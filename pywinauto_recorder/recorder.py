"""
Functions and classes that record a sequence of user actions.

All these functions and classes are private except the Recorder class.
"""

import sys
import os
import traceback
import time
import win32api
import win32con
import win32gui
import win32process
from threading import Event, Thread
import pywinauto
import overlay_arrows_and_more as oaam
import keyboard
import mouse
from collections import namedtuple
from enum import Enum
import pyperclip
import codecs
from .core import path_separator, type_separator, Strategy, is_int, \
	get_wrapper_path, get_entry_list, get_entry, get_sorted_rectangles, \
	read_config_file, SPECIAL_CHARACTERS, ABSOLUTE_PATH_SUFFIX, COMError
from .core import find_elements as not_ttl_cached_find_elements
from .player import playback
from cachetools import func


@func.ttl_cache(ttl=10)
def find_elements(full_element_path=None, window_handle=None):
	"""
	Returns the elements that match a path, like core.find_elements. The result is kept 10 seconds.

	:param full_element_path: The full path of the elements
	:param window_handle: The handle of the window where the elements are searched (see core.find_elements)
	"""
	return not_ttl_cached_find_elements(full_element_path=full_element_path, window_handle=window_handle)


__all__ = ['Recorder']

ElementEvent = namedtuple('ElementEvent', ['strategy', 'rectangle', 'path'])
SendKeysEvent = namedtuple('SendKeysEvent', ['line'])
MouseWheelEvent = namedtuple('MouseWheelEvent', ['delta'])
DragAndDropEvent = namedtuple('DragAndDropEvent', ['path', 'dx1', 'dy1', 'path2', 'dx2', 'dy2'])
ClickEvent = namedtuple('ClickEvent', ['button', 'click_count', 'path', 'dx', 'dy', 'time'])
FindEvent = namedtuple('FindEvent', ['path', 'dx', 'dy', 'time'])
MenuEvent = namedtuple('MenuEvent', ['path', 'menu_path'])


class IconSet:
	"""It loads the icons from the Icons folder of the package and stores them in the class."""

	path_icons = os.path.dirname(os.path.realpath(__file__))
	hicon_clipboard = oaam.load_ico(path_icons + r'\Icons\paste.ico', 48, 48)
	hicon_light_on = oaam.load_ico(path_icons + r'\Icons\light-on.ico', 48, 48)
	hicon_record = oaam.load_ico(path_icons + r'\Icons\record.ico', 48, 48)
	hicon_play = oaam.load_ico(path_icons + r'\Icons\play.ico', 48, 48)
	hicon_stop = oaam.load_ico(path_icons + r'\Icons\stop.ico', 48, 48)
	hicon_search = oaam.load_ico(path_icons + r'\Icons\search.ico', 48, 48)
	hicon_power = oaam.load_ico(path_icons + r'\Icons\power.ico', 48, 48)


def _escape_special_char(string):
	"""
	Escapes the special characters of a path, so that it can be written between the quotes of a string in a script.

	It is called on all paths to remove all special characters but it's not good.
	Should be moved in core.
	Should be called only in get_wrapper_path and unescape_special_char in get_entry,
	and in script += 'menu_click... in menu_click as it's already done.
	Should replace -> by _>.
	"""
	for character, escaped in SPECIAL_CHARACTERS:
		string = string.replace(character, escaped)
	return string


def _compute_dx_dy(x, y, rectangle):
	"""
	Returns the coordinates of a position relative to the center of a rectangle: -1 and 1 are its sides.

	A coordinate is 0 when the rectangle is too thin to compute it (2 pixels or less): before, the division by 0
	stopped the end of the recording, and the recording was lost.
	"""
	cx, cy = rectangle.mid_point()
	half_width, half_height = rectangle.width() / 2 - 1, rectangle.height() / 2 - 1
	dx = float(x - cx) / half_width if half_width > 0 else 0.0
	dy = float(y - cy) / half_height if half_height > 0 else 0.0
	return (dx, dy)


def _relative_coordinates(dx, dy, relative_coordinate_mode):
	"""
	Returns the relative coordinates written after the path of an element in the scripts: "%(dx,dy)".

	The coordinates are written in percent, with 2 decimals. Nothing is written when relative_coordinate_mode is False,
	or when one of the written coordinates is 0.

	:param dx: The horizontal coordinate, relative to the center of the element (-1 and 1 are its left and right sides)
	:param dy: The vertical coordinate, relative to the center of the element (-1 and 1 are its top and bottom sides)
	:param relative_coordinate_mode: True if the relative coordinates are recorded
	:return: "%(dx,dy)", or "" if no coordinates are written
	"""
	str_dx, str_dy = "{:.2f}".format(round(dx * 100, 2)), "{:.2f}".format(round(dy * 100, 2))
	if relative_coordinate_mode and float(str_dx) != 0 and float(str_dy) != 0:
		return '%(' + str_dx + ',' + str_dy + ')'
	return ''


_USER_EVENT_TYPES = (DragAndDropEvent, ClickEvent, FindEvent, MenuEvent)
"""The events of the user interface elements: the UIPath of the next lines of the script depends on their paths."""

_SCRIPT_EVENT_TYPES = (SendKeysEvent, MouseWheelEvent) + _USER_EVENT_TYPES
"""The events written in the scripts."""


def _write_in_file(events, relative_coordinate_mode=False):
	"""
	Writes the script that replays the events in a new file of the folder "Pywinauto recorder" of the home folder.

	The script is also copied to the clipboard.

	:param events: The processed events of the recording
	:param relative_coordinate_mode: True if the relative coordinates of the clicks are written
	:return: The path of the file
	"""
	from pathlib import Path
	home_dir = Path.home() / 'Pywinauto recorder'
	home_dir.mkdir(parents=True, exist_ok=True)
	record_file_name = home_dir / Path('recorded ' + time.asctime().replace(':', '_') + '.py')
	print('Recording in file: ' + str(record_file_name.absolute()))
	script = _get_script(events, relative_coordinate_mode)
	with codecs.open(record_file_name, "w", encoding=sys.getdefaultencoding()) as f:
		f.write(script)
	pyperclip.copy(script)
	return record_file_name


def _get_script(events, relative_coordinate_mode=False):
	"""
	Returns the Python script that replays the events.

	The paths of consecutive events in the same window, or in the same region, are written relatively to UIPath blocks.

	:param events: The processed events of the recording
	:param relative_coordinate_mode: True if the relative coordinates of the clicks are written
	:return: The Python script
	"""
	script = "# encoding: {}\n\n".format(sys.getdefaultencoding())
	script += "from pywinauto_recorder.player import *\n\n"
	common_path = ''
	common_window = ''
	common_region = ''
	for i, e_i in enumerate(events):
		if isinstance(e_i, _USER_EVENT_TYPES) and e_i.path != common_path:
			new_common_path = _find_new_common_path_in_next_user_events(events, i)
			if new_common_path != common_path:
				common_path = new_common_path
				uipath_lines, common_window, common_region = _get_uipath_lines(common_path, common_window, common_region)
				script += uipath_lines
		if isinstance(e_i, _SCRIPT_EVENT_TYPES):
			if common_window:
				script += '\t'
				if common_region:
					script += '\t'
			script += _get_script_line(e_i, common_path, relative_coordinate_mode)
	return script


def _get_uipath_lines(common_path, common_window, common_region):
	"""
	Returns the "with UIPath(...)" lines written when the common path of the next events changes.

	:param common_path: The new common path
	:param common_window: The window of the UIPath block of the previous events, or ''
	:param common_region: The region of the UIPath block of the previous events, or ''
	:return: The lines, the window and the region of the UIPath blocks of the next events
	"""
	lines = ''
	entry_list = get_entry_list(common_path)
	if entry_list[0] != common_window:
		common_window = entry_list[0]
		lines += '\nwith UIPath(u"' + _escape_special_char(common_window) + '"):\n'
	region = path_separator.join(entry_list[1:])
	if region != common_region and region:
		common_region = region
		lines += '\twith UIPath(u"' + _escape_special_char(common_region) + '"):\n'
	else:
		common_region = ''
	return lines, common_window, common_region


def _get_element_argument(path, dx, dy, common_path, relative_coordinate_mode):
	"""
	Returns the path of an element written in the script: relative to common_path, followed by the relative coordinates.

	A path that is not under common_path (see _is_under), like the destination of a drag and drop in another window, is
	written as an absolute path: the full path, ended by ABSOLUTE_PATH_SUFFIX, so that the path of the UIPath blocks is
	not added to it when the script is replayed. Before, the beginning of the path was cut wrongly.

	:param path: The full path of the element
	:param dx: The horizontal relative coordinate of the click (see _relative_coordinates)
	:param dy: The vertical relative coordinate of the click (see _relative_coordinates)
	:param common_path: The common path of the UIPath blocks, or ''
	:param relative_coordinate_mode: True if the relative coordinates are written
	:return: The text written between the quotes of the argument
	"""
	coordinates = _relative_coordinates(dx, dy, relative_coordinate_mode)
	if not common_path:
		return _escape_special_char(path) + coordinates
	if _is_under(path, common_path):
		return _escape_special_char(_get_relative_path(common_path, path)) + coordinates
	return _escape_special_char(_get_absolute_path(path)) + coordinates + ABSOLUTE_PATH_SUFFIX


def _get_script_line(event, common_path, relative_coordinate_mode):
	"""
	Returns the line of the script that replays an event.

	:param event: An event of a type of _SCRIPT_EVENT_TYPES
	:param common_path: The common path of the UIPath blocks, or ''
	:param relative_coordinate_mode: True if the relative coordinates are written
	:return: The line, without its indentation
	"""
	if isinstance(event, SendKeysEvent):
		return 'send_keys(' + event.line + ')\n'
	if isinstance(event, MouseWheelEvent):
		return 'mouse_wheel(' + str(event.delta) + ')\n'
	if isinstance(event, MenuEvent):
		return 'menu_click(u"' + _escape_special_char(event.menu_path) + '")\n'
	if isinstance(event, DragAndDropEvent):
		return 'drag_and_drop(u"' + \
			_get_element_argument(event.path, event.dx1, event.dy1, common_path, relative_coordinate_mode) + '", u"' + \
			_get_element_argument(event.path2, event.dx2, event.dy2, common_path, relative_coordinate_mode) + '")\n'
	argument = _get_element_argument(event.path, event.dx, event.dy, common_path, relative_coordinate_mode)
	if isinstance(event, ClickEvent):
		button = '' if event.button == 'left' else event.button + '_'
		return ['', '', 'double_', 'triple_'][event.click_count] + button + 'click(u"' + argument + '")\n'
	return 'wrapper = find(u"' + argument + '")\n'  # FindEvent


def _clean_events(events, remove_first_up=False):
	"""
	Removes duplicate or useless events.

	It removes all the last down events due or not to (CTRL+ALT+r) when ending record mode.

	:param remove_first_up: when True, removes the 2 first up events due to (CTRL+ALT+r) when starting record mode
	:param events: the copy of recorded event list
	"""
	if remove_first_up:
		i = 0
		i_up = 0
		while i < len(events):
			if isinstance(events[i], keyboard.KeyboardEvent) and events[i].event_type == 'up':
				i_up = i_up + 1
				events.pop(i)
				if i_up == 2:
					break
			else:
				i = i + 1
	i = 0
	previous_event_type = None
	while i < len(events):
		event_type = type(events[i])
		if event_type is previous_event_type and event_type in (ElementEvent, mouse.MoveEvent):
			del events[i - 1]  # only the last of consecutive element events, or of consecutive mouse moves, is kept
		else:
			previous_event_type = event_type
			i = i + 1
	i = len(events) - 1
	while i > 0:
		if isinstance(events[i], keyboard.KeyboardEvent) and events[i].event_type == 'down':
			events.pop(i)
		elif isinstance(events[i], keyboard.KeyboardEvent) and events[i].event_type == 'up':
			break
		i = i - 1


def _process_events(events, process_menu_click=True):
	"""
	Replaces the recorded events by the events written in the script.

	The keyboard events become SendKeysEvent, the mouse wheel events MouseWheelEvent, the presses and releases of the
	mouse buttons ClickEvent or DragAndDropEvent (from the last one: see _process_drag_and_drop_or_click_events), and
	the clicks on menu items MenuEvent if process_menu_click is True.

	:param events: The events of the recording, cleaned by _clean_events
	:param process_menu_click: True if the clicks on menu items become a MenuEvent
	"""
	send_keys_writer = _SendKeysWriter()
	i = 0
	while i < len(events):
		if isinstance(events[i], keyboard.KeyboardEvent):
			i = _process_keyboard_events(events, i, send_keys_writer)
			continue
		if isinstance(events[i], mouse.WheelEvent):
			_process_wheel_events(events, i)
		i = i + 1
	# the modifier keys still pressed at the end of the recording are released at the end of the script
	events.extend(SendKeysEvent(line=line) for line in send_keys_writer.write([], keep_modifiers=False))
	thresholds = _mouse_thresholds()
	i = len(events) - 1
	while i >= 0:
		if isinstance(events[i], mouse.ButtonEvent) and events[i].event_type == 'up':
			i = _process_drag_and_drop_or_click_events(events, i, thresholds)
		i = i - 1
	if process_menu_click:
		i = len(events) - 1
		while i >= 0:
			if isinstance(events[i], ClickEvent):
				i = _process_menu_select_events(events, i)
			i = i - 1


def _process_keyboard_events(events, i, send_keys_writer):
	"""
	Replaces the keyboard events that begin at index i by SendKeysEvent, one for each string of send_keys.

	The keyboard events are the ones before the next mouse button or wheel event: the element events and the mouse moves
	between them are kept after the SendKeysEvent. Before a mouse button or wheel event, the modifier keys that are
	pressed stay pressed in the script, for CTRL+click or SHIFT+click.

	:param events: The events of the recording
	:param i: The index of the first keyboard event
	:param send_keys_writer: The _SendKeysWriter of the recording
	:return: The index of the event after the SendKeysEvent
	"""
	i_end = i
	while i_end < len(events) and isinstance(events[i_end], (keyboard.KeyboardEvent, ElementEvent, mouse.MoveEvent)):
		i_end = i_end + 1
	keyboard_events = [event for event in events[i:i_end] if isinstance(event, keyboard.KeyboardEvent)]
	lines = send_keys_writer.write(keyboard_events, keep_modifiers=i_end < len(events))
	other_events = [event for event in events[i:i_end] if not isinstance(event, keyboard.KeyboardEvent)]
	events[i:i_end] = [SendKeysEvent(line=line) for line in lines] + other_events
	return i + len(lines)


def _process_wheel_events(events, i):
	delta = events[i].delta
	i_processed_events = []
	i0 = i + 1
	while i0 < len(events):
		if isinstance(events[i0], mouse.WheelEvent):
			delta = delta + events[i0].delta
			i_processed_events.append(i0)
			i0 = i0 + 1
		elif type(events[i0]) in (ElementEvent, mouse.MoveEvent):
			i0 = i0 + 1
		else:
			break
	for i_p_e in sorted(i_processed_events, reverse=True):
		del events[i_p_e]
	events[i] = MouseWheelEvent(delta=delta)


DEFAULT_DRAG_THRESHOLD = (4, 4)
"""
The distances, in pixels, that the mouse can move horizontally and vertically while a button is pressed, before it is
a drag and drop: the default values of Windows (SM_CXDRAG and SM_CYDRAG).
"""

DEFAULT_DOUBLE_CLICK_RECTANGLE = (4, 4)
"""
The size, in pixels, of the rectangle centered on a click where the next click must be to make a double click: the
default values of Windows (SM_CXDOUBLECLK and SM_CYDOUBLECLK).
"""

MouseThresholds = namedtuple('MouseThresholds', ['drag', 'double_click'])
"""
The distances, in pixels, that the mouse can move horizontally and vertically: while a button is pressed, before it is
a drag and drop (drag), and from a click to the next click of a double click (double_click).
"""

DEFAULT_MOUSE_THRESHOLDS = MouseThresholds(
	drag=DEFAULT_DRAG_THRESHOLD, double_click=(DEFAULT_DOUBLE_CLICK_RECTANGLE[0] / 2, DEFAULT_DOUBLE_CLICK_RECTANGLE[1] / 2))
"""The mouse thresholds of Windows with its default settings."""


def _system_distances(metric_x, metric_y, default):
	"""Returns 2 distances given by Windows (GetSystemMetrics), or the default distances if Windows does not give them."""
	distances = (win32api.GetSystemMetrics(metric_x), win32api.GetSystemMetrics(metric_y))
	if all(isinstance(distance, int) and distance > 0 for distance in distances):
		return distances
	return default


def _mouse_thresholds():
	"""
	Returns the MouseThresholds of Windows.

	The applications use them to tell a click from a drag and drop, and a double click from 2 clicks: a click is recorded
	as a click when the hand moves the mouse a little, and 2 quick clicks on 2 elements are recorded as 2 clicks.
	"""
	width, height = _system_distances(win32con.SM_CXDOUBLECLK, win32con.SM_CYDOUBLECLK, DEFAULT_DOUBLE_CLICK_RECTANGLE)
	return MouseThresholds(
		drag=_system_distances(win32con.SM_CXDRAG, win32con.SM_CYDRAG, DEFAULT_DRAG_THRESHOLD),
		double_click=(width / 2, height / 2))


def _is_far(position, origin, distances):
	"""Returns True if a position is farther from an origin than the distances, horizontally or vertically."""
	return abs(position[0] - origin[0]) > distances[0] or abs(position[1] - origin[1]) > distances[1]


def _contains(rectangle, position):
	"""Returns True if the rectangle of an element contains a position. Its right and bottom sides are outside it."""
	x, y = position
	return rectangle is not None and rectangle.left <= x < rectangle.right and rectangle.top <= y < rectangle.bottom


def _is_descendant(path, ancestor_path):
	"""
	Returns True if a path is the path of a descendant of the element of ancestor_path.

	The [] operator at the end of ancestor_path is ignored: "W||Window->L||List#[0,1]" is an ancestor of
	"W||Window->L||List->A||ListItem".
	"""
	i = ancestor_path.rfind("#[")
	if i != -1:
		ancestor_path = ancestor_path[:i]
	return path.startswith(ancestor_path + path_separator)


def _find_previous_event(events, i, event_type):
	"""
	Searches the last event of type event_type in events[:i + 1].

	:return: The index of the event and the event, or -1 and None if there is none
	"""
	while i >= 0:
		if isinstance(events[i], event_type):
			return i, events[i]
		i = i - 1
	return -1, None


def _position_at(events, i):
	"""Returns the position of the mouse cursor at events[i]: the one of the last mouse move before it, or None."""
	_, move_event = _find_previous_event(events, i - 1, mouse.MoveEvent)
	if move_event is None:
		return None
	return move_event.x, move_event.y


def _find_previous_button_event(events, i, button):
	"""Returns the index of the last event of a mouse button before events[i], or -1 if there is none."""
	for i0 in range(i - 1, -1, -1):
		if isinstance(events[i0], mouse.ButtonEvent) and events[i0].button == button:
			return i0
	return -1


def _find_press(events, i):
	"""
	Returns the index of the press of the mouse button released by events[i], or -1 if there is none.

	There is none when the previous event of the button is a release too: the button was pressed before the start of the
	recording, or on Pywinauto Recorder (see Recorder.__is_ignored).
	"""
	i_press = _find_previous_button_event(events, i, events[i].button)
	if i_press != -1 and events[i_press].event_type in ('down', 'double'):
		return i_press
	return -1


def _is_drag(events, i_press, i, threshold):
	"""Returns True if the mouse moved farther than the drag threshold from the press events[i_press] to events[i]."""
	position = _position_at(events, i_press)
	return any(
		isinstance(event, mouse.MoveEvent) and _is_far((event.x, event.y), position, threshold)
		for event in events[i_press + 1:i])


def _find_previous_click(events, i_press, thresholds):
	"""
	Returns the index of the press of the click before the press events[i_press], if they make a double click, or -1.

	They make a double click if the press events[i_press] is quick: the 'mouse' library gives it as a double click
	('double') when it is soon after the previous event of a mouse button, which must be the release of the same button.
	The previous click must not be a drag and drop, and the press must be near its press (see MouseThresholds): 2 quick
	clicks on 2 elements are 2 clicks, as for Windows.
	"""
	if events[i_press].event_type != 'double':
		return -1
	i_release, release = _find_previous_event(events, i_press - 1, mouse.ButtonEvent)
	if release is None or release.event_type != 'up' or release.button != events[i_press].button:
		return -1
	i_previous = _find_press(events, i_release)
	if i_previous == -1 or _position_at(events, i_previous) is None:
		return -1
	if _is_drag(events, i_previous, i_release, thresholds.drag):
		return -1
	if _is_far(_position_at(events, i_press), _position_at(events, i_previous), thresholds.double_click):
		return -1
	return i_previous


def _find_first_click(events, i_press, thresholds):
	"""
	Returns the index of the first press of the double or triple click that ends with the press events[i_press].

	:return: The index of the first press, and the number of clicks: 1, 2 or 3
	"""
	click_count = 1
	while click_count < 3:
		i_previous = _find_previous_click(events, i_press, thresholds)
		if i_previous == -1:
			break
		i_press, click_count = i_previous, click_count + 1
	return i_press, click_count


def _element_events_before(events, i, first):
	"""
	Returns the element events found just before events[i], from the last one.

	The search stops at an event that is neither an element event nor a mouse move, except the mouse button events from
	events[first]: they are the ones of the drag and drop of the release events[i].
	"""
	element_events = []
	for i0 in range(i - 1, -1, -1):
		event = events[i0]
		if isinstance(event, ElementEvent):
			element_events.append(event)
		elif not isinstance(event, mouse.MoveEvent) and not (i0 >= first and isinstance(event, mouse.ButtonEvent)):
			break
	return element_events


def _element_events_after(events, i, last, position, threshold):
	"""
	Returns the element events found just after events[i], while the mouse cursor stays near a position.

	The search stops at a mouse move farther than the drag threshold, and at an event that is neither an element event
	nor a mouse move, except the mouse button events up to events[last]: they are the ones of the click of events[i].
	"""
	element_events = []
	for i0 in range(i + 1, len(events)):
		event = events[i0]
		if isinstance(event, ElementEvent):
			element_events.append(event)
		elif isinstance(event, mouse.MoveEvent):
			if _is_far((event.x, event.y), position, threshold):
				break
		elif not (i0 <= last and isinstance(event, mouse.ButtonEvent)):
			break
	return element_events


def _find_element_event_at(element_events_before, element_events_after, position):
	"""
	Returns the element event of the element under the mouse cursor at a position, when a mouse button is pressed.

	The element events are added by the loop of the recorder, which needs some time to find the element under the mouse
	cursor: when the mouse moves quickly, the last one before the press may be the one of another element, or of an
	ancestor of the element, like the pane that contains the clicked button.
	So the element is the last one found before the press whose rectangle contains the position, or else the first one
	found after it. Then an element found after it at this position, inside this element, replaces it: the mouse cursor
	was already on it.

	:param element_events_before: The element events found just before, from the last one (see _element_events_before)
	:param element_events_after: The element events found just after, from the first one (see _element_events_after)
	:param position: The position of the mouse cursor
	:return: The element event, or None if none contains the position
	"""
	element_event = next((e for e in element_events_before if _contains(e.rectangle, position)), None)
	for e in element_events_after:
		if _contains(e.rectangle, position) and (element_event is None or _is_descendant(e.path, element_event.path)):
			element_event = e
	return element_event


def _press_element_event(events, i_press, i, threshold):
	"""
	Returns the element event of the element under the mouse cursor at the press of a click or of a drag and drop.

	The press is events[i_press], and the release events[i] (see _find_element_event_at). When no element is found at
	the position of the press, it is the last element found before the press, as before.

	:return: The element event, or None if no element was found
	"""
	position = _position_at(events, i_press)
	return _find_element_event_at(
		_element_events_before(events, i_press, i_press),
		_element_events_after(events, i_press, i, position, threshold),
		position) or _find_previous_event(events, i_press - 1, ElementEvent)[1]


def _release_element_event(events, i_press, i):
	"""
	Returns the element event of the element under the mouse cursor at the release events[i] of a drag and drop.

	It is the last element found before the release whose rectangle contains the position of the release, or else the
	last element found before the release. The elements found after the release are not used: the drop often changes the
	elements under the mouse cursor, for example the dragged element may be there.

	:param i_press: The index of the press of the drag and drop
	:return: The element event, or None if no element was found
	"""
	return _find_element_event_at(_element_events_before(events, i, i_press), [], _position_at(events, i)) or \
		_find_previous_event(events, i - 1, ElementEvent)[1]


def _click_event(events, i_press, i, click_count, threshold):
	"""
	Returns the ClickEvent of a click from the press events[i_press] to the release events[i].

	The click is at the position of the press. It returns None if no element was found under the mouse cursor.
	"""
	element_event = _press_element_event(events, i_press, i, threshold)
	if element_event is None:
		return None
	x, y = _position_at(events, i_press)
	dx, dy = _compute_dx_dy(x, y, element_event.rectangle)
	release = events[i]
	return ClickEvent(
		button=release.button, click_count=click_count, path=element_event.path, dx=dx, dy=dy, time=release.time)


def _drag_and_drop_event(events, i_press, i, threshold):
	"""
	Returns the DragAndDropEvent of a drag and drop from the press events[i_press] to the release events[i].

	It returns None if no element was found under the mouse cursor.
	"""
	start, end = _press_element_event(events, i_press, i, threshold), _release_element_event(events, i_press, i)
	if start is None or end is None:
		return None
	x1, y1 = _position_at(events, i_press)
	x2, y2 = _position_at(events, i)
	dx1, dy1 = _compute_dx_dy(x1, y1, start.rectangle)
	dx2, dy2 = _compute_dx_dy(x2, y2, end.rectangle)
	return DragAndDropEvent(path=start.path, dx1=dx1, dy1=dy1, path2=end.path, dx2=dx2, dy2=dy2)


def _remove_mouse_and_element_events(events, i1, i):
	"""
	Removes the mouse button, mouse move and element events from events[i1] to events[i - 1].

	:return: The new index of the event of index i
	"""
	removed = [i0 for i0 in range(max(i1, 0), i) if isinstance(events[i0], (mouse.ButtonEvent, mouse.MoveEvent, ElementEvent))]
	for i0 in reversed(removed):
		del events[i0]
	return i - len(removed)


def _process_drag_and_drop_or_click_events(events, i, thresholds=DEFAULT_MOUSE_THRESHOLDS):
	"""
	Replaces the release of a mouse button, events[i], by a ClickEvent or a DragAndDropEvent.

	It is a drag and drop if the mouse moved farther than the drag threshold while the button was pressed. Otherwise, it
	is a click at the position of the press: a double or triple click if the previous clicks were quick and near (see
	_find_previous_click). The mouse button, mouse move and element events from the press to the release are removed.
	A release without press is removed (see _find_press). A click or a drag and drop is removed, and a message is
	printed, when no element was found under the mouse cursor: before, the recording was lost.

	:param events: The events of the recording
	:param i: The index of a release of a mouse button
	:param thresholds: The MouseThresholds
	:return: The new index of the ClickEvent or of the DragAndDropEvent, or the index of the first removed event
	"""
	i_press = _find_press(events, i)
	if i_press == -1:
		del events[i]
		return i
	if _position_at(events, i_press) is None:
		event = None  # no mouse move before the press, which is impossible after start_recording
	elif _is_drag(events, i_press, i, thresholds.drag):
		event = _drag_and_drop_event(events, i_press, i, thresholds.drag)
	else:
		i_press, click_count = _find_first_click(events, i_press, thresholds)
		event = _click_event(events, i_press, i, click_count, thresholds.drag)
	i = _remove_mouse_and_element_events(events, i_press, i)
	if event is None:
		print("A mouse click is not recorded: no element was found under the mouse cursor.")
		del events[i]
	else:
		events[i] = event
	return i


def _split_reference_path(path):
	"""
	Splits the path of an element found with a reference element (Strategy.array_1d): "path#[reference path,x]".

	:param path: The full path of an element
	:return: The path without the [] operator, the path of the reference element and the column x. The path, None and
		None if the element was not found with a reference element.
	"""
	entry_list = get_entry_list(path)
	str_name, str_type, y_x, _ = get_entry(entry_list[-1])
	if y_x is None or is_int(y_x[0]):
		return path, None, None
	return path_separator.join(entry_list[:-1] + [str_name + type_separator + str_type]), y_x[0], y_x[1]


def _is_under(path, common_path):
	"""
	Returns True if a path can be written relatively to common_path.

	It starts with common_path, like the path of its reference element (Strategy.array_1d) if there is one.
	"""
	element_path, reference_path, _ = _split_reference_path(path)
	return all(
		p == common_path or p.startswith(common_path + path_separator)
		for p in (element_path, reference_path) if p is not None)


def _get_relative_path(common_path, path):
	"""
	Returns the path of an element relative to common_path, the path of one of its ancestors (see _is_under).

	The path of the reference element of the [] operator (Strategy.array_1d) is made relative to common_path too.

	:param common_path: The path of the UIPath blocks
	:param path: The full path of the element
	:return: The relative path: '' for the element of common_path
	"""
	if not path:
		return ''
	element_path, reference_path, x = _split_reference_path(path)
	relative_path = element_path[len(common_path) + len(path_separator):]
	if reference_path is not None:
		relative_path += "#[" + reference_path[len(common_path) + len(path_separator):] + "," + str(x) + "]"
	return relative_path


def _get_absolute_path(path):
	"""
	Returns the absolute path of an element, without ABSOLUTE_PATH_SUFFIX at the end (see _get_element_argument).

	It is the full path, where the path of the reference element of the [] operator (Strategy.array_1d), if there is
	one, is absolute too: "path#[reference path~Absolute_UIPath,x]".
	"""
	element_path, reference_path, x = _split_reference_path(path)
	if reference_path is None:
		return path
	return element_path + "#[" + reference_path + ABSOLUTE_PATH_SUFFIX + "," + str(x) + "]"


def _find_common_path(current_path, next_path):
	current_entry_list = get_entry_list(current_path)
	if len(current_entry_list) > 1:
		_, _, y_x, _ = get_entry(current_entry_list[-1])
		if (y_x is not None) and not is_int(y_x[0]):
			current_entry_list = get_entry_list(y_x[0])[:-1]
		else:
			current_entry_list = current_entry_list[:-1]
	next_entry_list = get_entry_list(next_path)
	if len(next_entry_list)>1:
		next_entry_list = next_entry_list[:-1]
	n = 0
	try:
		while current_entry_list[n] == next_entry_list[n]:
			n = n + 1
	except IndexError:
		common_path = path_separator.join(current_entry_list[0:n])
		return common_path
	common_path = path_separator.join(current_entry_list[0:n])
	return common_path


def _find_new_common_path_in_next_user_events(events, i):
	path_i = events[i].path
	i0 = i + 1
	new_common_path = ''
	while i0 < len(events):
		e = events[i0]
		if type(e) in (DragAndDropEvent, ClickEvent, FindEvent, MenuEvent):
			new_common_path = _find_common_path(path_i, e.path)
			break
		elif type(e) in (ElementEvent, mouse.MoveEvent):
			i0 = i0 + 1
		else:
			break
	if new_common_path == '':
		new_common_path = _find_common_path(path_i, path_i)
	return new_common_path


def _process_menu_select_events(events, i):
	i0 = i
	i_processed_events = []
	menu_path = []
	while i0 >= 0:
		if isinstance(events[i0], ClickEvent):
			entry_list = get_entry_list(events[i0].path)
			matching = [s for s in entry_list if "||MenuItem" in s]
			if matching:
				str_name, _, _, _ = get_entry(matching[0])
				menu_path.append(str_name)
				i_processed_events.append(i0)
				if [s for s in entry_list if "||MenuBar" in s]:
					break
			else:
				break
		i0 -= 1
	if menu_path:
		menu_path = path_separator.join(reversed(menu_path))
		i_menu_bar = i_processed_events.pop()
		menu_bar_path = get_entry_list(events[i_menu_bar].path)[0]
		events[i_menu_bar] = MenuEvent(path=menu_bar_path, menu_path=menu_path)
		for i_p_e in sorted(i_processed_events, reverse=True):
			del events[i_p_e]
			i -= 1
	return i


_KEYS_WITHOUT_CHARACTER = """
1 ESC esc
14 BACKSPACE backspace
15 TAB tab
28 ENTER enter
55 PRTSC print screen
59 F1 f1
60 F2 f2
61 F3 f3
62 F4 f4
63 F5 f5
64 F6 f6
65 F7 f7
66 F8 f8
67 F9 f9
68 F10 f10
69 VK_PAUSE pause
71 HOME home
72 UP up
73 PGUP page up
75 LEFT left
77 RIGHT right
79 END end
80 DOWN down
81 PGDN page down
82 INSERT insert
83 DELETE delete
87 F11 f11
88 F12 f12
93 VK_APPS menu
- HELP help
- VK_DECIMAL decimal
- VK_SEPARATOR separator
- VK_CLEAR clear
""".strip().splitlines() + ['- F{0} f{0}'.format(n) for n in range(13, 25)]
"""
The keys that type no character: their scan code ('-' if they are found by their name only), their code of send_keys,
and the name that the 'keyboard' library gives them.

Windows gives the names of some of these keys in its language, like "suppr" for the Delete key in French: they are
found by their scan code. The keys of the numeric keypad, which have the same scan codes, have English names.
"""

_SEND_KEYS_CODES = {key.split(maxsplit=2)[2]: key.split()[1] for key in _KEYS_WITHOUT_CHARACTER}
"""The codes of send_keys of the keys that type no character, by the names that the 'keyboard' library gives them."""

_SEND_KEYS_SCAN_CODES = {int(key.split()[0]): key.split()[1] for key in _KEYS_WITHOUT_CHARACTER if key[0] != '-'}
"""The codes of send_keys of the keys that type no character, by scan code (see _KEYS_WITHOUT_CHARACTER)."""

_IGNORED_KEYS = ('caps lock', 'num lock', 'scroll lock')
"""
The lock keys, which are not written: the 'keyboard' library gives the characters typed with caps lock, and send_keys
turns num lock off.
"""

_IGNORED_SCAN_CODES = (58, 70)
"""The scan codes of caps lock and scroll lock, whose names depend on the language of Windows (see _IGNORED_KEYS)."""

_MODIFIER_SCAN_CODES = {42: 'shift', 54: 'shift', 29: 'ctrl', 541: 'alt gr', 56: 'alt', 91: 'windows', 92: 'windows'}
"""
The modifier keys, by scan code: Windows gives their names in its language, like "maj" for shift in French.

541 is the scan code of the left ctrl that Windows sends when alt gr is pressed: the 'keyboard' library ignores the
right alt that follows.
"""

_MODIFIER_CODES = {'ctrl': 'VK_CONTROL', 'alt': 'VK_MENU', 'shift': 'VK_SHIFT', 'windows': 'VK_LWIN'}
"""
The codes of send_keys of the modifier keys written in the scripts, in the order in which they are pressed.

Alt gr is not written: the characters typed with it are.
"""

_ALONE_MODIFIER_CODES = {'alt': 'VK_MENU', 'windows': 'LWIN'}
"""
The codes of send_keys of the modifier keys written when they are pressed alone: alt selects the menu bar, and windows
opens the Start menu.
"""

_SEND_KEYS_SPECIAL_CHARACTERS = {'{': '{{}', '}': '{}}', '^': '{^}', '%': '{%}', '~': '{~}'}
"""
The characters that send_keys reads as keys, and how they are escaped to be typed.

'+', '(' and ')' are not escaped in the scripts: player.send_keys escapes them.
"""


def _modifier(event):
	"""Returns the modifier key of a keyboard event: 'ctrl', 'alt', 'shift', 'windows', 'alt gr', or None."""
	if event.scan_code in _MODIFIER_SCAN_CODES:
		return _MODIFIER_SCAN_CODES[event.scan_code]
	name = event.name or ''
	if name.startswith('maj'):
		return 'shift'
	for modifier in ('alt gr', 'ctrl', 'alt', 'shift', 'windows'):
		if modifier in name:
			return modifier
	return None


def _typed_key(event):
	"""
	Returns what a key that is not a modifier key types.

	:param event: A keyboard event
	:return: (the character, True), (the code of send_keys of a key that types no character, like '{TAB}', False), or
		(None, False) for a key that is not written
	"""
	name = event.name or ''
	if name == 'space':
		return ' ', True
	if len(name) == 1:
		return (name, True) if ord(name) >= 32 else (None, False)
	if name in _IGNORED_KEYS or (not event.is_keypad and event.scan_code in _IGNORED_SCAN_CODES):
		return None, False
	code = _SEND_KEYS_CODES.get(name)
	if code is None and not event.is_keypad:
		code = _SEND_KEYS_SCAN_CODES.get(event.scan_code)
	return ('{' + code + '}', False) if code else (None, False)


def _is_recorder_hotkey(character, modifiers):
	"""Returns True if a character typed with modifier keys is a keyboard shortcut of the recorder, like CTRL+ALT+R."""
	for hotkey in (HOTKEY_RECORD, HOTKEY_SMART_MODE, HOTKEY_ELEMENT_INFO, HOTKEY_COPY_CODE):
		keys = hotkey.lower().split('+')
		if character.lower() == keys[-1] and set(keys[:-1]) <= modifiers:
			return True
	return False


class _SendKeysWriter:
	"""
	Writes the keyboard events of a recording in the strings of send_keys, in the syntax of pywinauto.

	The characters are written as they are typed: the 'keyboard' library gives the characters typed with shift, alt gr
	and caps lock, so these keys are not written, except shift with a key that types no character, like SHIFT+TAB. The
	keys that type no character are written between braces, like {TAB} or {ENTER}, and the keyboard shortcuts with the
	presses and the releases of their modifier keys, like {VK_CONTROL down}s{VK_CONTROL up}.

	In a string, send_keys sends the characters that follow the press or the release of a key as virtual keys, without
	shift: a character that is typed after a press or a release is written in a new string.

	Before, a text that was not followed by ENTER was lost, like the keys that type no character, and the keys that type
	no character were written with names that send_keys does not know, like {f4}.
	"""

	def __init__(self):
		self.held = set()
		"""The modifier keys that are pressed: 'ctrl', 'alt', 'shift', 'windows' or 'alt gr'."""
		self.combined = set()
		"""The modifier keys that are pressed with another key: they are not written as pressed alone."""
		self.written = []
		"""The modifier keys written as pressed and not yet released, in the order of _MODIFIER_CODES."""
		self.strings = []
		"""The strings of send_keys of a group of keyboard events: [keys, True if a key is pressed or released in it]."""

	def write(self, keyboard_events, keep_modifiers):
		"""
		Returns the strings of send_keys of a group of keyboard events, between quotes and escaped like in the scripts.

		:param keyboard_events: The keyboard events of the group
		:param keep_modifiers: True if the modifier keys pressed at the end of the group stay pressed, because a mouse
			button or wheel event follows (CTRL+click): else, all the modifier keys are released
		"""
		self.strings = [[[], False]]
		for event in keyboard_events:
			modifier = _modifier(event)
			if modifier:
				self._modifier_event(modifier, event.event_type == 'down')
			elif event.event_type == 'down':
				self._key_down(event)
		if keep_modifiers:
			self.combined.update(self.held)
			self._write_modifiers(self.held)
		else:
			self._write_modifiers(())
		return ['"' + _escape_special_char(''.join(key for key, _ in keys)) + '"' for keys, _ in self.strings if keys]

	def _add(self, key, is_character=False, is_press_or_release=False):
		"""Adds a key, written in the syntax of send_keys, at the end of the current string."""
		self.strings[-1][0].append((key, is_character))
		if is_press_or_release:
			self.strings[-1][1] = True

	def _write_modifiers(self, modifiers, again=False):
		"""
		Writes the releases and the presses that make the given modifier keys the only ones written as pressed.

		:param modifiers: The modifier keys that must be pressed
		:param again: True if the modifier keys already pressed are pressed again when the current string has no press
			or release of a key, so that the characters that follow are sent as virtual keys (keyboard shortcut)
		"""
		modifiers = [modifier for modifier in _MODIFIER_CODES if modifier in modifiers]
		again = again and not self.strings[-1][1]
		for modifier in reversed(self.written):
			if modifier not in modifiers:
				self._add('{' + _MODIFIER_CODES[modifier] + ' up}', is_press_or_release=True)
		for modifier in modifiers:
			if again or modifier not in self.written:
				self._add('{' + _MODIFIER_CODES[modifier] + ' down}', is_press_or_release=True)
		self.written = modifiers

	def _modifier_event(self, modifier, is_down):
		"""Processes the press or the release of a modifier key."""
		if is_down:
			if modifier not in self.held and self.held:
				self.combined.update(self.held | {modifier})
			self.held.add(modifier)
			return
		alone = modifier in self.held and modifier not in self.combined
		self.held.discard(modifier)
		self.combined.discard(modifier)
		if modifier in self.written:
			self._write_modifiers([written for written in self.written if written != modifier])
		elif alone and modifier in _ALONE_MODIFIER_CODES:
			self._add('{' + _ALONE_MODIFIER_CODES[modifier] + '}')

	def _key_down(self, event):
		"""Processes the press of a key that is not a modifier key."""
		self.combined.update(self.held)
		key, is_character = _typed_key(event)
		if key is None or (is_character and _is_recorder_hotkey(key, self.held)):
			return
		modifiers = [modifier for modifier in ('ctrl', 'alt', 'windows') if modifier in self.held]
		if 'shift' in self.held and (modifiers or not is_character or key == ' '):
			modifiers.append('shift')
		if modifiers:
			self._write_modifiers(modifiers, again=is_character)
			if key == ' ':
				self._add('{SPACE}')
			else:
				self._add(_SEND_KEYS_SPECIAL_CHARACTERS.get(key, key) if is_character else key)
			return
		self._write_modifiers(())
		self._type(key, is_character)

	def _type(self, key, is_character):
		"""
		Writes a key typed without modifier key.

		A backspace removes the character typed before it in the string, if there is one: before, the backspaces were
		always removed.
		"""
		keys, has_press_or_release = self.strings[-1]
		if key == '{BACKSPACE}' and keys and keys[-1][1] and not has_press_or_release:
			keys.pop()
			return
		if is_character and has_press_or_release:
			self.strings.append([[], False])
		self._add(_SEND_KEYS_SPECIAL_CHARACTERS.get(key, key) if is_character else key, is_character=is_character)


t0_progress_icon_timings = time.time()
progress_icon_timings = [0, 0, 0, 0, 0, 0, 0]


def _overlay_add_icon_frame(main_overlay, x, y):
	"""Draws the frame of an icon of the recorder, 52x52 pixels, at (x, y)."""
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=x, y=y, width=52, height=52,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
	main_overlay.add(
		geometry=oaam.Shape.triangle,
		xyrgb_array=((x + 1, y + 1, 255, 255, 254), (x + 1, y + 52, 128, 128, 128), (x + 51, y + 52, 255, 255, 254)),
		thickness=0)


def _overlay_add_progress_icon(main_overlay, i, x, y):
	global t0_progress_icon_timings
	_overlay_add_icon_frame(main_overlay, x, y)
	dt = time.time() - t0_progress_icon_timings
	nb_dt = int(dt/0.01)
	if nb_dt > 255:
		nb_dt = int(255)
	progress_icon_timings[i % 6 -1] = nb_dt
	
	for b in range(i % 6):
		c = progress_icon_timings[b]
		main_overlay.add(
			geometry=oaam.Shape.rectangle, x=x + 6, y=y + 6 + b * 8, width=40, height=6,
			# color=(0, 255, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 200, 0))
			color=(c, int(255 - c/2), c ), thickness=1, brush=oaam.Brush.solid, brush_color=(c, int(255 - c), 0))
	t0_progress_icon_timings = time.time()


def _overlay_add_mode_icon(main_overlay, hicon, x, y):
	_overlay_add_icon_frame(main_overlay, x, y)
	main_overlay.add(
		geometry=oaam.Shape.image, hicon=hicon, x=int(x+2), y=int(y+2))


HOTKEY_RECORD = "CTRL+ALT+R"
"""Keyboard shortcut that starts or stops the recording (handled by Recorder.__key_on)."""
HOTKEY_SMART_MODE = "CTRL+ALT+S"
"""Keyboard shortcut that starts or stops the Smart mode (handled by Recorder.__key_on)."""
HOTKEY_ELEMENT_INFO = "CTRL+SHIFT+D"
"""Keyboard shortcut that starts or stops displaying element information (handled by Recorder.__key_on)."""
HOTKEY_COPY_CODE = "CTRL+SHIFT+F"
"""Keyboard shortcut that copies the Python code finding the element to the clipboard (handled by Recorder.__key_on)."""

QUIT_TIMEOUT = 10
"""Maximum duration, in seconds, of the wait for the end of the loop of the recorder when it quits."""

KEYBOARD_HOOKS_CHECK_TIMEOUT = 1.0
"""Maximum duration, in seconds, of the check of the keyboard hooks when the recorder starts."""

_VK_LSHIFT = 0xA0


def keyboard_hooks_work(timeout=KEYBOARD_HOOKS_CHECK_TIMEOUT, interval=0.2):
	"""
	Checks that the keyboard hooks receive the keyboard events.

	It sends a release of the left Shift key until a keyboard event is received, or until the timeout. The release is
	sent again because the hook is installed by a thread of the 'keyboard' library: the first release may be sent
	before the hook is installed.

	:param timeout: The maximum duration of the check, in seconds
	:param interval: The duration, in seconds, of the wait for a keyboard event after each release of the Shift key
	:return: True if a keyboard event was received
	"""
	received = Event()
	hooked = keyboard.hook(lambda event: received.set())
	try:
		deadline = time.monotonic() + timeout
		while True:
			win32api.keybd_event(_VK_LSHIFT, 0, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)
			if received.wait(interval):
				return True
			if time.monotonic() >= deadline:
				return False
	finally:
		keyboard.unhook(hooked)


def _is_hotkey_released(event, key, modifiers):
	"""
	Returns True if a keyboard event is the release of a key while modifier keys are pressed.

	:param event: The keyboard event
	:param key: The name of the key, for example 'r'
	:param modifiers: The names of the modifier keys, for example ("alt", "ctrl")
	"""
	if (event.name, event.event_type) != (key, 'up'):
		return False
	return {keyboard.key_to_scan_codes(modifier)[0] for modifier in modifiers}.issubset(keyboard._pressed_events)


def _top_level_window(wrapper):
	"""
	Returns the top level window of an element: the element itself, or its ancestor whose parent is the desktop.

	The parent of each ancestor is asked once: it is a request to the application. Before, it was asked 4 times at each
	loop of the recorder.
	"""
	window, parent = wrapper, wrapper.parent()
	while parent is not None:
		grandparent = parent.parent()
		if grandparent is None:
			break
		window, parent = parent, grandparent
	return window


def _is_cursor_in(rectangle):
	"""Returns True if the mouse cursor is inside the rectangle, and not on its border."""
	try:
		x, y = win32api.GetCursorPos()
	except win32api.error:  # the position is not available (see Recorder.__cursor_position)
		return False
	return rectangle.left < x < rectangle.right and rectangle.top < y < rectangle.bottom


_UIA_E_ELEMENTNOTAVAILABLE = 0x80040201
"""
The error of UI Automation for an element that does not exist anymore. Its message is misleading: "An event was unable to
invoke any of the subscribers".
"""


def _is_vanished_element_error(error):
	"""Returns True if an error is the one of UI Automation for an element that vanished: the application changed."""
	return isinstance(error, COMError) and bool(error.args) and isinstance(error.args[0], int) and \
		error.args[0] & 0xFFFFFFFF == _UIA_E_ELEMENTNOTAVAILABLE


def _no_tray_icon(x, y):
	"""Returns False: the position is not on a tray icon. It is Recorder.is_on_tray_icon when there is no tray icon."""
	return False


_MENU_WINDOW_CLASS = "#32768"
"""The class of the windows of the menus, like the menu of the tray icon of Pywinauto Recorder."""


def _is_menu_of_this_process(x, y):
	"""Returns True if the window at a position of the screen is a menu of this process, like the menu of the tray icon."""
	try:
		window = win32gui.WindowFromPoint((x, y))
		return win32gui.GetClassName(window) == _MENU_WINDOW_CLASS and \
			win32process.GetWindowThreadProcessId(window)[1] == os.getpid()
	except Exception:  # pywintypes.error, for example if there is no window at the position
		return False


class _ElementSearch:
	"""
	The state of the search of a unique path of the element under the mouse cursor, kept from a loop to the next one.

	The strategies are tried one after the other, a strategy at each loop, while the mouse cursor stays on the same
	element and no unique path is found. The strategy Strategy.array_1d is used only in Smart mode.
	"""

	STRATEGIES = (Strategy.unique_path, Strategy.array_2d, Strategy.array_1d)

	def __init__(self):
		"""Creates the state of a search that starts with the first strategy."""
		self.elements = []
		self.previous_wrapper_path = None
		self.previous_window_handle = None
		self.unique_wrapper_path = None
		self.i_strategy = 0
		self.searched_again_rectangle = None

	def next_strategy(self, wrapper_path, smart_mode, window_handle=None):
		"""
		Returns the strategy of the search of a unique path of the element of path wrapper_path.

		The elements of path wrapper_path are searched again, in the window of window_handle, when the element under the
		mouse cursor changes.

		:param wrapper_path: The path of the element under the mouse cursor
		:param smart_mode: True if the Smart mode is active
		:param window_handle: The handle of the window under the mouse cursor (see core.find_elements)
		:return: The strategy
		"""
		if (wrapper_path, window_handle) == (self.previous_wrapper_path, self.previous_window_handle):
			if (self.unique_wrapper_path is None) or (self.STRATEGIES[self.i_strategy] == Strategy.array_2d):
				self.i_strategy = min(self.i_strategy + 1, len(self.STRATEGIES) - 1)
				if (not smart_mode) and (self.STRATEGIES[self.i_strategy] == Strategy.array_1d):
					self.i_strategy = 1
		else:
			self.i_strategy = 0
			self.previous_wrapper_path = wrapper_path
			self.previous_window_handle = window_handle
			self.elements = find_elements(wrapper_path, window_handle)
			self.searched_again_rectangle = None
		self.unique_wrapper_path = None
		return self.STRATEGIES[self.i_strategy]

	def search_again(self, wrapper_rectangle):
		"""
		Searches again the elements of the path of the element under the mouse cursor, without the cache of find_elements.

		The elements found are kept 10 seconds by find_elements, and while the mouse cursor stays on elements of the same
		path: they may not be the elements of the application anymore, if it changed (another page, for example). Then
		the element under the mouse cursor is not among them, or one of them vanished. The elements are searched again
		once for each element under the mouse cursor, and the search starts again with the first strategy.

		:param wrapper_rectangle: The rectangle of the element under the mouse cursor
		:return: True if the elements were searched again, False if they were already searched again for this element
		"""
		if self.searched_again_rectangle is not None and self.searched_again_rectangle == wrapper_rectangle:
			return False
		self.searched_again_rectangle = wrapper_rectangle
		find_elements.cache_clear()
		self.elements = find_elements(self.previous_wrapper_path, self.previous_window_handle)
		self.i_strategy = 0
		return True


class RecorderMode(str, Enum):
	"""
	The modes of the recorder.

	It is a 'str' enumeration, so a mode is equal to its value: RecorderMode.RECORD == "Record".
	The code that compares the mode of the recorder with a string keeps working.
	"""

	INITIALIZING = "Initializing"
	RECORD = "Record"
	PLAY = "Play"
	INFO = "Info"
	STOP = "Stop"
	QUIT = "Quit"


class Recorder(Thread):
	"""
	Recorder is a class thread used to record UI events in clipboard or in a file.

	.. code-block:: python
		:caption: Example of code using 'Recorder'::
		
		from pywinauto_recorder.recorder import Recorder
		from pywinauto_recorder.player import UIPath, click, move, playback

		recorder = Recorder()
		recorder.start_recording()
		with UIPath("Untitled - Notepad||Window"):
			doc = move("Text editor||Document")
			time.sleep(0.5)
			click(doc)
			utf8 = move("||Pane-> UTF-8||Text")
			time.sleep(0.5)
			click(utf8)
		recorded_python_script = recorder.stop_recording()
		recorder.quit()
		playback(filename=recorded_python_script)

	The above code clicks on the text field and then on 'utf8'.
	All these events are recorded in a file. Then the file is replayed.
	"""

	def __init__(self):
		"""Creates the recorder, and starts its loop in a thread, then the process of the element information tooltip."""
		from .element_observer import ElementInfoTooltip
		Thread.__init__(self)
		from win32api import GetSystemMetrics
		self._loop_t0 = None
		self.screen_width = GetSystemMetrics(0)
		self.screen_height = GetSystemMetrics(1)
		self.main_overlay = oaam.Overlay(transparency=0.4)
		self.desktop = pywinauto.Desktop(backend='uia', allow_magic_lookup=False)
		self.daemon = True
		self.event_list = []
		self._copy_count = 0
		self._mode = RecorderMode.INITIALIZING
		self._process_menu_click_mode = False
		self._smart_mode = False
		self._relative_coordinate_mode = False
		self.wrapper_old_info_tip = None
		self.common_path_info_tip = ""
		self.last_element_event = None
		self.started_recording_with_keyboard = False
		self.is_on_tray_icon = _no_tray_icon
		"""
		A function is_on_tray_icon(x, y) that returns True if a position of the screen is on the tray icon of the
		application. The presses of the mouse buttons on the tray icon are not recorded (see __is_ignored).
		"""
		self._mouse_position = None
		self._ignored_buttons = set()
		self._window_handle = None
		self._cursor_position_unavailable = False
		self.element_info_tooltip = ElementInfoTooltip()
		self.start()
	
	def __overlay_add_bold_rectangle(self, wrapper_rectangle, color=(0, 255, 0)):
		thickness = 5
		r = wrapper_rectangle
		sides = (  # top, bottom, left and right sides: x, y, width, height
			(r.left, r.top, r.width(), thickness), (r.left, r.bottom - thickness, r.width(), thickness),
			(r.left, r.top, thickness, r.height()), (r.right - thickness, r.top, thickness, r.height()))
		for x, y, width, height in sides:
			self.main_overlay.add(
				geometry=oaam.Shape.rectangle, x=x, y=y, width=width, height=height,
				thickness=0, color=(0, 128, 0), brush=oaam.Brush.solid, brush_color=color)
		
	def __find_unique_element_array_1d(self, wrapper_rectangle, elements, wrapper_path):
		"""
		Searches a reference element for the element under the mouse cursor (Strategy.array_1d, used in Smart mode).

		:param wrapper_rectangle: The rectangle of the element under the mouse cursor
		:param elements: The elements found with the path of the element
		:param wrapper_path: The path of the element: the reference element must be in the same window
		:return: The [] operator: "#[path of the reference element,column]", or None if no reference element is found
		"""
		window_title = get_entry_list(wrapper_path)[0]
		for row in get_sorted_rectangles(elements)[2]:
			for r_x, (rectangle, _) in enumerate(row):
				if rectangle == wrapper_rectangle:
					return self.__find_reference_element(rectangle, r_x, window_title)
		return None

	def __find_reference_element(self, wrapper_rectangle, column, window_title):
		"""
		Searches the reference element of an element: an element with a unique path, on its left, in the same window.

		:param wrapper_rectangle: The rectangle of the element
		:param column: The column of the element (see get_sorted_rectangles)
		:param window_title: The first entry of the path of the element: its window
		:return: The [] operator: "#[path of the reference element,column]", or None if no reference element is found
		"""
		xx, yy = wrapper_rectangle.left, wrapper_rectangle.mid_point()[1]
		previous_wrapper_path2 = None
		while xx > 0:  # TODO: limiter la recherche à la fenêtre courante
			xx = xx - 9
			wrapper2 = self.desktop.from_point(xx, yy)
			if wrapper2 is None:
				continue
			wrapper2_rectangle = wrapper2.rectangle()
			if wrapper2_rectangle.height() > wrapper_rectangle.height() * 2:
				continue
			wrapper_path2 = get_wrapper_path(wrapper2)
			if not wrapper_path2 or wrapper_path2 == previous_wrapper_path2 or get_entry_list(wrapper_path2)[0] != window_title:
				continue
			previous_wrapper_path2 = wrapper_path2
			if find_elements(wrapper_path2, self._window_handle):
				self.__overlay_add_bold_rectangle(wrapper2_rectangle, color=(0, 0, 255))
				self.__overlay_add_bold_rectangle(wrapper_rectangle, color=(255, 200, 0))
				return '#[' + wrapper_path2 + ',' + str(column) + ']'
		return None

	def __find_unique_element_array_2d(self, wrapper_rectangle, elements):
		"""
		Searches the element under the mouse cursor among the elements sorted in rows and columns (Strategy.array_2d).

		The element under the mouse cursor is displayed in orange, and the other elements in red.

		:param wrapper_rectangle: The rectangle of the element under the mouse cursor
		:param elements: The elements found with the path of the element
		:return: The [] operator: "#[row,column]", or None if the element is not among the sorted elements. Before, it
			was '': the path, which matches several elements, was taken for a unique path (issue #68).
		"""
		unique_array_2d = None
		for r_y, row in enumerate(get_sorted_rectangles(elements)[2]):
			for r_x, (rectangle, _) in enumerate(row):
				if rectangle == wrapper_rectangle:
					self.__overlay_add_bold_rectangle(rectangle, color=(255, 200, 0))
					unique_array_2d = '#[' + str(r_y) + ',' + str(r_x) + ']'
				else:
					self.__overlay_add_bold_rectangle(rectangle, color=(255, 0, 0))
		return unique_array_2d

	def __mouse_on(self, mouse_event):
		"""
		Records a mouse event in Record mode. The 'mouse' library calls it in its thread, in the order of the events.

		Consecutive mouse moves are recorded as one: the last one. The presses on Pywinauto Recorder, and their releases,
		are not recorded (see __is_ignored).
		"""
		if isinstance(mouse_event, mouse.MoveEvent):
			self._mouse_position = (mouse_event.x, mouse_event.y)
		if self.mode != RecorderMode.RECORD or self.__is_ignored(mouse_event):
			return
		if isinstance(mouse_event, mouse.MoveEvent) and (len(self.event_list) > 0):
			if isinstance(self.event_list[-1], mouse.MoveEvent):
				self.event_list = self.event_list[:-1]
		self.event_list.append(mouse_event)

	def __is_ignored(self, mouse_event):
		"""
		Returns True if a mouse event must not be recorded: a press on Pywinauto Recorder, or the release of this press.

		The presses on its tray icon (see is_on_tray_icon) and on the windows of its menus start or stop the recording:
		before, they were recorded, and the replay failed on them.
		"""
		if not isinstance(mouse_event, mouse.ButtonEvent):
			return False
		if mouse_event.event_type == 'up':
			ignored = mouse_event.button in self._ignored_buttons
			self._ignored_buttons.discard(mouse_event.button)
			return ignored
		x, y = self._mouse_position or win32api.GetCursorPos()
		if self.is_on_tray_icon(x, y) or _is_menu_of_this_process(x, y):
			self._ignored_buttons.add(mouse_event.button)
			return True
		self._ignored_buttons.discard(mouse_event.button)
		return False

	def __start_stop_recording_by_key(self):
		if self.mode != RecorderMode.RECORD:
			self.started_recording_with_keyboard = True
			self.start_recording()
		else:
			self.stop_recording()
	
	def __start_stop_displaying_info_by_key(self):
		self.toggle_element_info()
			
	def __display_found_element_by_key(self):
		if self.last_element_event:
			self._copy_count = 2
			x, y = win32api.GetCursorPos()
			l_e_e = self.last_element_event
			dx, dy = _compute_dx_dy(x, y, l_e_e.rectangle)
			i = l_e_e.path.find(path_separator)
			window_title = l_e_e.path[0:i]
			p = _get_relative_path(window_title, l_e_e.path)
			code = 'with UIPath(u"' + _escape_special_char(window_title) + '"):\n'
			code += '\twrapper = find(u"' + _escape_special_char(p)
			code += _relative_coordinates(dx, dy, self.relative_coordinate_mode)
			code += '")\n'
			code += '\twrapper.draw_outline()\n'
			pyperclip.copy(code)
			if self.event_list and self.mode == RecorderMode.RECORD:
				self.event_list.append(FindEvent(path=l_e_e.path, dx=dx, dy=dy, time=time.time()))
				
	def __key_on(self, e):
		if self.mode == RecorderMode.PLAY:
			# The shortcuts are ignored during a replay: CTRL+ALT+R would record the replayed actions,
			# and the end of the replay would then lose this recording.
			return
		if _is_hotkey_released(e, 'r', ("alt", "ctrl")):
			self.__start_stop_recording_by_key()
		elif _is_hotkey_released(e, 's', ("alt", "ctrl")):
			self.smart_mode = not self.smart_mode
		elif _is_hotkey_released(e, 'F', ("shift", "ctrl")):
			self.__display_found_element_by_key()
		elif _is_hotkey_released(e, 'D', ("shift", "ctrl")):
			self.__start_stop_displaying_info_by_key()
		elif self.mode == RecorderMode.RECORD:
			self.event_list.append(e)

	def run(self):
		"""
		The loop of the recorder: it searches a unique path of the element under the mouse cursor, and displays it.

		It stops when the mode of the recorder is "Quit". An error in a loop is printed (only once if it is raised again
		at each loop), and the next loop starts 0.1 s later.
		"""
		import comtypes.client
		print("")
		print("COMTYPES CACHE FOLDER:", comtypes.client._code_cache._find_gen_dir())
		
		dir_path = os.path.dirname(os.path.realpath(__file__))
		print("PYWINAUTO RECORDER FOLDER:", dir_path)
		
		keyboard.hook(self.__key_on)
		mouse.hook(self.__mouse_on)
		if not keyboard_hooks_work():
			print(
				"Warning: the keyboard hooks do not work. The keyboard shortcuts and the recording of the keyboard are "
				"not available: use the menu of the tray icon, or restart Pywinauto Recorder.")
		search = _ElementSearch()
		if self.mode == RecorderMode.INITIALIZING:
			# a replay of a file, or quit(), may have set the mode while the recorder was initializing
			self.mode = RecorderMode.INFO
		i = 0
		previous_loop_error = None
		while self.mode != RecorderMode.QUIT:
			i = i + 1
			loop_error = None
			try:
				self.__loop(search, i)
			except Exception as error:
				# An element that vanished is not an error: the application changed, the next loop searches again
				if not _is_vanished_element_error(error):
					exc_type, exc_value, exc_traceback = sys.exc_info()
					loop_error = repr(traceback.format_exception(exc_type, exc_value, exc_traceback))
					if loop_error != previous_loop_error:
						# An error that is raised again at each loop is printed only once
						print(loop_error)
				self.common_path_info_tip = ""
				self.wrapper_old_info_tip = None
				time.sleep(0.1)  # same pace as a loop without error: do not loop at full speed
			finally:
				previous_loop_error = loop_error
		mouse.unhook_all()
		keyboard.unhook_all()
		self.main_overlay.quit()
		print("Run end")

	def __cursor_position(self):
		"""
		Returns the position of the mouse cursor, or None if Windows does not give it.

		Windows refuses it ("Access is denied") while the secure desktop is displayed: locked session, UAC prompt or
		CTRL+ALT+DEL screen. Then a message is printed once, instead of an error at each loop (issue #68).
		"""
		try:
			position = win32api.GetCursorPos()
		except win32api.error:
			if not self._cursor_position_unavailable:
				print(
					"The position of the mouse cursor is not available (locked session, UAC prompt or CTRL+ALT+DEL "
					"screen): the recorder waits.")
			self._cursor_position_unavailable = True
			return None
		self._cursor_position_unavailable = False
		return position

	def __loop(self, search, i):
		"""
		A loop of the recorder: it searches a unique path of the element under the mouse cursor, and displays it.

		The loop lasts at least 0.1 s. It ends at once, without displaying anything, when there is no element under the
		mouse cursor, or when the mouse cursor leaves the element during the search.

		:param search: The state of the search, kept from a loop to the next one
		:param i: The number of the loop, displayed by the progress icon
		"""
		self._loop_t0 = time.time()
		self.main_overlay.clear_all()
		position = self.__cursor_position()
		if position is None:
			time.sleep(0.1)  # same pace as a loop
			return
		wrapper = self.desktop.from_point(*position)
		if wrapper is None:
			time.sleep(0.01)
			return
		# The elements are searched in the window under the mouse cursor. Before, this window was set with
		# set_native_window_handle, which also restricted the searches of a replay to this window.
		self._window_handle = _top_level_window(wrapper).handle
		wrapper_path = get_wrapper_path(wrapper)
		if not wrapper_path:
			time.sleep(0.01)
			return
		strategy = search.next_strategy(wrapper_path, self.smart_mode, self._window_handle)
		wrapper_rectangle = wrapper.rectangle()
		strategy, unique_wrapper_path, in_elements = self.__search_unique_path(
			search.elements, strategy, wrapper_path, wrapper_rectangle)
		if not in_elements and search.search_again(wrapper_rectangle):
			self.main_overlay.clear_all()  # the elements found before are not displayed
			strategy, unique_wrapper_path, _ = self.__search_unique_path(
				search.elements, Strategy.unique_path, wrapper_path, wrapper_rectangle)
		if strategy is None:
			search.i_strategy = 0  # the mouse cursor left the element: the search starts again with the first strategy
			return
		search.unique_wrapper_path = unique_wrapper_path
		if unique_wrapper_path is not None:
			self.last_element_event = ElementEvent(strategy, wrapper_rectangle, unique_wrapper_path)
			if self.event_list and self.mode == RecorderMode.RECORD:
				self.event_list.append(self.last_element_event)
		self.__display_mode_icons(i)
		self.main_overlay.refresh()
		
		loop_duration = time.time() - self._loop_t0
		while loop_duration < 0.1:
			time.sleep(0.01)
			loop_duration = time.time() - self._loop_t0
		time.sleep(0.01)  # main_overlay.clear_all() doit attendre la fin de main_overlay.refresh()

	def __search_unique_path(self, elements, strategy, wrapper_path, wrapper_rectangle):
		"""
		Searches a unique path of the element under the mouse cursor with a strategy, and displays the elements found.

		:param elements: The elements found with wrapper_path
		:param strategy: The strategy of the search
		:param wrapper_path: The path of the element
		:param wrapper_rectangle: The rectangle of the element
		:return: The strategy that was used, the unique path or None if it is not found, and False if the element under
			the mouse cursor is not among the elements found, or if one of them vanished: they may be the result of a
			previous search (see _ElementSearch.search_again). (None, None, True) if the mouse cursor left the element.
		"""
		try:
			if strategy is Strategy.unique_path:
				if not _is_cursor_in(wrapper_rectangle):
					return None, None, True
				return self.__search_with_unique_path(elements, wrapper_path, wrapper_rectangle)
			if strategy == Strategy.array_1d and elements:
				if not _is_cursor_in(wrapper_rectangle):
					return None, None, True
				unique_array_1d = self.__find_unique_element_array_1d(wrapper_rectangle, elements, wrapper_path)
				if unique_array_1d is not None:
					return strategy, wrapper_path + unique_array_1d, True
				strategy = Strategy.array_2d
			if strategy == Strategy.array_2d and elements:
				if not _is_cursor_in(wrapper_rectangle):
					return None, None, True
				unique_array_2d = self.__find_unique_element_array_2d(wrapper_rectangle, elements)
				if unique_array_2d is not None:
					return strategy, wrapper_path + unique_array_2d, True
			return strategy, None, False
		except COMError:  # an element found vanished: the application changed since the search
			return strategy, None, False

	def __search_with_unique_path(self, elements, wrapper_path, wrapper_rectangle):
		"""
		Searches the element under the mouse cursor with its path (Strategy.unique_path).

		The path is unique if only one element has this path: it is displayed in green. Otherwise, the elements found are
		displayed in red.

		:return: Strategy.unique_path, the path or None if it is not unique, and False if the element under the mouse
			cursor is not among the elements found
		"""
		rectangles = [element.rectangle() for element in elements]
		if len(rectangles) == 1:
			self.__overlay_add_bold_rectangle(wrapper_rectangle, color=(0, 255, 0))
			return Strategy.unique_path, wrapper_path, rectangles[0] == wrapper_rectangle
		for rectangle in rectangles:
			self.__overlay_add_bold_rectangle(rectangle, color=(255, 0, 0))
		return Strategy.unique_path, None, wrapper_rectangle in rectangles

	def __display_mode_icons(self, i):
		"""
		Displays the icons of the modes of the recorder, and waits while the recorder is stopped or replays a script.

		:param i: The number of the loop, displayed by the progress icon
		"""
		nb_icons = 0
		if self.mode == RecorderMode.RECORD:
			_overlay_add_mode_icon(self.main_overlay, IconSet.hicon_record, 10, 10)
			nb_icons += 1
		elif self.mode == RecorderMode.STOP:
			self.element_info_tooltip.hide()
			self.main_overlay.clear_all()
			_overlay_add_mode_icon(self.main_overlay, IconSet.hicon_stop, 10, 10)
			self.main_overlay.refresh()
			while self.mode == RecorderMode.STOP:
				time.sleep(0.1)
		elif self.mode == RecorderMode.PLAY:
			self.element_info_tooltip.hide()
			self.main_overlay.clear_all()
			_overlay_add_mode_icon(self.main_overlay, IconSet.hicon_play, 10, 10)
			self.main_overlay.refresh()
			while self.mode == RecorderMode.PLAY:
				time.sleep(1.0)
		if self.mode in (RecorderMode.RECORD, RecorderMode.INFO):
			_overlay_add_progress_icon(self.main_overlay, i, 10 + 60 * nb_icons, 10)
			nb_icons += 1
		if self.mode == RecorderMode.INFO:
			self.element_info_tooltip.show()
		if self.smart_mode:
			_overlay_add_mode_icon(self.main_overlay, IconSet.hicon_light_on, 10 + 60 * nb_icons, 10)
			nb_icons += 1
		if self._copy_count > 0:
			_overlay_add_mode_icon(self.main_overlay, IconSet.hicon_clipboard, 10 + 60 * nb_icons, 10)
			nb_icons += 1
			self._copy_count = self._copy_count - 1

	@property
	def process_menu_click_mode(self):
		"""
		If True, the process menu events are recorded else they are ignored.
	
		:return: The state of the process menu click mode.
		"""
		return self._process_menu_click_mode
	
	@process_menu_click_mode.setter
	def process_menu_click_mode(self, value):
		"""
		It sets the state of the process menu click mode.
		
		:param value: If the value is True, the process menu events are recorded else they are ignored.
		"""
		self._process_menu_click_mode = value

	@property
	def relative_coordinate_mode(self):
		"""
		If True, the relative coordinates are recorded else they are ignored.

		:return: The state of the relative coordinates mode.
		"""
		return self._relative_coordinate_mode
	
	@relative_coordinate_mode.setter
	def relative_coordinate_mode(self, value):
		"""
		It sets the state of the relative coordinates mode.

		:param value: If the value is True, the relative coordinates are recorded else they are ignored.
		"""
		self._relative_coordinate_mode = value

	@property
	def smart_mode(self):
		"""
		If True, the smart mode is activated.

		:return: The state of the smart mode.
		"""
		return self._smart_mode
	
	@smart_mode.setter
	def smart_mode(self, value):
		"""
		It sets the state of the smart mode.

		:param value: If the value is True, the smart mode is activated else it is not activated.
		"""
		self._smart_mode = value

	@property
	def mode(self):
		"""
		It returns the mode of the recorder.

		:return: The mode of the recorder, a RecorderMode. It is equal to its value, for example "Record".
		"""
		return self._mode
	
	@mode.setter
	def mode(self, value):
		"""
		It sets the mode of the recorder.

		:param value: A RecorderMode, or its value: "Record", "Play", "Info", "Stop" or "Quit"
		:raises ValueError: if the value is not a mode of the recorder
		"""
		self._mode = RecorderMode(value)

	def start_recording(self):
		"""
		It starts the recording: the mode of the recorder becomes "Record".

		It adds a mouse move event to the event list, displays the record icon to the main overlay,
		clears and refreshes the main and info overlays, and then sets the mode to "Record".
		"""
		time.sleep(0.6)  # wait the recorder to be fully ready
		x, y = win32api.GetCursorPos()
		self._ignored_buttons = set()
		self.event_list = [mouse.MoveEvent(x, y, time.time())]
		_overlay_add_mode_icon(self.main_overlay, IconSet.hicon_record, 10, 10)
		self.element_info_tooltip.hide()
		self.main_overlay.clear_all()
		self.main_overlay.refresh()
		self.mode = RecorderMode.RECORD

	def stop_recording(self):
		"""
		It stops the recording, and writes the Python script: the mode of the recorder becomes "Stop".

		It cleans the event list, displays the stop icon to the main overlay,
		clears and refreshes the main and info overlays, writes the Python script, and then sets the mode to "Stop".

		:return: The name of the file that was created.
		"""
		if self.mode == RecorderMode.RECORD and len(self.event_list) > 2:
			events = list(self.event_list)
			self.event_list = []
			self.mode = RecorderMode.STOP
			time.sleep(0.6)  # wait the recorder to be fully ready
			if self.started_recording_with_keyboard:
				_clean_events(events, remove_first_up=True)
			else:
				_clean_events(events)
			self.started_recording_with_keyboard = False
			_process_events(events, process_menu_click=self.process_menu_click_mode)
			_clean_events(events)
			return _write_in_file(events, relative_coordinate_mode=self.relative_coordinate_mode)
		self.main_overlay.clear_all()
		_overlay_add_mode_icon(self.main_overlay, IconSet.hicon_stop, 10, 10)
		self.main_overlay.refresh()
		self.mode = RecorderMode.STOP
		return None

	def get_last_element_event(self):
		"""
		It returns the last element of the event.

		:return: The last element event.
		"""
		return self.last_element_event
	
	def toggle_element_info(self):
		"""
		It starts or stops displaying element information.

		A recording in progress is stopped first with stop_recording(), so that the recorded script is written.
		"""
		if self.mode == RecorderMode.INFO:
			self.mode = RecorderMode.STOP
		else:
			if self.mode == RecorderMode.RECORD:
				self.stop_recording()
			self.mode = RecorderMode.INFO

	def playback(self, str_code='', filename='', wait_for_enter=True):
		"""
		It plays back a string of code or a Python file.
		
		:param str_code: The code to be played back
		:param filename: The name of the file corresponding to the code to be played back
		:param wait_for_enter: False to never wait for the Enter key after an error (see player.playback)
		:return: True if the code was played back without error, False otherwise
		"""
		self.mode = RecorderMode.PLAY
		self.main_overlay.refresh()
		try:
			return playback(str_code, filename, wait_for_enter=wait_for_enter)
		finally:
			self.mode = RecorderMode.STOP
	
	def quit(self):
		"""
		It quits the recorder: it ends its loop and stops the element information tooltip.

		It writes the recording in progress, clears the main overlay, stops the loop of the recorder, and then stops the
		element information tooltip.
		The end of the loop is waited for at most QUIT_TIMEOUT seconds, because an application that does not respond
		may block it. The thread of the recorder is a daemon thread: it does not prevent the application from exiting.
		"""
		if self.mode == RecorderMode.RECORD:
			self.stop_recording()  # before, the recording in progress was lost
		self.main_overlay.clear_all()
		self.main_overlay.refresh()
		self.mode = RecorderMode.QUIT
		self.join(timeout=QUIT_TIMEOUT)
		if self.is_alive():
			print("The recorder did not stop within " + str(QUIT_TIMEOUT) + " seconds.")
		del self.element_info_tooltip  # it stops the process of the tooltip, now that the loop does not use it
		print("Quit")


read_config_file()