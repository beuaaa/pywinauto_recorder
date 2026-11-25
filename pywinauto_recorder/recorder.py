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
	get_wrapper_path, get_entry_list, get_entry, get_sorted_region, \
	read_config_file, set_native_window_handle, SPECIAL_CHARACTERS
from .core import find_elements as not_ttl_cached_find_elements
from .player import playback
from cachetools import func


@func.ttl_cache(ttl=10)
def find_elements(full_element_path=None):
	return not_ttl_cached_find_elements(full_element_path=full_element_path)


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
	cx, cy = rectangle.mid_point()
	dx, dy = float(x - cx) / (rectangle.width() / 2 - 1), float(y - cy) / (rectangle.height() / 2 - 1)
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

	:param path: The full path of the element
	:param dx: The horizontal relative coordinate of the click (see _relative_coordinates)
	:param dy: The vertical relative coordinate of the click (see _relative_coordinates)
	:param common_path: The common path of the UIPath blocks, or ''
	:param relative_coordinate_mode: True if the relative coordinates are written
	:return: The text written between the quotes of the argument
	"""
	if common_path:
		path = _get_relative_path(common_path, path)
	return _escape_special_char(path) + _relative_coordinates(dx, dy, relative_coordinate_mode)


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
	i = 0
	while i < len(events):
		if isinstance(events[i], keyboard.KeyboardEvent):
			_process_keyboard_events(events, i)
		elif isinstance(events[i], mouse.WheelEvent):
			_process_wheel_events(events, i)
		i = i + 1
	i = len(events) - 1
	while i >= 0:
		if isinstance(events[i], mouse.ButtonEvent) and events[i].event_type == 'up':
			i = _process_drag_and_drop_or_click_events(events, i)
		i = i - 1
	if process_menu_click:
		i = len(events) - 1
		while i >= 0:
			if isinstance(events[i], ClickEvent):
				i = _process_menu_select_events(events, i)
			i = i - 1


def _process_keyboard_events(events, i):
	keyboard_events = [events[i]]
	i0 = i + 1
	i_processed_events = []
	while i0 < len(events):
		if isinstance(events[i0], keyboard.KeyboardEvent):
			keyboard_events.append(events[i0])
			i_processed_events.append(i0)
			i0 = i0 + 1
		elif isinstance(events[i0], ElementEvent):
			i0 = i0 + 1
		else:
			break
	line = _get_send_keys_strings(keyboard_events)
	for i_p_e in sorted(i_processed_events, reverse=True):
		del events[i_p_e]
	if line:
		events[i] = SendKeysEvent(line=line)


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


def _find_button_down(events, i, move_event_end):
	"""
	Searches the press of the mouse button released by events[i], and counts the clicks.

	:param events: The events of the recording
	:param i: The index of a release of a mouse button
	:param move_event_end: The last mouse move before the release
	:return: The index of the press (-1 if it is not found), True if the mouse moved while the button was pressed,
		and the number of clicks (2 or 3 for a double or triple click)
	"""
	drag_and_drop = False
	click_count = 0
	i0 = i - 1
	while i0 >= 0:
		event = events[i0]
		if isinstance(event, mouse.MoveEvent):
			if event.x != move_event_end.x or event.y != move_event_end.y:
				drag_and_drop = True
		elif isinstance(event, mouse.ButtonEvent) and event.event_type in ('down', 'double'):
			click_count = click_count + 1
			if event.event_type == 'down' or click_count == 3:
				return i0, drag_and_drop, click_count
		i0 = i0 - 1
	return -1, drag_and_drop, click_count


def _remove_mouse_and_element_events(events, i1, i):
	"""
	Removes the mouse button, mouse move and element events from events[i1] to events[i - 1].

	:return: The new index of the event of index i
	"""
	removed = [i0 for i0 in range(max(i1, 0), i) if isinstance(events[i0], (mouse.ButtonEvent, mouse.MoveEvent, ElementEvent))]
	for i0 in reversed(removed):
		del events[i0]
	return i - len(removed)


def _process_drag_and_drop_or_click_events(events, i):
	"""
	Replaces the release of a mouse button, events[i], by a ClickEvent or a DragAndDropEvent.

	The mouse button, mouse move and element events from the press to the release of the button are removed.

	:param events: The events of the recording
	:param i: The index of a release of a mouse button
	:return: The new index of the ClickEvent or of the DragAndDropEvent
	"""
	i0, element_event_before_button_up = _find_previous_event(events, i - 1, ElementEvent)
	_, move_event_end = _find_previous_event(events, i0, mouse.MoveEvent)
	i1, drag_and_drop, click_count = _find_button_down(events, i, move_event_end)
	i0, element_event_before_button_down = _find_previous_event(events, i1, ElementEvent)
	if drag_and_drop:
		_, move_event_start = _find_previous_event(events, i0, mouse.MoveEvent)
		dx1, dy1 = _compute_dx_dy(move_event_start.x, move_event_start.y, element_event_before_button_down.rectangle)
		dx2, dy2 = _compute_dx_dy(move_event_end.x, move_event_end.y, element_event_before_button_up.rectangle)
		events[i] = DragAndDropEvent(
			path=element_event_before_button_down.path, dx1=dx1, dy1=dy1,
			path2=element_event_before_button_up.path, dx2=dx2, dy2=dy2)
	else:
		up_event = events[i]
		dx, dy = _compute_dx_dy(move_event_end.x, move_event_end.y, element_event_before_button_down.rectangle)
		events[i] = ClickEvent(
			button=up_event.button, click_count=click_count,
			path=element_event_before_button_down.path, dx=dx, dy=dy, time=up_event.time)
	return _remove_mouse_and_element_events(events, i1, i)


def _get_relative_path(common_path, path):
	if not path:
		return ''
	# TODO: check if common_path is the beginning of path
	path = path[len(common_path) + len(path_separator):]
	entry_list = get_entry_list(path)
	str_name, str_type, y_x, dx_dy = get_entry(entry_list[-1])
	if (y_x is not None) and not is_int(y_x[0]):
		y_x[0] = y_x[0][len(common_path) + 2:]
		path = path_separator.join(entry_list[:-1]) + path_separator + str_name
		if path == path_separator:
			path = ''
		path = path + type_separator + str_type + "#[" + y_x[0] + "," + str(y_x[1]) + "]"
		if dx_dy is not None and dx_dy[0] != 0 and dx_dy[1] != 0:
			path = path + "%(" + str(dx_dy[0]) + "," + str(dx_dy[1]) + ")"
	return path


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


def _common_start(sa, sb):
	"""Returns the longest common substring from the beginning of sa and sb."""
	
	def _iter():
		for a, b in zip(sa, sb):
			if a == b:
				yield a
			else:
				return
	
	return ''.join(_iter())


def _get_typed_keys(keyboard_events):
	string = ''
	previous_event = None
	for event in keyboard_events:
		event_name = event.name.replace('windows gauche', 'left windows')
		event_name = event_name.replace('windows droite', 'right windows')
		if previous_event:
			common_event_name = _common_start(event.name, previous_event.name)
			if common_event_name:
				if previous_event.event_type == 'down' and event.event_type == 'up':
					if len(common_event_name) == 1:
						if previous_event and len(previous_event.name) == 1:
							string = string[:-len('""{? down}"')]
						else:
							string = string[:-len('{? down}"')]
						string = string + event_name + '"'
						previous_event = event
						continue
					else:
						string = string[:-len(' down}"')] + '}"'
						previous_event = event
						continue
		previous_event = event
		if event_name in keyboard.all_modifiers | {'maj', 'enter'}:
			string = string + '"' + "{VK_"
			if 'left' in event_name:
				string = string + "L"
			if 'right' in event_name or 'gr' in event_name:
				string = string + "R"
			if 'alt' in event_name:
				string = string + "MENU"
			elif 'ctrl' in event_name:
				string = string + "CONTROL"
			elif 'shift' in event_name or 'maj' in event_name:
				string = string + "SHIFT"
			elif 'windows' in event_name:
				string = string + "WIN"
			elif 'enter' in event_name:
				string = string[:-len("VK_")] + "ENTER"
			string = string + ' ' + event.event_type + "}" + '"'
		else:
			string = string + '"{' + event_name + ' ' + event.event_type + '}"'
	return string


def _get_typed_strings(keyboard_events, allow_backspace=True):
	"""
	Given a sequence of events, tries to deduce what strings were typed.

	Strings are separated when a non-textual key is pressed (such as tab or
	enter). Characters are converted to uppercase according to shift and
	capslock status. If `allow_backspace` is True, backspaces remove the last
	character typed. Control keys are converted into pywinauto.keyboard key codes.
	"""
	backspace_name = 'backspace'
	
	shift_pressed = False
	capslock_pressed = False
	string = ''
	for event in keyboard_events:
		name = event.name
		
		# Space is the only key that we _parse_hotkey to the spelled out name
		# because of legibility. Now we have to undo that.
		if event.name == 'space':
			name = ' '
		
		if 'shift' in event.name:
			shift_pressed = event.event_type == 'down'
		elif event.name == 'caps lock' and event.event_type == 'down':
			capslock_pressed = not capslock_pressed
		elif allow_backspace and event.name == backspace_name and event.event_type == 'down':
			string = string[:-1]
		elif event.event_type == 'down':
			if len(name) == 1:
				if shift_pressed ^ capslock_pressed:
					name = name.upper()
				string = string + name
			else:
				if string:
					yield '"' + _escape_special_char(string) + '"'
				if 'windows' in event.name:
					yield '"' + '{LWIN}' + '"'
				elif 'enter' in event.name:
					yield '"' + '{ENTER}' + '"'
				string = ''


def _get_send_keys_strings(keyboard_events):
	is_typed_words = True
	alnum_count = 0
	for event in keyboard_events:
		if event.name in keyboard.all_modifiers:
			is_typed_words = False
			break
		if event.name.isalnum():
			alnum_count += 1
			if alnum_count > 1:
				break
	if alnum_count <= 1:
		is_typed_words = False
	if is_typed_words:
		return ''.join(format(code) for code in _get_typed_strings(keyboard_events))
	else:
		return _get_typed_keys(keyboard_events)


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
	"""Returns the top level window of an element: the element itself, or its ancestor whose parent is the desktop."""
	window = wrapper
	while window.parent() is not None and window.parent().parent() is not None:
		window = window.parent()
	return window


def _is_cursor_in(rectangle):
	"""Returns True if the mouse cursor is inside the rectangle, and not on its border."""
	x, y = win32api.GetCursorPos()
	return rectangle.left < x < rectangle.right and rectangle.top < y < rectangle.bottom


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
		self.unique_wrapper_path = None
		self.i_strategy = 0

	def next_strategy(self, wrapper_path, smart_mode):
		"""
		Returns the strategy of the search of a unique path of the element of path wrapper_path.

		The elements of path wrapper_path are searched again when the element under the mouse cursor changes.

		:param wrapper_path: The path of the element under the mouse cursor
		:param smart_mode: True if the Smart mode is active
		:return: The strategy
		"""
		if wrapper_path == self.previous_wrapper_path:
			if (self.unique_wrapper_path is None) or (self.STRATEGIES[self.i_strategy] == Strategy.array_2d):
				self.i_strategy = min(self.i_strategy + 1, len(self.STRATEGIES) - 1)
				if (not smart_mode) and (self.STRATEGIES[self.i_strategy] == Strategy.array_1d):
					self.i_strategy = 1
		else:
			self.i_strategy = 0
			self.previous_wrapper_path = wrapper_path
			self.elements = find_elements(wrapper_path)
		self.unique_wrapper_path = None
		return self.STRATEGIES[self.i_strategy]


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
		
	def __find_unique_element_array_1d(self, wrapper_rectangle, elements):
		nb_y, nb_x, candidates = get_sorted_region(elements)
		window_title = get_entry_list((get_wrapper_path(elements[0])))[0]
		for r_y in range(nb_y):
			for r_x in range(nb_x):
				try:
					r = candidates[r_y][r_x].rectangle()
				except IndexError:
					continue
				if r == wrapper_rectangle:
					xx, yy = r.left, r.mid_point()[1]
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
						if not wrapper_path2:
							continue
						if wrapper_path2 == previous_wrapper_path2:
							continue
						if get_entry_list(wrapper_path2)[0] != window_title:
							continue
						previous_wrapper_path2 = wrapper_path2
						
						if find_elements(get_wrapper_path(wrapper2)):
							self.__overlay_add_bold_rectangle(wrapper2_rectangle, color=(0, 0, 255))
							self.__overlay_add_bold_rectangle(wrapper_rectangle, color=(255, 200, 0))
							return '#[' + wrapper_path2 + ',' + str(r_x) + ']'
					return None
		return None

	def __find_unique_element_array_2d(self, wrapper_rectangle, elements):
		nb_y, nb_x, candidates = get_sorted_region(elements)
		unique_array_2d = ''
		for r_y in range(nb_y):
			for r_x in range(nb_x):
				try:
					r = candidates[r_y][r_x].rectangle()
				except IndexError:
					continue
				if r == wrapper_rectangle:
					self.__overlay_add_bold_rectangle(r, color=(255, 200, 0))
					unique_array_2d = '#[' + str(r_y) + ',' + str(r_x) + ']'
				else:
					self.__overlay_add_bold_rectangle(r, color=(255, 0, 0))
		return unique_array_2d

	def __mouse_on(self, mouse_event):
		if self.mode == RecorderMode.RECORD:
			if isinstance(mouse_event, mouse.MoveEvent) and (len(self.event_list) > 0):
				if isinstance(self.event_list[-1], mouse.MoveEvent):
					self.event_list = self.event_list[:-1]
			self.event_list.append(mouse_event)
			
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
			except Exception:
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
		wrapper = self.desktop.from_point(*win32api.GetCursorPos())
		if wrapper is None:
			time.sleep(0.01)
			return
		set_native_window_handle(_top_level_window(wrapper).handle)
		wrapper_path = get_wrapper_path(wrapper)
		if not wrapper_path:
			time.sleep(0.01)
			return
		strategy = search.next_strategy(wrapper_path, self.smart_mode)
		wrapper_rectangle = wrapper.rectangle()
		strategy, unique_wrapper_path = self.__search_unique_path(
			search.elements, strategy, wrapper, wrapper_path, wrapper_rectangle)
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

	def __search_unique_path(self, elements, strategy, wrapper, wrapper_path, wrapper_rectangle):
		"""
		Searches a unique path of the element under the mouse cursor with a strategy, and displays the elements found.

		:param elements: The elements found with wrapper_path
		:param strategy: The strategy of the search
		:param wrapper: The element under the mouse cursor
		:param wrapper_path: The path of the element
		:param wrapper_rectangle: The rectangle of the element
		:return: The strategy that was used, and the unique path, or None if it is not found.
			(None, None) if the mouse cursor left the element.
		"""
		unique_wrapper_path = None
		if strategy is Strategy.unique_path:
			if not _is_cursor_in(wrapper_rectangle):
				return None, None
			if len(elements) == 1:
				unique_wrapper_path = wrapper_path
				self.__overlay_add_bold_rectangle(wrapper_rectangle, color=(0, 255, 0))
			else:
				for e in elements:
					self.__overlay_add_bold_rectangle(e.rectangle(), color=(255, 0, 0))
		if strategy == Strategy.array_1d and elements:
			if not _is_cursor_in(wrapper_rectangle):
				return None, None
			unique_array_1d = self.__find_unique_element_array_1d(wrapper.rectangle(), elements)
			if unique_array_1d is not None:
				unique_wrapper_path = wrapper_path + unique_array_1d
			else:
				strategy = Strategy.array_2d
		if strategy == Strategy.array_2d and elements:
			if not _is_cursor_in(wrapper_rectangle):
				return None, None
			unique_array_2d = self.__find_unique_element_array_2d(wrapper.rectangle(), elements)
			if unique_array_2d is not None:
				unique_wrapper_path = wrapper_path + unique_array_2d
		return strategy, unique_wrapper_path

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

	def playback(self, str_code='', filename=''):
		"""
		It plays back a string of code or a Python file.
		
		:param str_code: The code to be played back
		:param filename: The name of the file corresponding to the code to be played back
		:return: True if the code was played back without error, False otherwise
		"""
		self.mode = RecorderMode.PLAY
		self.main_overlay.refresh()
		try:
			return playback(str_code, filename)
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