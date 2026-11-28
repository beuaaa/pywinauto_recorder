"""This module contains functions and classes allowing to record a sequence of user actions.
All these functions and classes are private except the Recorder class.
"""

import sys
import os
import traceback
import time
import win32api
import win32con
from threading import Thread
import pywinauto
import overlay_arrows_and_more as oaam
import keyboard
import mouse
from collections import namedtuple
import pyperclip
import codecs
from .core import path_separator, type_separator, Strategy, is_int, \
                    get_wrapper_path, get_entry_list, get_entry, get_sorted_region, \
                    read_config_file
from .core import find_elements as not_ttl_cached_find_elements
#from .core import find_elements
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
	""" It loads the icons from the Icons folder and stores them in the class."""
	if "__compiled__" in globals():
		path_icons = os.path.dirname(os.path.realpath(__file__)) + r'\..'
	else:
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
		Is called on all paths to remove all special characters but it's not good.
		Should be moved in core
		Should be called only in get_wrapper_path and unescape_special_char in get_entry
		and in script += 'menu_click... in menu_click as it's already done
		should replace -> by _>
	"""
	for r in (("\\", "\\\\"), ("\t", "\\t"), ("\n", "\\n"), ("\r", "\\r"), ("\v", "\\v"), ("\f", "\\f"), ('"', '\\"')):
		string = string.replace(*r)
	return string


def _compute_dx_dy(x, y, rectangle):
	cx, cy = rectangle.mid_point()
	dx, dy = float(x - cx) / (rectangle.width() / 2 - 1), float(y - cy) / (rectangle.height() / 2 - 1)
	return (dx, dy)


def _write_in_file(events, relative_coordinate_mode=False):
	from pathlib import Path
	home_dir = Path.home() / 'Pywinauto recorder'
	home_dir.mkdir(parents=True, exist_ok=True)
	record_file_name = home_dir / Path('recorded ' + time.asctime().replace(':', '_') + '.py')
	print('Recording in file: ' + str(record_file_name.absolute()))
	script = "# encoding: {}\n\n".format(sys.getdefaultencoding())
	script += "from pywinauto_recorder.player import *\n\n"
	common_path = ''
	common_window = ''
	common_region = ''
	i = 0
	while i < len(events):
		e_i = events[i]
		if type(e_i) in (DragAndDropEvent, ClickEvent, FindEvent, MenuEvent):
			if e_i.path != common_path:
				new_common_path = _find_new_common_path_in_next_user_events(events, i)
				if new_common_path != common_path:
					common_path = new_common_path
					entry_list = get_entry_list(common_path)
					e_i_window = entry_list[0]
					if e_i_window != common_window:
						common_window = e_i_window
						script += '\nwith UIPath(u"' + _escape_special_char(common_window) + '"):\n'
					e_i_region = path_separator.join(entry_list[1:])
					if e_i_region != common_region and e_i_region:
						common_region = e_i_region
						script += '\twith UIPath(u"' + _escape_special_char(common_region) + '"):\n'
					else:
						common_region = ''
		if type(e_i) in (SendKeysEvent, MouseWheelEvent, DragAndDropEvent, ClickEvent, FindEvent, MenuEvent):
			if common_window:
				script += '\t'
				if common_region:
					script += '\t'
			if isinstance(e_i, SendKeysEvent):
				script += 'send_keys(' + e_i.line + ')\n'
			elif isinstance(e_i, MouseWheelEvent):
				script += 'mouse_wheel(' + str(e_i.delta) + ')\n'
			elif isinstance(e_i, DragAndDropEvent):
				p1, p2 = e_i.path, e_i.path2
				dx1, dy1 = "{:.2f}".format(round(e_i.dx1 * 100, 2)), "{:.2f}".format(round(e_i.dy1 * 100, 2))
				dx2, dy2 = "{:.2f}".format(round(e_i.dx2 * 100, 2)), "{:.2f}".format(round(e_i.dy2 * 100, 2))
				if common_path:
					p1 = _get_relative_path(common_path, p1)
					p2 = _get_relative_path(common_path, p2)
				script += 'drag_and_drop(u"' + _escape_special_char(p1)
				if relative_coordinate_mode and eval(dx1) != 0 and eval(dy1) != 0:
					script += '%(' + dx1 + ',' + dy1 + ')'
				script += '", u"' + _escape_special_char(p2)
				if relative_coordinate_mode and eval(dx2) != 0 and eval(dy2) != 0:
					script += '%(' + dx2 + ',' + dy2 + ')'
				script += '")\n'
			elif isinstance(e_i, ClickEvent):
				p = e_i.path
				dx, dy = "{:.2f}".format(round(e_i.dx * 100, 2)), "{:.2f}".format(round(e_i.dy * 100, 2))
				if common_path:
					p = _get_relative_path(common_path, p)
				str_c = ['', '', 'double_', 'triple_']
				if e_i.button == 'left':
					if e_i.count == 1:
						script += 'click(u"' + _escape_special_char(p)
					else:
						script += str_c[e_i.click_count] + 'click(u"' + _escape_special_char(p)
				else:
					script += str_c[e_i.click_count] + e_i.button + '_click(u"' + _escape_special_char(p)
				if relative_coordinate_mode and eval(dx) != 0 and eval(dy) != 0:
					script += '%(' + dx + ',' + dy + ')'
				script += '")\n'
			elif isinstance(e_i, FindEvent):
				p = e_i.path
				dx, dy = "{:.2f}".format(round(e_i.dx * 100, 2)), "{:.2f}".format(round(e_i.dy * 100, 2))
				if common_path:
					p = _get_relative_path(common_path, p)
				script += 'wrapper = find(u"' + _escape_special_char(p)
				if relative_coordinate_mode and eval(dx) != 0 and eval(dy) != 0:
					script += '%(' + dx + ',' + dy + ')'
				script += '")\n'
			elif isinstance(e_i, MenuEvent):
				script += 'menu_click(u"' + _escape_special_char(e_i.menu_path) + '")\n'
		i += 1
	with codecs.open(record_file_name, "w", encoding=sys.getdefaultencoding()) as f:
		f.write(script)
	pyperclip.copy(script)
	return record_file_name


def _clean_events(events, remove_first_up=False):
	""""
	removes duplicate or useless events
	removes all the last down events due or not to (CTRL+ALT+r) when ending record mode
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
		if type(events[i]) is previous_event_type:
			if type(events[i]) in (ElementEvent, mouse.MoveEvent):
				del events[i - 1]
			else:
				previous_event_type = type(events[i])
				i = i + 1
		else:
			previous_event_type = type(events[i])
			i = i + 1
	i = len(events) - 1
	while i > 0:
		if isinstance(events[i], keyboard.KeyboardEvent) and events[i].event_type == 'down':
			events.pop(i)
		elif isinstance(events[i], keyboard.KeyboardEvent) and events[i].event_type == 'up':
			break
		i = i - 1


def _process_events(events, process_menu_click=True):
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


def _process_drag_and_drop_or_click_events(events, i):
	i0 = i - 1
	while i0 >= 0:
		if isinstance(events[i0], ElementEvent):
			element_event_before_button_up = events[i0]
			break
		i0 = i0 - 1
	while i0 >= 0:
		if isinstance(events[i0], mouse.MoveEvent):
			move_event_end = events[i0]
			break
		i0 = i0 - 1
	i0 = i - 1
	drag_and_drop = False
	click_count = 0
	while i0 >= 0:
		if isinstance(events[i0], mouse.MoveEvent):
			if events[i0].x != move_event_end.x or events[i0].y != move_event_end.y:
				drag_and_drop = True
		elif isinstance(events[i0], mouse.ButtonEvent) and events[i0].event_type in ('down', 'double'):
			click_count = click_count + 1
			if events[i0].event_type == 'down' or click_count == 3:
				i1 = i0
				break
		i0 = i0 - 1
	element_event_before_button_down = None
	while i0 >= 0:
		if isinstance(events[i0], ElementEvent):
			element_event_before_button_down = events[i0]
			break
		i0 = i0 - 1
	if drag_and_drop:
		move_event_start = None
		while i0 >= 0:
			if isinstance(events[i0], mouse.MoveEvent):
				move_event_start = events[i0]
				break
			i0 = i0 - 1
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
	i_processed_events = []
	i0 = i - 1
	while i0 >= i1:
		if type(events[i0]) in (mouse.ButtonEvent, mouse.MoveEvent, ElementEvent):
			i_processed_events.append(i0)
		i0 = i0 - 1
	for i_p_e in sorted(i_processed_events, reverse=True):
		del events[i_p_e]
		i = i - 1
	return i


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


HOTKEY_RECORD = "CTRL+ALT+R"
"""Keyboard shortcut that starts or stops the recording (handled by Recorder.__key_on)."""
HOTKEY_SMART_MODE = "CTRL+ALT+S"
"""Keyboard shortcut that starts or stops the Smart mode (handled by Recorder.__key_on)."""
HOTKEY_ELEMENT_INFO = "CTRL+SHIFT+D"
"""Keyboard shortcut that starts or stops displaying element information (handled by Recorder.__key_on)."""
HOTKEY_COPY_CODE = "CTRL+SHIFT+F"
"""Keyboard shortcut that copies the Python code finding the element to the clipboard (handled by Recorder.__key_on)."""


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


def _overlay_add_progress_icon(main_overlay, i, x, y):
	global t0_progress_icon_timings
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=x, y=y, width=52, height=52,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
	main_overlay.add(
		geometry=oaam.Shape.triangle,
		xyrgb_array=((x + 1, y + 1, 255, 255, 254), (x + 1, y + 52, 128, 128, 128), (x + 51, y + 52, 255, 255, 254)),
		thickness=0)
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
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=x, y=y, width=52, height=52,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
	main_overlay.add(
		geometry=oaam.Shape.triangle,
		xyrgb_array=((x + 1, y + 1, 255, 255, 254), (x + 1, y + 52, 128, 128, 128), (x + 51, y + 52, 255, 255, 254)),
		thickness=0)
	main_overlay.add(
		geometry=oaam.Shape.image, hicon=hicon, x=int(x+2), y=int(y+2))


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
		self._mode = "Initializing"
		self._process_menu_click_mode = False
		self._smart_mode = False
		self._relative_coordinate_mode = False
		self.wrapper_old_info_tip = None
		self.common_path_info_tip = ""
		self.last_element_event = None
		self.started_recording_with_keyboard = False
		self._window_handle = None
		self.element_info_tooltip = ElementInfoTooltip()
		self.start()
	
	def __overlay_add_bold_rectangle(self, wrapper_rectangle, color=(0, 255, 0)):
		thickness = 5
		r = wrapper_rectangle
		self.main_overlay.add(
			geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(), height=thickness,
			thickness=0, color=(0, 128, 0), brush=oaam.Brush.solid, brush_color=color)
		self.main_overlay.add(
			geometry=oaam.Shape.rectangle, x=r.left, y=r.bottom - thickness, width=r.width(), height=thickness,
			thickness=0, color=(0, 128, 0), brush=oaam.Brush.solid, brush_color=color)
		self.main_overlay.add(
			geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=thickness, height=r.height(),
			thickness=0, color=(0, 128, 0), brush=oaam.Brush.solid, brush_color=color)
		self.main_overlay.add(
			geometry=oaam.Shape.rectangle, x=r.right - thickness, y=r.top, width=thickness, height=r.height(),
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
						
						if find_elements(get_wrapper_path(wrapper2), self._window_handle):
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
		if self.mode == "Record":
			if isinstance(mouse_event, mouse.MoveEvent) and (len(self.event_list) > 0):
				if isinstance(self.event_list[-1], mouse.MoveEvent):
					self.event_list = self.event_list[:-1]
			self.event_list.append(mouse_event)
			
	def __start_stop_recording_by_key(self):
		if self.mode != "Record":
			self.started_recording_with_keyboard = True
			self.start_recording()
		else:
			self.stop_recording()
	
	def __start_stop_displaying_info_by_key(self):
		if self.mode == "Info":
			self.mode = "Stop"
		else:
			self.mode = "Info"
			
	def __display_found_element_by_key(self):
		if self.last_element_event:
			self._copy_count = 2
			x, y = win32api.GetCursorPos()
			l_e_e = self.last_element_event
			dx, dy = _compute_dx_dy(x, y, l_e_e.rectangle)
			str_dx, str_dy = "{:.2f}".format(round(dx * 100, 2)), "{:.2f}".format(round(dy * 100, 2))
			i = l_e_e.path.find(path_separator)
			window_title = l_e_e.path[0:i]
			# element_path = l_e_e.path[i+len(path_separator):]
			p = _get_relative_path(window_title, l_e_e.path)
			code = 'with UIPath(u"' + _escape_special_char(window_title) + '"):\n'
			code += '\twrapper = find(u"' + _escape_special_char(p)
			if self.relative_coordinate_mode and eval(str_dx) != 0 and eval(str_dy) != 0:
				code += '%(' + str_dx + ',' + str_dy + ')'
			code += '")\n'
			code += '\twrapper.draw_outline()\n'
			pyperclip.copy(code)
			if self.event_list and self.mode == "Record":
				self.event_list.append(FindEvent(path=l_e_e.path, dx=dx, dy=dy, time=time.time()))
				
	def __key_on(self, e):
		key_to_scan_codes = keyboard.key_to_scan_codes
		if (
				(e.name, e.event_type) == ('r', 'up') and
				set([key_to_scan_codes("alt")[0], key_to_scan_codes("ctrl")[0]]).issubset(keyboard._pressed_events)):
			self.__start_stop_recording_by_key()
		elif (
				(e.name, e.event_type) == ('s', 'up') and
				set([key_to_scan_codes("alt")[0], key_to_scan_codes("ctrl")[0]]).issubset(keyboard._pressed_events)):
			self.smart_mode = not self.smart_mode
		elif (
				(e.name, e.event_type) == ('F', 'up') and
				set([key_to_scan_codes("shift")[0], key_to_scan_codes("ctrl")[0]]).issubset(keyboard._pressed_events)):
			self.__display_found_element_by_key()
		elif (
				(e.name, e.event_type) == ('D', 'up') and
				set([key_to_scan_codes("shift")[0], key_to_scan_codes("ctrl")[0]]).issubset(keyboard._pressed_events)):
			self.__start_stop_displaying_info_by_key()
		elif self.mode == "Record":
			self.event_list.append(e)


	# Ne fonctionne pas dans tous les cas car un rectangle père n'englobe pas forcément un rectangle fils
	# (par exemple un TreeItem) Mais peut être utilisé en ultime recours
	'''
	def my_from_point(self, x, y, wrapper=None):
		def get_children_of_children_with_empty_rectangle(wrapper_children):
			children_of_children_with_empty_rectangle = []
			for child in wrapper_children:
				child_rectangle = child.rectangle()
				if child_rectangle.width() ==0 or child_rectangle.width()==0:
					children_of_children_with_empty_rectangle += child.children()
			return children_of_children_with_empty_rectangle
		def get_children_of_children_not_in_parent_rectangle(wrapper_children):
			children_of_children_with_empty_rectangle = []
			for child in wrapper_children:
				child_rectangle = child.rectangle()
				grandchildren = child.children()
				if grandchildren:
					x = grandchildren[0].rectangle().left
					y = grandchildren[0].rectangle().top
					if not ((child_rectangle.left < x < child_rectangle.right) and (child_rectangle.top < y < child_rectangle.bottom)):
						return grandchildren
			return []
		if not wrapper:
			wrapper = self.desktop.top_from_point(x, y)
		wrapper_children = wrapper.children()
		for child in wrapper_children:
			child_rectangle = child.rectangle()
			if (child_rectangle.left < x < child_rectangle.right) and (child_rectangle.top < y < child_rectangle.bottom):
				return self.my_from_point(x, y, wrapper=child)
		for grandchild in get_children_of_children_with_empty_rectangle(wrapper_children):
			child_rectangle = grandchild.rectangle()
			if (child_rectangle.left < x < child_rectangle.right) and (child_rectangle.top < y < child_rectangle.bottom):
				return self.my_from_point(x, y, wrapper=grandchild)
			
		for grandchild in get_children_of_children_not_in_parent_rectangle(wrapper_children):
			child_rectangle = grandchild.rectangle()
			if (child_rectangle.left < x < child_rectangle.right) and (child_rectangle.top < y < child_rectangle.bottom):
				return self.my_from_point(x, y, wrapper=grandchild)

		return wrapper
	'''
	
	def run(self):
		"""
		The function is called in a loop, and it tries to find the unique element under the mouse cursor.
		"""
		import comtypes.client
		print("")
		print("COMTYPES CACHE FOLDER:", comtypes.client._code_cache._find_gen_dir())
		
		dir_path = os.path.dirname(os.path.realpath(__file__))
		print("PYWINAUTO RECORDER FOLDER:", dir_path)
		
		keyboard.hook(self.__key_on)
		mouse.hook(self.__mouse_on)
		keyboard.start_recording()
		win32api.keybd_event(160, 0, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)
		ev_list = keyboard.stop_recording()
		if not ev_list and os.path.isfile(dir_path + r"\pywinauto_recorder.exe"):
			print("Couldn't set keyboard hooks. Trying once again...\n")
			time.sleep(0.5)
			os.system(dir_path + r"\pywinauto_recorder.exe --no_splash_screen")
			sys.exit(1)
		elements = []
		i = 0
		previous_wrapper_path = None
		previous_window_handle = None
		unique_wrapper_path = None
		strategies = [Strategy.unique_path, Strategy.array_2D, Strategy.array_1D]
		i_strategy = 0
		self.mode = "Info"
		#strategy_unique_path_again_done = False
		previous_loop_error = None
		while self.mode != "Quit":
			i = i + 1
			loop_error = None
			try:
				self._loop_t0 = time.time()
				self.main_overlay.clear_all()
				cursor_pos = win32api.GetCursorPos()
				wrapper = self.desktop.from_point(*cursor_pos)
				#wrapper = self.my_from_point(*cursor_pos)
				if wrapper is None:
					time.sleep(0.01)
					continue
				
				""" """
				window = wrapper
				while window.parent() is not None and window.parent().parent() is not None:
					window = window.parent()
				# The elements are searched in the window under the mouse cursor. Before, this window was set with
				# set_native_window_handle, which also restricted the searches of a replay to this window.
				self._window_handle = window.handle
				""" """
				
				wrapper_path = get_wrapper_path(wrapper)
				if not wrapper_path:
					time.sleep(0.01)
					continue
				# The elements are searched again when the mouse cursor goes to another element, or to another window
				# that has the same title
				if (wrapper_path, self._window_handle) == (previous_wrapper_path, previous_window_handle):
					if (unique_wrapper_path is None) or (strategies[i_strategy] == Strategy.array_2D):
						i_strategy = i_strategy + 1
						if (not self.smart_mode) and (strategies[i_strategy] == Strategy.array_1D):
							i_strategy = 1
						if i_strategy >= len(strategies):
							i_strategy = len(strategies) - 1
				else:
					# strategy_unique_path_again_done = False
					i_strategy = 0
					previous_wrapper_path = wrapper_path
					previous_window_handle = self._window_handle
					elements = find_elements(wrapper_path, self._window_handle)
				#if wrapper_path == previous_wrapper_path and unique_wrapper_path:
				#	strategy = Strategy.unique_path_again
				# else:
				#	strategy = strategies[i_strategy]
				strategy = strategies[i_strategy]
				unique_wrapper_path = None
				# *** ----> this block of code must start a new while iteration if mouse cursor is outside wrapper rectangle
				# => add tests to leave if mouse cursor is outside wrapper rectangle
				wrapper_rectangle = wrapper.rectangle()
				
				#if strategy in [Strategy.unique_path, Strategy.unique_path_again]:
				if strategy is Strategy.unique_path:
					x_new, y_new = win32api.GetCursorPos()
					if not ((wrapper_rectangle.left < x_new < wrapper_rectangle.right) and (
							wrapper_rectangle.top < y_new < wrapper_rectangle.bottom)):
						i_strategy = 0
						continue
					if len(elements)==1:
						unique_wrapper_path = wrapper_path
						self.__overlay_add_bold_rectangle(wrapper_rectangle, color=(0, 255, 0))
					else:
						for e in elements:
							self.__overlay_add_bold_rectangle(e.rectangle(), color=(255, 0, 0))

				if strategy == Strategy.array_1D and elements:
					x_new, y_new = win32api.GetCursorPos()
					if not ((wrapper_rectangle.left < x_new < wrapper_rectangle.right) and (
							wrapper_rectangle.top < y_new < wrapper_rectangle.bottom)):
						i_strategy = 0
						continue
					unique_array_1d = self.__find_unique_element_array_1d(wrapper.rectangle(), elements)
					if unique_array_1d is not None:
						unique_wrapper_path = wrapper_path + unique_array_1d
					else:
						strategy = Strategy.array_2D
				if strategy == Strategy.array_2D and elements:
					x_new, y_new = win32api.GetCursorPos()
					if not ((wrapper_rectangle.left < x_new < wrapper_rectangle.right) and (
							wrapper_rectangle.top < y_new < wrapper_rectangle.bottom)):
						i_strategy = 0
						continue
					unique_array_2d = self.__find_unique_element_array_2d(wrapper.rectangle(), elements)
					if unique_array_2d is not None:
						unique_wrapper_path = wrapper_path + unique_array_2d
				# <----- ***
				if unique_wrapper_path is not None:
					#self.last_element_event = ElementEvent(strategy, wrapper.rectangle(), unique_wrapper_path)
					self.last_element_event = ElementEvent(strategy, wrapper_rectangle, unique_wrapper_path)
					if self.event_list and self.mode == "Record":
						self.event_list.append(self.last_element_event)
				nb_icons = 0
				if self.mode == "Record":
					_overlay_add_mode_icon(self.main_overlay, IconSet.hicon_record, 10, 10)
					nb_icons += 1
				elif self.mode == "Stop":
					self.element_info_tooltip.hide()
					self.main_overlay.clear_all()
					_overlay_add_mode_icon(self.main_overlay, IconSet.hicon_stop, 10, 10)
					self.main_overlay.refresh()
					while self.mode == "Stop":
						time.sleep(0.1)
				elif self.mode == "Play":
					self.element_info_tooltip.hide()
					self.main_overlay.clear_all()
					_overlay_add_mode_icon(self.main_overlay, IconSet.hicon_play, 10, 10)
					self.main_overlay.refresh()
					while self.mode == "Play":
						time.sleep(1.0)
				if self.mode in ("Record", "Info"):
					_overlay_add_progress_icon(self.main_overlay, i, 10 + 60 * nb_icons, 10)
					nb_icons += 1
				if self.mode == "Info":
					self.element_info_tooltip.show()
				if self.smart_mode:
					_overlay_add_mode_icon(self.main_overlay, IconSet.hicon_light_on, 10 + 60 * nb_icons, 10)
					nb_icons += 1
				if self._copy_count > 0:
					_overlay_add_mode_icon(self.main_overlay, IconSet.hicon_clipboard, 10 + 60 * nb_icons, 10)
					nb_icons += 1
					self._copy_count = self._copy_count - 1
				
				self.main_overlay.refresh()
				
				loop_duration = time.time() - self._loop_t0
				while loop_duration < 0.1:
					time.sleep(0.01)
					loop_duration = time.time() - self._loop_t0
				time.sleep(0.01)  # main_overlay.clear_all() doit attendre la fin de main_overlay.refresh()
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
		if self.event_list:
			self.stop_recording()
		mouse.unhook_all()
		keyboard.unhook_all()
		self.main_overlay.quit()
		print("Run end")

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
    It returns the mode of the recorder: "Record", "Play", "Info", "Stop", "Quit"
			
		:return: The mode of the recorder.
		"""
		return self._mode
	
	@mode.setter
	def mode(self, value):
		"""
		It sets the mode of the recorder: "Record", "Play", "Info", "Stop", "Quit"
		
		:param value: The mode of the recorder.
		"""
		self._mode = value

	def start_recording(self):
		"""
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
		self.mode = "Record"

	def stop_recording(self):
		"""
		It cleans the event list, displays the stop icon to the main overlay,
		clears and refreshes the main and info overlays, writes the Python script, and then sets the mode to "Stop".
		:return: The name of the file that was created.
		"""
		if self.mode == "Record" and len(self.event_list) > 2:
			events = list(self.event_list)
			self.event_list = []
			self.mode = "Stop"
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
		self.mode = "Stop"
		return None

	def get_last_element_event(self):
		"""
		It returns the last element of the event.
		:return: The last element event.
		"""
		return self.last_element_event
	
	def playback(self, str_code='', filename=''):
		"""
		This function plays back a string of code or a Python file.
		
		:param str_code: The code to be played back
		:param filename: The name of the file corresponding to the code to be played back
		"""
		self.mode = "Play"
		self.main_overlay.refresh()
		playback(str_code, filename)
		self.mode = "Stop"
	
	def quit(self):
		"""
		The function clears the main and info overlays, sets the mode to 'Quit', and then joins the thread.
		"""
		self.main_overlay.clear_all()
		self.main_overlay.refresh()
		del self.element_info_tooltip
		self.mode = 'Quit'
		self.join()
		print("Quit")

read_config_file()