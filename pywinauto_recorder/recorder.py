# -*- coding: utf-8 -*-

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
import math
import codecs
from .core import *

ElementEvent = namedtuple('ElementEvent', ['strategy', 'rectangle', 'path'])
SendKeysEvent = namedtuple('SendKeysEvent', ['line'])
MouseWheelEvent = namedtuple('MouseWheelEvent', ['delta'])
DragAndDropEvent = namedtuple('DragAndDropEvent', ['path', 'dx1', 'dy1', 'path2', 'dx2', 'dy2'])
ClickEvent = namedtuple('ClickEvent', ['button', 'click_count', 'path', 'dx', 'dy', 'time'])
FindEvent = namedtuple('FindEvent', ['path', 'dx', 'dy', 'time'])
MenuEvent = namedtuple('MenuEvent', ['path', 'menu_path', 'menu_type'])


class IconSet:
	dir_path = os.path.dirname(os.path.realpath(__file__))
	hicon_clipboard = oaam.load_ico(dir_path + r'\..\Icons\paste.ico', 48, 48)
	hicon_light_on = oaam.load_ico(dir_path + r'\..\Icons\light-on.ico', 48, 48)
	hicon_record = oaam.load_ico(dir_path + r'\..\Icons\record.ico', 48, 48)
	hicon_play = oaam.load_ico(dir_path + r'\..\Icons\play.ico', 48, 48)
	hicon_stop = oaam.load_ico(dir_path + r'\..\Icons\stop.ico', 48, 48)
	hicon_search = oaam.load_ico(dir_path + r'\..\Icons\search.ico', 48, 48)
	hicon_power = oaam.load_ico(dir_path + r'\..\Icons\power.ico', 48, 48)


def escape_special_char(string):
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


def compute_dx_dy(x, y, rectangle):
	cx, cy = rectangle.mid_point()
	dx, dy = float(x - cx) / (rectangle.width() / 2 - 1), float(y - cy) / (rectangle.height() / 2 - 1)
	return (dx, dy)


def write_in_file(events, relative_coordinate_mode=False, menu_click=False):
	from pathlib import Path
	home_dir = Path.home() / Path('Pywinauto recorder')
	home_dir.mkdir(parents=True, exist_ok=True)
	record_file_name = home_dir / Path('recorded ' + time.asctime().replace(':', '_') + '.py')
	print('Recording in file: ' + str(record_file_name.absolute()))
	script = "# encoding: {}\n\n".format(sys.getdefaultencoding())
	script += u"import os, sys\n"
	script += u"script_dir = os.path.dirname(__file__)\n"
	script += u"sys.path.append(script_dir)\n"
	script += "from pywinauto_recorder.player import *\n\n"
	common_path = ''
	common_window = ''
	common_region = ''
	i = 0
	while i < len(events):
		e_i = events[i]
		if type(e_i) in [DragAndDropEvent, ClickEvent, FindEvent, MenuEvent]:
			if e_i.path != common_path:
				new_common_path = find_new_common_path_in_next_user_events(events, i)
				if new_common_path != common_path:
					common_path = new_common_path
					entry_list = get_entry_list(common_path)
					e_i_window = entry_list[0]
					if e_i_window != common_window:
						common_window = e_i_window
						script += '\nwith Window(u"' + escape_special_char(common_window) + '"):\n'
					e_i_region = path_separator.join(entry_list[1:])
					if e_i_region != common_region and e_i_region:
						common_region = e_i_region
						script += '\twith Region(u"' + escape_special_char(common_region) + '"):\n'
					else:
						common_region = ''
		if type(e_i) in [SendKeysEvent, MouseWheelEvent, DragAndDropEvent, ClickEvent, FindEvent, MenuEvent]:
			if common_window:
				script += '\t'
				if common_region:
					script += '\t'
			if type(e_i) is SendKeysEvent:
				script += 'send_keys(' + e_i.line + ')\n'
			elif type(e_i) is MouseWheelEvent:
				script += 'mouse_wheel(' + str(e_i.delta) + ')\n'
			elif type(e_i) is DragAndDropEvent:
				p1, p2 = e_i.path, e_i.path2
				dx1, dy1 = "{:.2f}".format(round(e_i.dx1 * 100, 2)), "{:.2f}".format(round(e_i.dy1 * 100, 2))
				dx2, dy2 = "{:.2f}".format(round(e_i.dx2 * 100, 2)), "{:.2f}".format(round(e_i.dy2 * 100, 2))
				if common_path:
					p1 = get_relative_path(common_path, p1)
					p2 = get_relative_path(common_path, p2)
				script += 'drag_and_drop(u"' + escape_special_char(p1)
				if relative_coordinate_mode and eval(dx1) != 0 and eval(dy1) != 0:
					script += '%(' + dx1 + ',' + dy1 + ')'
				script += '", u"' + escape_special_char(p2)
				if relative_coordinate_mode and eval(dx2) != 0 and eval(dy2) != 0:
					script += '%(' + dx2 + ',' + dy2 + ')'
				script += '")\n'
			elif type(e_i) is ClickEvent:
				p = e_i.path
				dx, dy = "{:.2f}".format(round(e_i.dx * 100, 2)), "{:.2f}".format(round(e_i.dy * 100, 2))
				if common_path:
					p = get_relative_path(common_path, p)
				str_c = ['', '', 'double_', 'triple_']
				script += str_c[e_i.click_count] + e_i.button + '_click(u"' + escape_special_char(p)
				if relative_coordinate_mode and eval(dx) != 0 and eval(dy) != 0:
					script += '%(' + dx + ',' + dy + ')'
				script += '")\n'
			elif type(e_i) is FindEvent:
				p = e_i.path
				dx, dy = "{:.2f}".format(round(e_i.dx * 100, 2)), "{:.2f}".format(round(e_i.dy * 100, 2))
				if common_path:
					p = get_relative_path(common_path, p)
				script += 'wrapper = find(u"' + escape_special_char(p)
				if relative_coordinate_mode and eval(dx) != 0 and eval(dy) != 0:
					script += '%(' + dx + ',' + dy + ')'
				script += '")\n'
			elif type(e_i) is MenuEvent:
				p, m_p = e_i.path, e_i.menu_path
				if common_path:
					p = get_relative_path(common_path, p)
				script += 'menu_click(u"' + escape_special_char(p) + '", r"' + escape_special_char(m_p) + '"'
				if e_i.menu_type == 'NPP':
					script += ', menu_type="NPP")\n'
				else:
					script += ')\n'
		i = i + 1
	with codecs.open(record_file_name, "w", encoding=sys.getdefaultencoding()) as f:
		f.write(script)
	pyperclip.copy(script)
	return record_file_name


def clean_events(events, remove_first_up=False):
	""""
	removes duplicate or useless events
	removes all the last down events due or not to (CRTL+ALT+r) when ending record mode
	:param remove_first_up: when True, removes the 2 first up events due to (CRTL+ALT+r) when starting record mode
	:param events: the copy of recorded event list
	"""
	if remove_first_up:
		i = 0
		i_up = 0
		while i < len(events):
			if type(events[i]) is keyboard.KeyboardEvent and events[i].event_type == 'up':
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
			if type(events[i]) in [ElementEvent, mouse.MoveEvent]:
				del events[i - 1]
			else:
				previous_event_type = type(events[i])
				i = i + 1
		else:
			previous_event_type = type(events[i])
			i = i + 1
	i = len(events) - 1
	while i > 0:
		if type(events[i]) is keyboard.KeyboardEvent and events[i].event_type == 'down':
			events.pop(i)
		elif type(events[i]) is keyboard.KeyboardEvent and events[i].event_type == 'up':
			break
		i = i - 1


def process_events(events, process_menu_click=True):
	i = 0
	while i < len(events):
		if type(events[i]) is keyboard.KeyboardEvent:
			process_keyboard_events(events, i)
		elif type(events[i]) is mouse.WheelEvent:
			process_wheel_events(events, i)
		i = i + 1
	i = len(events) - 1
	while i >= 0:
		if type(events[i]) is mouse.ButtonEvent and events[i].event_type == 'up':
			i = process_drag_and_drop_or_click_events(events, i)
		i = i - 1
	if process_menu_click:
		i = len(events) - 1
		while i >= 0:
			if type(events[i]) is ClickEvent:
				i = process_menu_select_events(events, i)
			i = i - 1


def process_keyboard_events(events, i):
	keyboard_events = [events[i]]
	i0 = i + 1
	i_processed_events = []
	while i0 < len(events):
		if type(events[i0]) == keyboard.KeyboardEvent:
			keyboard_events.append(events[i0])
			i_processed_events.append(i0)
			i0 = i0 + 1
		elif type(events[i0]) == ElementEvent:
			i0 = i0 + 1
		else:
			break
	line = get_send_keys_strings(keyboard_events)
	for i_p_e in sorted(i_processed_events, reverse=True):
		del events[i_p_e]
	if line:
		events[i] = SendKeysEvent(line=line)


def process_wheel_events(events, i):
	delta = events[i].delta
	i_processed_events = []
	i0 = i + 1
	while i0 < len(events):
		if type(events[i0]) == mouse.WheelEvent:
			delta = delta + events[i0].delta
			i_processed_events.append(i0)
			i0 = i0 + 1
		elif type(events[i0]) in [ElementEvent, mouse.MoveEvent]:
			i0 = i0 + 1
		else:
			break
	for i_p_e in sorted(i_processed_events, reverse=True):
		del events[i_p_e]
	events[i] = MouseWheelEvent(delta=delta)


def process_drag_and_drop_or_click_events(events, i):
	i0 = i - 1
	while i0 >= 0:
		if type(events[i0]) == ElementEvent:
			element_event_before_button_up = events[i0]
			break
		i0 = i0 - 1
	while i0 >= 0:
		if type(events[i0]) == mouse.MoveEvent:
			move_event_end = events[i0]
			break
		i0 = i0 - 1
	i0 = i - 1
	drag_and_drop = False
	click_count = 0
	while i0 >= 0:
		if type(events[i0]) == mouse.MoveEvent:
			if events[i0].x != move_event_end.x or events[i0].y != move_event_end.y:
				drag_and_drop = True
		elif type(events[i0]) == mouse.ButtonEvent and events[i0].event_type in ['down', 'double']:
			click_count = click_count + 1
			if events[i0].event_type == 'down' or click_count == 3:
				i1 = i0
				break
		i0 = i0 - 1
	element_event_before_button_down = None
	while i0 >= 0:
		if type(events[i0]) == ElementEvent:
			element_event_before_button_down = events[i0]
			break
		i0 = i0 - 1
	if drag_and_drop:
		move_event_start = None
		while i0 >= 0:
			if type(events[i0]) == mouse.MoveEvent:
				move_event_start = events[i0]
				break
			i0 = i0 - 1
		dx1, dy1 = compute_dx_dy(move_event_start.x, move_event_start.y, element_event_before_button_down.rectangle)
		dx2, dy2 = compute_dx_dy(move_event_end.x, move_event_end.y, element_event_before_button_up.rectangle)
		events[i] = DragAndDropEvent(
			path=element_event_before_button_down.path, dx1=dx1, dy1=dy1,
			path2=element_event_before_button_up.path, dx2=dx2, dy2=dy2)
	else:
		up_event = events[i]
		dx, dy = compute_dx_dy(move_event_end.x, move_event_end.y, element_event_before_button_down.rectangle)
		events[i] = ClickEvent(
			button=up_event.button, click_count=click_count,
			path=element_event_before_button_down.path, dx=dx, dy=dy, time=up_event.time)
	i_processed_events = []
	i0 = i - 1
	while i0 >= i1:
		if type(events[i0]) in [mouse.ButtonEvent, mouse.MoveEvent, ElementEvent]:
			i_processed_events.append(i0)
		i0 = i0 - 1
	for i_p_e in sorted(i_processed_events, reverse=True):
		del events[i_p_e]
		i = i - 1
	return i


def get_relative_path(common_path, path):
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


def find_common_path(current_path, next_path):
	current_entry_list = get_entry_list(current_path)
	if len(current_entry_list) > 1:
		_, _, y_x, _ = get_entry(current_entry_list[-1])
		if (y_x is not None) and not is_int(y_x[0]):
			current_entry_list = get_entry_list(y_x[0])[:-1]
		else:
			current_entry_list = current_entry_list[:-1]
	
	next_entry_list = get_entry_list(next_path)
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


def find_new_common_path_in_next_user_events(events, i):
	path_i = events[i].path
	i0 = i + 1
	new_common_path = ''
	while i0 < len(events):
		e = events[i0]
		if type(e) in [DragAndDropEvent, ClickEvent, FindEvent, MenuEvent]:
			new_common_path = find_common_path(path_i, e.path)
			break
		elif type(e) in [ElementEvent, mouse.MoveEvent]:
			i0 = i0 + 1
		else:
			break
	if new_common_path == '':
		new_common_path = find_common_path(path_i, path_i)
	return new_common_path


def process_menu_select_events(events, i):
	i0 = i
	i_processed_events = []
	menu_bar_path = None
	menu_path = []
	menu_type = 'QT'
	while i0 >= 0:
		if type(events[i0]) is ClickEvent:
			entry_list = get_entry_list(events[i0].path)
			str_name, str_type, _, _ = get_entry(entry_list[-1])
			if str_type == 'MenuItem':
				str_name_root, str_type_root, _, _ = get_entry(entry_list[0])
				if str_name_root == 'Context' and str_type_root == 'Menu':
					break
				menu_path.append(str_name)
				str_name, str_type, _, _ = get_entry(entry_list[-2])
				if str_type == 'MenuBar':
					if str_name:
						menu_type = 'NPP'
					# menu_path.append(str_name)
					menu_bar_path = path_separator.join(entry_list[0:-2])
					menu_path = path_separator.join(reversed(menu_path))
					break
				else:
					i_processed_events.append(i0)
		i0 = i0 - 1
	if menu_bar_path:
		events[i0] = MenuEvent(path=menu_bar_path, menu_path=menu_path, menu_type=menu_type)
		for i_p_e in sorted(i_processed_events, reverse=True):
			del events[i_p_e]
			i = i - 1
	return i


def common_start(sa, sb):
	""" returns the longest common substring from the beginning of sa and sb """
	
	def _iter():
		for a, b in zip(sa, sb):
			if a == b:
				yield a
			else:
				return
	
	return ''.join(_iter())


def get_typed_keys(keyboard_events):
	string = ''
	previous_event = None
	for event in keyboard_events:
		event_name = event.name.replace('windows gauche', 'left windows')
		event_name = event_name.replace('windows droite', 'right windows')
		if previous_event:
			common_event_name = common_start(event.name, previous_event.name)
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


def get_typed_strings(keyboard_events, allow_backspace=True):
	"""
	Given a sequence of events, tries to deduce what strings were typed.
	Strings are separated when a non-textual key is pressed (such as tab or
	enter). Characters are converted to uppercase according to shift and
	capslock status. If `allow_backspace` is True, backspaces remove the last
	character typed. Control keys are converted into pywinauto.keyboard key codes
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
					yield '"' + escape_special_char(string) + '"'
				if 'windows' in event.name:
					yield '"' + '{LWIN}' + '"'
				elif 'enter' in event.name:
					yield '"' + '{ENTER}' + '"'
				string = ''


def get_send_keys_strings(keyboard_events):
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
		return ''.join(format(code) for code in get_typed_strings(keyboard_events))
	else:
		return get_typed_keys(keyboard_events)


def overlay_add_progress_icon(main_overlay, i, x, y):
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=x, y=y, width=52, height=52,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
	main_overlay.add(
		geometry=oaam.Shape.triangle,
		xyrgb_array=((x + 1, y + 1, 255, 255, 254), (x + 1, y + 52, 128, 128, 128), (x + 51, y + 52, 255, 255, 254)),
		thickness=0)
	for b in range(i % 6):
		main_overlay.add(
			geometry=oaam.Shape.rectangle, x=x + 6, y=y + 6 + b * 8, width=40, height=6,
			color=(0, 255, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 200, 0))


def overlay_add_mode_icon(main_overlay, hicon, x, y):
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
	def __init__(self):
		Thread.__init__(self)
		from win32api import GetSystemMetrics
		self.screen_width = GetSystemMetrics(0)
		self.screen_height = GetSystemMetrics(1)
		self.main_overlay = oaam.Overlay(transparency=0.4)
		self.info_overlay = oaam.Overlay(transparency=0.0)
		self.info_overlay2 = oaam.Overlay(transparency=0.0)
		self.desktop = pywinauto.Desktop(backend='uia', allow_magic_lookup=False)
		self.daemon = True
		self.event_list = []
		self._mode = "Info"
		self._process_menu_click_mode = True
		self._smart_mode = False
		self._relative_coordinate_mode = False
		self._display_info_tip_mode = False
		self.x_info_tip = None
		self.y_info_tip = None
		self.path_info_tip = None
		self.last_element_event = None
		self.started_recording_with_keyboard = False
		self.start()

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
					while xx > 0:  # TODO: limiter la recherche à la fenètre courante
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
						
						entry_list2 = get_entry_list(wrapper_path2)
						unique_candidate2, _ = find_element(self.desktop, entry_list2, window_candidates=[])
						
						if unique_candidate2 is not None:
							r = wrapper2_rectangle
							self.main_overlay.add(
								geometry=oaam.Shape.rectangle, x=r.left, y=r.top,
								width=r.width(), height=r.height(),
								thickness=1, color=(0, 0, 255), brush=oaam.Brush.solid,
								brush_color=(0, 0, 255))
							r = wrapper_rectangle
							self.main_overlay.add(
								geometry=oaam.Shape.rectangle, x=r.left, y=r.top,
								width=r.width(), height=r.height(),
								thickness=1, color=(255, 0, 0), brush=oaam.Brush.solid,
								brush_color=(255, 200, 0))
							return '#[' + wrapper_path2 + ',' + str(r_x) + ']'
					else:
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
					color = (255, 200, 0)
					unique_array_2d = '#[' + str(r_y) + ',' + str(r_x) + ']'
				else:
					color = (255, 0, 0)
				self.main_overlay.add(
					geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(),
					height=r.height(),
					thickness=1, color=(255, 0, 0), brush=oaam.Brush.solid, brush_color=color)
		return unique_array_2d

	def __mouse_on(self, mouse_event):
		if self.mode == "Record":
			if (type(mouse_event) == mouse.MoveEvent) and (len(self.event_list) > 0):
				if type(self.event_list[-1]) == mouse.MoveEvent:
					self.event_list = self.event_list[:-1]
			self.event_list.append(mouse_event)

	def __key_on(self, e):
		if (
				e.name == 'r' and e.event_type == 'up'
				and keyboard.key_to_scan_codes("alt")[0] in keyboard._pressed_events
				and keyboard.key_to_scan_codes("ctrl")[0] in keyboard._pressed_events):
			if self.mode != "Record":
				self.started_recording_with_keyboard = True
				self.start_recording()
			else:
				self.stop_recording()
		elif (
				(e.name == 's') and (e.event_type == 'up')
				and keyboard.key_to_scan_codes("alt")[0] in keyboard._pressed_events
				and keyboard.key_to_scan_codes("ctrl")[0] in keyboard._pressed_events):
			self.smart_mode = not self.smart_mode
		elif (
				(e.name == 'F') and (e.event_type == 'up')
				and keyboard.key_to_scan_codes("shift")[0] in keyboard._pressed_events
				and keyboard.key_to_scan_codes("ctrl")[0] in keyboard._pressed_events):
			if self.last_element_event:
				x, y = win32api.GetCursorPos()
				l_e_e = self.last_element_event
				dx, dy = compute_dx_dy(x, y, l_e_e.rectangle)
				str_dx, str_dy = "{:.2f}".format(round(dx * 100, 2)), "{:.2f}".format(round(dy * 100, 2))
				#self.info_overlay.add(
				#	geometry=oaam.Shape.image, hicon=IconSet.hicon_clipboard, x=x, y=l_e_e.rectangle.top - 70)
				info_left = (dx+1)/2 * (self.screen_width - 500)
				info_top = (dy+1)/2 * (self.screen_height - 50)
				overlay_add_mode_icon(self.info_overlay, IconSet.hicon_clipboard, info_left, info_top)
				i = l_e_e.path.find(path_separator)
				window_title = l_e_e.path[0:i]
				# element_path = l_e_e.path[i+len(path_separator):]
				p = get_relative_path(window_title, l_e_e.path)
				code = 'with Window(u"' + escape_special_char(window_title) + '"):\n'
				code += '\twrapper = find(u"' + escape_special_char(p)
				if self.relative_coordinate_mode and eval(str_dx) != 0 and eval(str_dy) != 0:
					code += '%(' + str_dx + ',' + str_dy + ')'
				code += '")\n'
				pyperclip.copy(code)
				if self.event_list:
					self.event_list.append(FindEvent(path=l_e_e.path, dx=dx, dy=dy, time=time.time()))
		elif self.mode == "Record":
			self.event_list.append(e)

	def __display_info_tip(self, x, y, wrapper):
		if (self.x_info_tip == x) and (self.y_info_tip == y):
			return
		
		tooltip_width = 500
		tooltip_height = 25
		r = wrapper.rectangle()
		
		self.x_info_tip = x
		self.y_info_tip = y
		
		common_path = ""
		if self.path_info_tip:
			common_path = find_common_path(get_wrapper_path(wrapper), self.path_info_tip)
		self.path_info_tip = get_wrapper_path(wrapper)
		end_path = self.path_info_tip[len(common_path)::]
		
		text = ''
		text_width = 0
		for c in common_path:
			if c == '\n':
				text_width = 0
				tooltip_height = tooltip_height + 16
			else:
				text_width = text_width + 6.8
			text = text + c
			if text_width > tooltip_width:
				text = text + '\n'
				text_width = 0
				tooltip_height = tooltip_height + 16
				
		tooltip2_height = 25
		text2 = ''
		text_width = 0
		for c in end_path:
			if c == '\n':
				text_width = 0
				tooltip2_height = tooltip2_height + 16
			else:
				text_width = text_width + 6.8
			text2 = text2 + c
			if text_width > tooltip_width:
				text2 = text2 + '\n'
				text_width = 0
				tooltip2_height = tooltip2_height + 16
		dx, dy = (x - r.left) / r.width(), (y - r.top) / r.height()
		info_left = dx * (self.screen_width - tooltip_width)
		info_top = dy * (self.screen_height - (tooltip_height+tooltip2_height))
		self.info_overlay.clear_all()
		self.info_overlay2.clear_all()
		
		self.info_overlay2.add(
			geometry=oaam.Shape.rectangle, x=info_left, y=info_top - 2, width=tooltip_width, height=tooltip_height,
			thickness=1, color=(0, 0, 0), brush=oaam.Brush.solid, brush_color=(222, 254, 255))
		self.info_overlay.add(
			x=info_left + 5, y=info_top + 1, width=tooltip_width,
			height=25,
			text=text,
			text_format="win32con.DT_LEFT|win32con.DT_TOP|win32con.DT_WORDBREAK|win32con.DT_NOCLIP|win32con.DT_VCENTER",
			font_size=16, text_color=(0, 0, 0), color=(254, 255, 255),
			geometry=oaam.Shape.rectangle, thickness=0
		)
		self.info_overlay2.add(
			geometry=oaam.Shape.rectangle, x=info_left, y=info_top - 2 + tooltip_height, width=tooltip_width, height=tooltip2_height,
			thickness=1, color=(0, 0, 0), brush=oaam.Brush.solid, brush_color=(180, 254, 255))
		self.info_overlay.add(
			x=info_left + 5, y=info_top + tooltip_height, width=tooltip_width,
			height=25,
			text=text2,
			text_format="win32con.DT_LEFT|win32con.DT_TOP|win32con.DT_WORDBREAK|win32con.DT_NOCLIP|win32con.DT_VCENTER",
			font_size=16, text_color=(0, 0, 0), color=(254, 255, 255),
			geometry=oaam.Shape.rectangle, thickness=0
		)
		text = ""
		try:
			has_get_value = getattr(wrapper, "get_value", None)
			if callable(has_get_value):
				text = "wrapper.get_value(): " + wrapper.get_value()
		except:
			has_get_value = False
			pass
		if wrapper.legacy_properties()['Value']:
			if not has_get_value or (has_get_value and wrapper.get_value() != wrapper.legacy_properties()['Value']):
				if text:
					text = text + "\n"
				text = text + " wrapper.legacy_properties()['Value']: " + wrapper.legacy_properties()['Value']
		_, str_type, _, _ = get_entry(end_path)
		if str_type in ["RadioButton", "CheckBox"]:
			if text:
				text = text + "\n"
			text = text + " wrapper.legacy_properties()['State']: " + str(wrapper.legacy_properties()['State'])
		tooltip3_height = 25
		text3 = ''
		text_width = 0
		for c in text:
			if c == '\n':
				text_width = 0
				tooltip3_height = tooltip3_height + 16
			else:
				text_width = text_width + 6.8
			text3 = text3 + c
			if text_width > tooltip_width:
				text3 = text3 + '\n'
				text_width = 0
				tooltip3_height = tooltip3_height + 16
		if text3:
			self.info_overlay2.add(
				geometry=oaam.Shape.rectangle, x=info_left, y=info_top - 4 + tooltip_height + tooltip2_height, width=tooltip_width,
				height=tooltip3_height, thickness=1, color=(0, 0, 0), brush=oaam.Brush.solid, brush_color=(2, 254, 255))
			self.info_overlay.add(
				x=info_left + 5, y=info_top + tooltip_height + tooltip2_height, width=tooltip_width,
				height=tooltip3_height,
				text=text3,
				text_format="win32con.DT_LEFT|win32con.DT_TOP|win32con.DT_WORDBREAK|win32con.DT_NOCLIP|win32con.DT_VCENTER",
				font_size=16, text_color=(0, 0, 0), color=(254, 255, 255),
				geometry=oaam.Shape.rectangle, thickness=0
			)
		self.info_overlay2.refresh()
		#time.sleep(0.1)
		#self.info_overlay.refresh()

	def run(self):
		import comtypes.client
		print("COMPTYPES CACHE FOLDER:", comtypes.client._code_cache._find_gen_dir())
		
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
		unique_candidate = None
		elements = []
		i = 0
		previous_wrapper_path = None
		unique_wrapper_path = None
		strategies = [Strategy.unique_path, Strategy.array_2D, Strategy.array_1D]
		i_strategy = 0
		while self.mode != "Quit":
			try:
				self.main_overlay.clear_all()
				# self.info_overlay.clear_all()
				x, y = win32api.GetCursorPos()
				wrapper = self.desktop.from_point(x, y)
				if wrapper is None:
					continue
				wrapper_path = get_wrapper_path(wrapper)
				if not wrapper_path:
					continue
				if wrapper_path == previous_wrapper_path:
					if (unique_wrapper_path is None) or (strategies[i_strategy] == Strategy.array_2D):
						i_strategy = i_strategy + 1
						if (not self.smart_mode) and (strategies[i_strategy] == Strategy.array_1D):
							i_strategy = 1
						if i_strategy >= len(strategies):
							i_strategy = len(strategies) - 1
				else:
					i_strategy = 0
					previous_wrapper_path = wrapper_path
					entry_list = get_entry_list(wrapper_path)
					unique_candidate, elements = find_element(self.desktop, entry_list, window_candidates=[])
				strategy = strategies[i_strategy]
				unique_wrapper_path = None
				if strategy == Strategy.unique_path:
					if unique_candidate is not None:
						unique_wrapper_path = get_wrapper_path(unique_candidate)
						r = wrapper.rectangle()
						self.main_overlay.add(
							geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(), height=r.height(),
							thickness=1, color=(0, 128, 0), brush=oaam.Brush.solid, brush_color=(0, 255, 0))
					else:
						for e in elements:
							r = e.rectangle()
							self.main_overlay.add(
								geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(), height=r.height(),
								thickness=1, color=(0, 128, 0), brush=oaam.Brush.solid, brush_color=(255, 0, 0))
				if strategy == Strategy.array_1D and elements:
					unique_array_1d = self.__find_unique_element_array_1d(wrapper.rectangle(), elements)
					if unique_array_1d is not None:
						unique_wrapper_path = wrapper_path + unique_array_1d
					else:
						strategy = Strategy.array_2D
				if strategy == Strategy.array_2D and elements:
					unique_array_2d = self.__find_unique_element_array_2d(wrapper.rectangle(), elements)
					if unique_array_2d is not None:
						unique_wrapper_path = wrapper_path + unique_array_2d
				if unique_wrapper_path is not None:
					self.last_element_event = ElementEvent(strategy, wrapper.rectangle(), unique_wrapper_path)
					if self.event_list and unique_wrapper_path is not None:
						self.event_list.append(self.last_element_event)
				nb_icons = 0
				if self.mode == "Record":
					overlay_add_mode_icon(self.main_overlay, IconSet.hicon_record, 10, 10)
					#overlay_add_record_icon(self.main_overlay, 10, 10)
					nb_icons = nb_icons + 1
				elif self.mode == "Play":
					self.info_overlay.clear_all()
					self.info_overlay2.clear_all()
					self.main_overlay.clear_all()
					overlay_add_mode_icon(self.main_overlay, IconSet.hicon_play, 10, 10)
					while self.mode == "Play":
						time.sleep(1.0)
					self.main_overlay.refresh()
					self.info_overlay2.refresh()
					self.main_overlay.refresh()
				#elif self.mode == "Info":
				#	overlay_add_pause_icon(self.main_overlay, 10, 10)
				elif self.mode == "Stop":
					self.info_overlay.clear_all()
					self.info_overlay2.clear_all()
					self.main_overlay.clear_all()
					overlay_add_mode_icon(self.main_overlay, IconSet.hicon_stop, 10, 10)
					self.main_overlay.refresh()
					self.info_overlay2.refresh()
					self.main_overlay.refresh()
					while self.mode == "Stop":
						time.sleep(1.0)
				if self.mode in ["Record", "Info"]:
					overlay_add_progress_icon(self.main_overlay, i, 10+60*nb_icons, 10)
					nb_icons = nb_icons + 1
					overlay_add_mode_icon(self.main_overlay, IconSet.hicon_search, 10 + 60 * nb_icons, 10)
					nb_icons = nb_icons + 1
				if self.smart_mode:
					overlay_add_mode_icon(self.main_overlay, IconSet.hicon_light_on, 10 + 60 * nb_icons, 10)
					nb_icons = nb_icons + 1
				i = i + 1
				
				if self.display_info_tip_mode:
					self.__display_info_tip(x, y, wrapper)
				self.main_overlay.refresh()
				#self.info_overlay.refresh()
				self.info_overlay.refresh()
				time.sleep(0.005)  # main_overlay.clear_all() doit attendre la fin de main_overlay.refresh()
			except Exception as e:
				exc_type, exc_value, exc_traceback = sys.exc_info()
				print(repr(traceback.format_exception(exc_type, exc_value, exc_traceback)))
		self.main_overlay.clear_all()
		self.main_overlay.refresh()
		if self.event_list:
			self.stop_recording()
		mouse.unhook_all()
		keyboard.unhook_all()
		print("Run end")

	@property
	def process_menu_click_mode(self):
		return self._process_menu_click_mode
	
	@process_menu_click_mode.setter
	def process_menu_click_mode(self, value):
		self._process_menu_click_mode = value

	@property
	def relative_coordinate_mode(self):
		return self._relative_coordinate_mode
	
	@relative_coordinate_mode.setter
	def relative_coordinate_mode(self, value):
		self._relative_coordinate_mode = value

	@property
	def display_info_tip_mode(self):
		return self._mode == "Info"
		return self._display_info_tip_mode
	
	@display_info_tip_mode.setter
	def display_info_tip_mode(self, value):
		if value:
			self._mode = "Info"
		else:
			self._mode = "Stop"
		#self._display_info_tip_mode = value

	@property
	def smart_mode(self):
		return self._smart_mode
	
	@smart_mode.setter
	def smart_mode(self, value):
		self._smart_mode = value

	@property
	def mode(self):
		return self._mode
	
	@mode.setter
	def mode(self, value):
		self._mode = value

	def start_recording(self):
		x, y = win32api.GetCursorPos()
		self.event_list = [mouse.MoveEvent(x, y, time.time())]
		overlay_add_mode_icon(self.main_overlay, IconSet.hicon_record, 10, 10)
		self.info_overlay.clear_all()
		self.info_overlay2.clear_all()
		self.main_overlay.clear_all()
		self.main_overlay.refresh()
		self.info_overlay2.refresh()
		self.info_overlay.refresh()
		self.mode = "Record"

	def stop_recording(self):
		if self.mode == "Record" and len(self.event_list) > 2:
			events = list(self.event_list)
			self.event_list = []
			self.mode = "Stop"
			if self.started_recording_with_keyboard:
				clean_events(events, remove_first_up=True)
			else:
				clean_events(events)
			self.started_recording_with_keyboard = False
			process_events(events, process_menu_click=self.process_menu_click_mode)
			clean_events(events)
			return write_in_file(events, relative_coordinate_mode=self.relative_coordinate_mode)
		self.main_overlay.clear_all()
		overlay_add_pause_icon(self.main_overlay, 10, 10)
		self.main_overlay.refresh()
		self.mode = "Stop"
		return None

	def get_last_element_event(self):
		return self.last_element_event

	def quit(self):
		self.info_overlay.clear_all()
		self.info_overlay2.clear_all()
		self.main_overlay.clear_all()
		self.main_overlay.refresh()
		self.info_overlay2.refresh()
		self.info_overlay.refresh()
		self.mode = 'Quit'
		print("Quit")
