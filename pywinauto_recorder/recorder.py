# -*- coding: utf-8 -*-

import sys
import os
import traceback
import time
import win32api
from threading import Thread
import pywinauto
import overlay_arrows_and_more as oaam
import keyboard
import mouse
import core
from collections import namedtuple
import pyperclip

ElementEvent = namedtuple('ElementEvent', ['strategy', 'rectangle', 'path'])
SendKeysEvent = namedtuple('SendKeysEvent', ['line'])
MouseWheelEvent = namedtuple('MouseWheelEvent', ['delta'])
DragAndDropEvent = namedtuple('DragAndDropEvent', ['path', 'dx1', 'dy1', 'path2', 'dx2', 'dy2'])
ClickEvent = namedtuple('ClickEvent', ['button', 'click_count', 'path', 'dx', 'dy', 'time'])
CommonPathEvent = namedtuple('CommonPathEvent', ['path'])
FindEvent = namedtuple('FindEvent', ['path', 'dx', 'dy', 'time'])
MenuEvent = namedtuple('MenuEvent', ['path', 'menu_path', 'menu_type'])


# reload(sys)
# sys.setdefaultencoding('utf-8')


def escape_special_char(string):
	for r in (("\t", "\\t"), ("\n", "\\n"), ("\r", "\\r"), ("\v", "\\v"), ("\f", "\\f"), ('"', '\\"')):
		string = string.replace(*r)
	return string


def write_in_file(events):
	new_path = 'Record files'
	if not os.path.exists(new_path):
		os.makedirs(new_path)
	record_file_name = './Record files/recorded ' + time.asctime() + '.py'
	record_file_name = record_file_name.replace(':', '_')
	print('Recording in file: ' + record_file_name)
	record_file = open(record_file_name, "w")
	record_file.write("# coding: utf-8\n\n")
	record_file.write("from pywinauto_recorder import *\n\n")
	i = 0
	common_path = ''
	common_window = ''
	while i < len(events):
		e_i = events[i]
		if type(e_i) is SendKeysEvent:
			if common_path:
				record_file.write('\t\t')
			record_file.write('send_keys(' + e_i.line + ')\n')
		elif type(e_i) is MouseWheelEvent:
			if common_path:
				record_file.write("\t\t")
			record_file.write('mouse_wheel(' + str(e_i.delta) + ')\n')
		elif type(e_i) is CommonPathEvent:
			if e_i.path != common_path:
				entry_list = core.get_entry_list(e_i.path)
				e_i_window = entry_list[0]
				if e_i_window != common_window:
					record_file.write('\nwith Window(u"' + escape_special_char(e_i_window) + '") as w:\n')
					common_window = e_i_window
				rel_path = core.path_separator.join(entry_list[1:])
				if rel_path:
					record_file.write('\twith Region(u"' + escape_special_char(rel_path) + '") as r:\n')
				else:
					record_file.write('\twith Region() as r:\n')
				common_path = e_i.path
		elif type(e_i) is DragAndDropEvent:
			p1, p2 = e_i.path, e_i.path2
			dx1, dy1 = "{:.2f}".format(round(e_i.dx1 * 100, 2)), "{:.2f}".format(round(e_i.dy1 * 100, 2))
			dx2, dy2 = "{:.2f}".format(round(e_i.dx2 * 100, 2)), "{:.2f}".format(round(e_i.dy2 * 100, 2))
			if common_path:
				p1 = get_relative_path(common_path, p1)
				p2 = get_relative_path(common_path, p2)
			record_file.write('\t\tr.drag_and_drop(u"' + escape_special_char(p1) + '%(' + dx1 + ',' + dy1 + ')", ')
			record_file.write('u"' + escape_special_char(p2) + '%(' + dx2 + ',' + dy2 + ')")\n')
		elif type(e_i) is ClickEvent:
			p = e_i.path
			dx, dy = "{:.2f}".format(round(e_i.dx * 100, 2)), "{:.2f}".format(round(e_i.dy * 100, 2))
			if common_path:
				p = get_relative_path(common_path, p)
			str_c = ['', '\t\tr.', '\tr.double_', '\tr.triple_']
			record_file.write(
				str_c[e_i.click_count] + e_i.button + '_click(u"' + escape_special_char(p) +
				'%(' + dx + ',' + dy + ')")\n')
		elif type(e_i) is FindEvent:
			p = e_i.path
			dx, dy = "{:.2f}".format(round(e_i.dx * 100, 2)), "{:.2f}".format(round(e_i.dy * 100, 2))
			record_file.write('\t\twrapper = r.find(u"' + escape_special_char(p) + '%(' + dx + ',' + dy + ')")\n')
		elif type(e_i) is MenuEvent:
			p, m_p = e_i.path, e_i.menu_path
			if common_path:
				p = get_relative_path(common_path, p)
			record_file.write('\t\tr.menu_click(u"' + escape_special_char(p) + '", r"' + escape_special_char(m_p) + '"')
			if e_i.menu_type == 'NPP':
				record_file.write(', menu_type="NPP")\n')
			else:
				record_file.write(')\n')
		i = i + 1
	record_file.close()
	with open(record_file_name, 'r') as my_file:
		data = my_file.read()
		pyperclip.copy(data)
	return record_file_name


def clean_events(events):
	""""
	remove duplicate or useless events
	:param events: the copy of recorded event list
	"""
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


def process_events(events):
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
	common_path = None
	while i < len(events):
		if type(events[i]) in [DragAndDropEvent, ClickEvent]:
			common_path = process_common_path_events(events, i, common_path)
		i = i + 1
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
		w_r = element_event_before_button_down.rectangle
		rx, ry = w_r.mid_point()
		dx1, dy1 = float(move_event_start.x - rx)/w_r.width(), float(move_event_start.y - ry)/w_r.height()
		w_r = element_event_before_button_up.rectangle
		rx, ry = w_r.mid_point()
		dx2, dy2 = float(move_event_end.x - rx)/w_r.width(), float(move_event_end.y - ry)/w_r.height()
		events[i] = DragAndDropEvent(
			path=element_event_before_button_down.path, dx1=dx1, dy1=dy1,
			path2=element_event_before_button_up.path, dx2=dx2, dy2=dy2)
	else:
		up_event = events[i]
		w_r = element_event_before_button_down.rectangle
		rx, ry = w_r.mid_point()
		dx, dy = float(move_event_end.x - rx) / w_r.width(), float(move_event_end.y - ry) / w_r.height()
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
	path = path[len(common_path) + len(core.path_separator):]
	entry_list = core.get_entry_list(path)
	str_name, str_type, y_x, dx_dy = core.get_entry(entry_list[-1])
	if (y_x is not None) and not core.is_int(y_x[0]):
		y_x[0] = y_x[0][len(common_path) + 2:]
		path = core.path_separator.join(entry_list[:-1]) + core.path_separator + str_name
		path = path + core.type_separator + str_type + "#[" + y_x[0] + "," + str(y_x[1]) + "]"
		if dx_dy is not None:
			path = path + "%(" + str(dx_dy[0]) + "," + str(dx_dy[1]) + ")"
	return path


def find_common_path(current_path, next_path):
	current_entry_list = core.get_entry_list(current_path)
	_, _, y_x, _ = core.get_entry(current_entry_list[-1])
	if (y_x is not None) and not core.is_int(y_x[0]):
		current_entry_list = core.get_entry_list(y_x[0])[:-1]
	else:
		current_entry_list = current_entry_list[:-1]

	next_entry_list = core.get_entry_list(next_path)
	next_entry_list = next_entry_list[:-1]
	n = 0
	try:
		while current_entry_list[n] == next_entry_list[n]:
			n = n + 1
	except IndexError:
		common_path = core.path_separator.join(current_entry_list[0:n])
		return common_path
	common_path = core.path_separator.join(current_entry_list[0:n])
	return common_path


def process_common_path_events(events, i, common_path):
	path_i = events[i].path
	i0 = i + 1
	new_common_path = ''
	while i0 < len(events):
		e = events[i0]
		if type(e) in [DragAndDropEvent, ClickEvent]:
			new_common_path = find_common_path(path_i, e.path)
			break
		elif type(e) in [ElementEvent, mouse.MoveEvent]:
			i0 = i0 + 1
		else:
			break
	if new_common_path == '':
		new_common_path = find_common_path(path_i, path_i)
	if new_common_path != common_path:
		events.insert(i, CommonPathEvent(path=new_common_path))
		return new_common_path
	return common_path


def process_menu_select_events(events, i):
	i0 = i
	i_processed_events = []
	menu_bar_path = None
	menu_path = []
	menu_type = 'QT'
	while i0 >= 0:
		if type(events[i0]) is ClickEvent:
			entry_list = core.get_entry_list(events[i0].path)
			str_name, str_type, _, _ = core.get_entry(entry_list[-1])
			if str_type == 'MenuItem':
				str_name, str_type, _, _ = core.get_entry(entry_list[0])
				if str_name == 'Context' and str_type == 'Menu':
					break
				menu_path.append(str_name)
				str_name, str_type, _, _ = core.get_entry(entry_list[-2])
				if str_type == 'MenuBar':
					if str_name:
						menu_type = 'NPP'
					menu_path.append(str_name)
					menu_bar_path = core.path_separator.join(entry_list[0:-2])
					menu_path = core.path_separator.join(reversed(menu_path))
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


def get_wrapper_path(wrapper):
	try:
		path = ''
		wrapper_top_level_parent = wrapper.top_level_parent()
		while wrapper != wrapper_top_level_parent:
			path = core.path_separator + wrapper.window_text() + core.type_separator + wrapper.element_info.control_type + path
			wrapper = wrapper.parent()
		return wrapper.window_text() + core.type_separator + wrapper.element_info.control_type + path
	except Exception as e:
		traceback.print_exc()
		print(e.message)
		return ''


def get_typed_keys(keyboard_events):
	string = ''
	for event in keyboard_events:
		if event.name in keyboard.all_modifiers | {'maj'}:
			string = string + '"' + "{VK_"
			if 'left' in event.name:
				string = string + "L"
			if 'right' in event.name or 'gr' in event.name:
				string = string + "R"
			if 'alt' in event.name:
				string = string + "MENU"
			elif 'ctrl' in event.name:
				string = string + "CONTROL"
			elif 'shift' in event.name or 'maj' in event.name:
				string = string + "SHIFT"
			elif 'windows' in event.name:
				string = string + "WIN"
			string = string + ' ' + event.event_type + "}" + '"'
		else:
			string = string + '"{' + event.name + ' ' + event.event_type + '}"'
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
					yield '"' + string + '"'
				if 'windows' in event.name:
					yield '"' + '{LWIN}' + '"'
				elif 'enter' in event.name:
					yield '"' + '{ENTER}' + '"'
				string = ''


def get_send_keys_strings(keyboard_events):
	print keyboard_events
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


def overlay_add_play_icon(main_overlay, x, y):
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=x, y=y, width=40, height=40,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
	main_overlay.add(
		geometry=oaam.Shape.triangle,
		xyrgb_array=((x + 5, y + 5, 0, 255, 0), (x + 5, y + 35, 0, 255, 0), (x + 35, y + 20, 0, 255, 0)))
	main_overlay.refresh()


def overlay_add_record_icon(main_overlay, x, y):
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=x, y=y, width=40, height=40,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
	main_overlay.add(
		geometry=oaam.Shape.ellipse, x=x + 5, y=y + 5, width=29, height=29,
		color=(255, 99, 99), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 0, 0))


def overlay_add_pause_icon(main_overlay, x, y):
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=x, y=y, width=40, height=40,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=x + 5, y=y + 5, width=12, height=30,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 0, 0))
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=x + 22, y=y + 5, width=12, height=30,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 0, 0))


def overlay_add_progress_icon(main_overlay, i, x, y):
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=x, y=y, width=40, height=40,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
	for b in range(i % 5):
		main_overlay.add(
			geometry=oaam.Shape.rectangle, x=x + 5, y=y + 5 + b * 8, width=30, height=6,
			color=(0, 255, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 200, 0))


def overlay_add_search_mode_icon(main_overlay, x, y):
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=x, y=y, width=40, height=40,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=x + 5, y=y + 5, width=30, height=30,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 255, 0))
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=x + 10, y=y + 5 + 1 * 8, width=15, height=6,
		color=(0, 255, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 0, 0))
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=x + 10, y=y + 5 + 2 * 8, width=15, height=6,
		color=(0, 255, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 0, 0))


class Recorder(Thread):
	def __init__(self, path_separator='->', type_separator='||'):
		Thread.__init__(self)
		core.path_separator = path_separator
		core.type_separator = type_separator
		self.main_overlay = oaam.Overlay(transparency=0.5)
		self.desktop = pywinauto.Desktop(backend='uia', allow_magic_lookup=False)
		self._is_running = False
		self.daemon = True
		self.event_list = []
		self.last_element_event = None
		self.start()

	def __find_unique_element_array_1d(self, wrapper_rectangle, elements):
		nb_y, nb_x, candidates = core.get_sorted_region(elements)
		for r_y in range(nb_y):
			for r_x in range(nb_x):
				try:
					r = candidates[r_y][r_x].rectangle()
				except IndexError:
					continue
				if r == wrapper_rectangle:
					xx, yy = r.left, r.mid_point()[1]
					previous_wrapper_path2 = None
					while xx > 0:
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

						previous_wrapper_path2 = wrapper_path2

						entry_list2 = core.get_entry_list(wrapper_path2)
						unique_candidate2, _ = core.find_element(self.desktop, entry_list2, window_candidates=[])

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
		nb_y, nb_x, candidates = core.get_sorted_region(elements)
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
		if self.event_list:
			if (type(mouse_event) == mouse.MoveEvent) and (len(self.event_list) > 0):
				if type(self.event_list[-1]) == mouse.MoveEvent:
					self.event_list = self.event_list[:-1]

			self.event_list.append(mouse_event)

	def __key_on(self, e):
		if (
				e.name == 'r' and e.event_type == 'up'
				and keyboard.key_to_scan_codes("alt")[0] in keyboard._pressed_events
				and keyboard.key_to_scan_codes("ctrl")[0] in keyboard._pressed_events):
			if not self.event_list:
				self.start_recording()
			else:
				self.stop_recording()
		elif (
				(e.name == 'q') and (e.event_type == 'up')
				and keyboard.key_to_scan_codes("alt")[0] in keyboard._pressed_events
				and keyboard.key_to_scan_codes("ctrl")[0] in keyboard._pressed_events):
			self.quit()
		elif (
				(e.name == 'F') and (e.event_type == 'up')
				and keyboard.key_to_scan_codes("shift")[0] in keyboard._pressed_events
				and keyboard.key_to_scan_codes("ctrl")[0] in keyboard._pressed_events):
			# remplacer l'icone play_icon par une icone Clipboard animée
			if self.last_element_event:
				x, y = win32api.GetCursorPos()
				l_e_e = self.last_element_event
				rx, ry = l_e_e.rectangle.mid_point()
				dx, dy = x - rx, y - ry
				overlay_add_play_icon(self.main_overlay, x, y)
				pyperclip.copy(
					'wrapper = r.find(u"' + escape_special_char(l_e_e.path) + '%(' + str(dx) + ',' + str(dy) + ')")\n')
				if self.event_list:
					self.event_list.append(FindEvent(path=l_e_e.path, dx=dx, dy=dy, time=time.time()))
		elif self.event_list:
			self.event_list.append(e)

	def run(self):
		keyboard.hook(self.__key_on)
		mouse.hook(self.__mouse_on)
		unique_candidate = None
		elements = []
		i = 0
		previous_wrapper_path = None
		unique_wrapper_path = None
		strategies = [core.Strategy.unique_path, core.Strategy.array_2D, core.Strategy.array_1D]
		i_strategy = 0
		self._is_running = True
		while self._is_running:
			try:
				self.main_overlay.clear_all()
				x, y = win32api.GetCursorPos()
				wrapper = self.desktop.from_point(x, y)
				if wrapper is None:
					continue
				wrapper_path = get_wrapper_path(wrapper)
				if not wrapper_path:
					continue
				if wrapper_path == previous_wrapper_path:
					if (unique_wrapper_path is None) or (strategies[i_strategy] == core.Strategy.array_2D):
						i_strategy = i_strategy + 1
						if i_strategy >= len(strategies):
							i_strategy = len(strategies) - 1
				else:
					i_strategy = 0
					previous_wrapper_path = wrapper_path
					entry_list = core.get_entry_list(wrapper_path)
					unique_candidate, elements = core.find_element(self.desktop, entry_list, window_candidates=[])
				strategy = strategies[i_strategy]
				unique_wrapper_path = None
				if strategy == core.Strategy.unique_path:
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
				if strategy == core.Strategy.array_1D:
					unique_array_1d = self.__find_unique_element_array_1d(wrapper.rectangle(), elements)
					if unique_array_1d is not None:
						unique_wrapper_path = wrapper_path + unique_array_1d
					else:
						strategy = core.Strategy.array_2D
				if strategy == core.Strategy.array_2D:
					unique_array_2d = self.__find_unique_element_array_2d(wrapper.rectangle(), elements)
					if unique_array_2d is not None:
						unique_wrapper_path = wrapper_path + unique_array_2d
				if unique_wrapper_path is not None:
					self.last_element_event = ElementEvent(strategy, wrapper.rectangle(), unique_wrapper_path)
					if self.event_list and unique_wrapper_path is not None:
						self.event_list.append(self.last_element_event)
				if self.event_list:
					overlay_add_record_icon(self.main_overlay, 10, 10)
				else:
					overlay_add_pause_icon(self.main_overlay, 10, 10)
				overlay_add_progress_icon(self.main_overlay, i, 60, 10)
				overlay_add_search_mode_icon(self.main_overlay, 110, 10)
				i = i + 1
				self.main_overlay.refresh()
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

	def start_recording(self):
		x, y = win32api.GetCursorPos()
		self.event_list = [mouse.MoveEvent(x, y, time.time())]
		overlay_add_record_icon(self.main_overlay, 10, 10)
		self.main_overlay.refresh()

	def stop_recording(self):
		if self.event_list:
			alt_ctrl = 0
			i = 0
			while alt_ctrl < 2:
				if type(self.event_list[i]) == keyboard.KeyboardEvent and self.event_list[i].name in {'alt', 'ctrl'}:
					alt_ctrl += 1
					self.event_list.pop(i)
				else:
					i += 1
			alt_ctrl = 0
			i = len(self.event_list) - 1
			while alt_ctrl < 2:
				if type(self.event_list[i]) == keyboard.KeyboardEvent and self.event_list[i].name in {'alt', 'ctrl'}:
					alt_ctrl += 1
				i -= 1
			events = list(self.event_list[0:i])
			self.event_list = []
			clean_events(events)
			process_events(events)
			return write_in_file(events)
		self.main_overlay.clear_all()
		overlay_add_pause_icon(self.main_overlay, 10, 10)
		self.main_overlay.refresh()
		return None

	def quit(self):
		print("Quit")
		self._is_running = False
