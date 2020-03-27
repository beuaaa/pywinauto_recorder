# -*- coding: utf-8 -*-

import sys
import os
import traceback
from ctypes.wintypes import tagPOINT
import time
import win32api
from threading import Thread
import pywinauto
import overlay_arrows_and_more as oaam
import keyboard
import mouse
import core
from collections import namedtuple

ElementEvent = namedtuple('ElementEvent', ['strategy', 'rectangle', 'path'])
SendKeysEvent = namedtuple('SendKeysEvent', ['line'])
MouseWheelEvent = namedtuple('MouseWheelEvent', ['delta'])
DragAndDropEvent = namedtuple('DragAndDropEvent', ['path', 'dx1', 'dy1', 'dx2', 'dy2'])
ClickEvent = namedtuple('ClickEvent', ['button', 'click_count', 'path', 'dx', 'dy', 'time'])
CommonPathEvent = namedtuple('CommonPathEvent', ['path'])

reload(sys)
sys.setdefaultencoding('utf-8')

event_list = []


def write_in_file(event_list_copy):
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
	while i < len(event_list_copy):
		e_i = event_list_copy[i]
		if type(e_i) is SendKeysEvent:
			if common_path:
				record_file.write("\t")
			record_file.write("send_keys(" + e_i.line + ")\n")
		elif type(e_i) is MouseWheelEvent:
			if common_path:
				record_file.write("\t")
			record_file.write("mouse_wheel(" + str(e_i.delta) + ")\n")
		elif type(e_i) is CommonPathEvent:
			record_file.write('\nwith Region("""' + e_i.path + '""") as r:\n')
			common_path = e_i.path
		elif type(e_i) is DragAndDropEvent:
			p, dx1, dy1, dx2, dy2 = e_i.path, str(e_i.dx1), str(e_i.dy1), str(e_i.dx2), str(e_i.dy2)
			if common_path:
				p = get_relative_path(common_path, p)
			record_file.write('\tr.drag_and_drop("""' + p + '%(' + dx1 + ',' + dy1 + ')%(' + dx2 + ',' + dy2 + ')""")\n')
		elif type(e_i) is ClickEvent:
			p, dx, dy = e_i.path, str(e_i.dx), str(e_i.dy)
			if common_path:
				p = get_relative_path(common_path, p)
			str_c = ['', '\tr.', '\tr.double_', '\tr.triple_']
			record_file.write(str_c[e_i.click_count] + e_i.button + '_click("""' + p + '%(' + dx + ',' + dy + ')""")\n')
		i = i + 1
	record_file.close()
	return record_file_name


def clean_events(event_list_copy):
	"""
	remove duplicate or useless events
	:param event_list_copy: the copy of recorded event list
	"""
	i = 0
	previous_event_type = None
	while i < len(event_list_copy):
		if type(event_list_copy[i]) is previous_event_type:
			if type(event_list_copy[i]) in [ElementEvent, mouse.MoveEvent]:
				del event_list_copy[i - 1]
			else:
				previous_event_type = type(event_list_copy[i])
				i = i + 1
		else:
			previous_event_type = type(event_list_copy[i])
			i = i + 1


def process_events(event_list_copy):
	i = 0
	while i < len(event_list_copy):
		if type(event_list_copy[i]) is keyboard.KeyboardEvent:
			process_keyboard_events(event_list_copy, i)
		elif type(event_list_copy[i]) is mouse.WheelEvent:
			process_wheel_events(event_list_copy, i)
		i = i + 1
	i = len(event_list_copy) - 1
	while i >= 0:
		if type(event_list_copy[i]) is mouse.ButtonEvent and event_list_copy[i].event_type == 'up':
			i = process_drag_and_drop_or_click_events(event_list_copy, i)
		i = i - 1

	common_path = None
	while i < len(event_list_copy):
		if type(event_list_copy[i]) in [DragAndDropEvent, ClickEvent]:
			common_path = process_common_path_events(event_list_copy, i, common_path)
		i = i + 1


def process_keyboard_events(event_list_copy, i):
	keyboard_events = [event_list_copy[i]]
	i0 = i + 1
	i_processed_events = []
	while i0 < len(event_list_copy):
		if type(event_list_copy[i0]) == keyboard.KeyboardEvent:
			keyboard_events.append(event_list_copy[i0])
			i_processed_events.append(i0)
			i0 = i0 + 1
		elif type(event_list_copy[i0]) == ElementEvent:
			i0 = i0 + 1
		else:
			break
	line = get_send_keys_strings(keyboard_events)
	for i_p_e in sorted(i_processed_events, reverse=True):
		del event_list_copy[i_p_e]
	if line:
		event_list_copy[i] = SendKeysEvent(line=line)


def process_wheel_events(event_list_copy, i):
	delta = event_list_copy[i].delta
	i_processed_events = []
	i0 = i + 1
	while i0 < len(event_list_copy):
		if type(event_list_copy[i0]) == mouse.WheelEvent:
			delta = delta + event_list_copy[i0].delta
			i_processed_events.append(i0)
			i0 = i0 + 1
		elif type(event_list_copy[i0]) in [ElementEvent, mouse.MoveEvent]:
			i0 = i0 + 1
		else:
			break
	for i_p_e in sorted(i_processed_events, reverse=True):
		del event_list_copy[i_p_e]
	event_list_copy[i] = MouseWheelEvent(delta=delta)


def process_drag_and_drop_or_click_events(event_list_copy, i):
	i0 = i - 1
	move_event_end = None
	drag_and_drop = False
	click_count = 0
	while i0 >= 0:
		if type(event_list_copy[i0]) == mouse.MoveEvent:
			if move_event_end:
				if event_list_copy[i0].x != move_event_end.x or event_list_copy[i0].y != move_event_end.y:
					drag_and_drop = True
			else:
				move_event_end = event_list_copy[i0]
		elif type(event_list_copy[i0]) == mouse.ButtonEvent and event_list_copy[i0].event_type in ['down', 'double']:
			click_count = click_count + 1
			if event_list_copy[i0].event_type == 'down' or click_count == 3:
				i1 = i0
				break
		i0 = i0 - 1
	element_event_before_click = None
	while i0 >= 0:
		if type(event_list_copy[i0]) == ElementEvent:
			element_event_before_click = event_list_copy[i0]
			break
		i0 = i0 - 1
	if drag_and_drop:
		move_event_start = None
		while i0 >= 0:
			if type(event_list_copy[i0]) == mouse.MoveEvent:
				move_event_start = event_list_copy[i0]
				break
			i0 = i0 - 1
		rx, ry = element_event_before_click.rectangle.mid_point()
		dx1, dy1 = move_event_start.x - rx, move_event_start.y - ry
		dx2, dy2 = move_event_end.x - rx, move_event_end.y - ry
		event_list_copy[i] = DragAndDropEvent(path=element_event_before_click.path, dx1=dx1, dy1=dy1, dx2=dx2, dy2=dy2)
	else:
		while i0 >= 0:
			if move_event_end:
				break
			if type(event_list_copy[i0]) == mouse.MoveEvent:
				move_event_end = event_list_copy[i0]
				break
			i0 = i0 - 1
		up_event = event_list_copy[i]
		rx, ry = element_event_before_click.rectangle.mid_point()
		dx, dy = move_event_end.x - rx, move_event_end.y - ry
		event_list_copy[i] = ClickEvent(button=up_event.button, click_count=click_count,
										path=element_event_before_click.path, dx=dx, dy=dy, time=up_event.time)
	i_processed_events = []
	i0 = i - 1
	while i0 >= i1:
		if type(event_list_copy[i0]) in [mouse.ButtonEvent, mouse.MoveEvent, ElementEvent]:
			i_processed_events.append(i0)
		i0 = i0 - 1
	for i_p_e in sorted(i_processed_events, reverse=True):
		del event_list_copy[i_p_e]
		i = i - 1
	return i


def get_relative_path(common_path, path):
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


def process_common_path_events(event_list_copy, i, common_path):
	path_i = event_list_copy[i].path
	i0 = i + 1
	new_common_path = ''
	while i0 < len(event_list_copy):
		e = event_list_copy[i0]
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
		event_list_copy.insert(i, CommonPathEvent(path=new_common_path))
		return new_common_path
	return common_path


def get_wrapper_path(wrapper):
	try:
		path = ''
		wrapper_top_level_parent = wrapper.top_level_parent()
		while wrapper != wrapper_top_level_parent:
			if not wrapper:
				break
			path = core.path_separator + wrapper.window_text() + core.type_separator + wrapper.element_info.control_type + path
			wrapper = wrapper.parent()

		return wrapper.window_text() + core.type_separator + wrapper.element_info.control_type + path
	except Exception:
		return ''


def get_type_strings(keyboard_events, allow_backspace=True):
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
	return ''.join(format(code) for code in get_type_strings(keyboard_events))


def key_on(e):
	global recorder
	global event_list

	if e.name == 'r' and e.event_type == 'up' and 56 in keyboard._pressed_events:
		if not event_list:
			recorder.start_recording()
		else:
			recorder.stop_recording()
	elif e.name == 'q' and e.event_type == 'up' and 56 in keyboard._pressed_events:
		recorder.quit()
	elif event_list:
		event_list.append(e)


def mouse_on(mouse_event):
	global event_list
	if event_list:
		if (type(mouse_event) == mouse.MoveEvent) and (len(event_list) > 0):
			if type(event_list[-1]) == mouse.MoveEvent:
				event_list = event_list[:-1]

		event_list.append(mouse_event)


class Recorder(Thread):
	def __init__(self, path_separator='->', type_separator='||'):
		Thread.__init__(self)
		core.path_separator = path_separator
		core.type_separator = type_separator
		self.main_overlay = oaam.Overlay(transparency=0.5)
		self.desktop = pywinauto.Desktop(backend='uia', allow_magic_lookup=False)
		self._is_running = False
		self.daemon = True
		self.start()

	def main_overlay_add_record_icon(self):
		self.main_overlay.add(
			geometry=oaam.Shape.rectangle, x=10, y=10, width=40, height=40,
			color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
		self.main_overlay.add(
			geometry=oaam.Shape.ellipse, x=15, y=15, width=29, height=29,
			color=(255, 90, 90), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 0, 0))

	def main_overlay_add_pause_icon(self):
		self.main_overlay.add(
			geometry=oaam.Shape.rectangle, x=10, y=10, width=40, height=40,
			color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
		self.main_overlay.add(
			geometry=oaam.Shape.rectangle, x=15, y=15, width=12, height=30,
			color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 0, 0))
		self.main_overlay.add(
			geometry=oaam.Shape.rectangle, x=32, y=15, width=12, height=30,
			color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 0, 0))

	def main_overlay_add_progress_icon(self, i):
		self.main_overlay.add(
			geometry=oaam.Shape.rectangle, x=60, y=10, width=40, height=40,
			color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
		for b in range(i % 5):
			self.main_overlay.add(
				geometry=oaam.Shape.rectangle, x=65, y=15 + b * 8, width=30, height=6,
				color=(0, 255, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 200, 0))

	def main_overlay_add_search_mode_icon(self):
		self.main_overlay.add(
			geometry=oaam.Shape.rectangle, x=110, y=10, width=40, height=40,
			color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
		self.main_overlay.add(
			geometry=oaam.Shape.rectangle, x=115, y=15, width=30, height=30,
			color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 255, 0))
		self.main_overlay.add(
			geometry=oaam.Shape.rectangle, x=120, y=15 + 1 * 8, width=15, height=6,
			color=(0, 255, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 0, 0))
		self.main_overlay.add(
			geometry=oaam.Shape.rectangle, x=120, y=15 + 2 * 8, width=15, height=6,
			color=(0, 255, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 0, 0))

	def find_unique_element_array_1D(self, element_info, elements):
		nb_y, nb_x, candidates = core.get_sorted_region(elements)
		for r_y in range(nb_y):
			for r_x in range(nb_x):
				try:
					r = candidates[r_y][r_x].rectangle()
				except IndexError:
					continue
				if r == element_info.rectangle:
					xx, yy = r.left, r.mid_point()[1]
					previous_element_path2 = None
					while xx > 0:
						xx = xx - 9
						element_from_point2 = pywinauto.uia_defines.IUIA().iuia.ElementFromPoint(tagPOINT(xx, yy))
						element_info2 = pywinauto.uia_element_info.UIAElementInfo(element_from_point2)
						if element_info2.control_type != "Text":
							continue
						if element_info2.rectangle.height() > element_info.rectangle.height() * 2:
							continue
						wrapper2 = pywinauto.controls.uiawrapper.UIAWrapper(element_info2)
						if wrapper2 is None:
							continue
						element_path2 = get_wrapper_path(wrapper2)
						if not element_path2:
							continue
						if element_path2 == previous_element_path2:
							continue

						previous_element_path2 = element_path2

						entry_list2 = core.get_entry_list(element_path2)
						unique_candidate2, _ = core.find_element(self.desktop, entry_list2, window_candidates=[])

						if unique_candidate2 is not None:
							r = element_info2.rectangle
							self.main_overlay.add(
								geometry=oaam.Shape.rectangle, x=r.left, y=r.top,
								width=r.width(), height=r.height(),
								thickness=1, color=(0, 0, 255), brush=oaam.Brush.solid,
								brush_color=(0, 0, 255))
							r = element_info.rectangle
							self.main_overlay.add(
								geometry=oaam.Shape.rectangle, x=r.left, y=r.top,
								width=r.width(), height=r.height(),
								thickness=1, color=(255, 0, 0), brush=oaam.Brush.solid,
								brush_color=(255, 200, 0))
							return '#[' + element_path2 + ',' + str(r_x) + ']'
					else:
						return None
		return None

	def find_unique_element_array_2D(self, element_info, elements):
		nb_y, nb_x, candidates = core.get_sorted_region(elements)
		unique_array_2D = ''
		for r_y in range(nb_y):
			for r_x in range(nb_x):
				try:
					r = candidates[r_y][r_x].rectangle()
				except IndexError:
					continue
				if r == element_info.rectangle:
					color = (255, 200, 0)
					unique_array_2D = '#[' + str(r_y) + ',' + str(r_x) + ']'
				else:
					color = (255, 0, 0)
				self.main_overlay.add(
					geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(),
					height=r.height(),
					thickness=1, color=(255, 0, 0), brush=oaam.Brush.solid, brush_color=color)
		return unique_array_2D

	def run(self):
		global event_list
		keyboard.hook(key_on)
		mouse.hook(mouse_on)
		unique_candidate = None
		elements = []
		i = 0
		previous_element_path = None
		unique_element_path = None
		strategies = [core.Strategy.unique_path, core.Strategy.array_2D, core.Strategy.array_1D]
		i_strategy = 0
		self._is_running = True
		while self._is_running:
			try:
				self.main_overlay.clear_all()
				x, y = win32api.GetCursorPos()
				element_from_point = pywinauto.uia_defines.IUIA().iuia.ElementFromPoint(tagPOINT(x, y))
				element_info = pywinauto.uia_element_info.UIAElementInfo(element_from_point)
				wrapper = pywinauto.controls.uiawrapper.UIAWrapper(element_info)
				if wrapper is None:
					continue
				element_path = get_wrapper_path(wrapper)
				if not element_path:
					continue
				if element_path == previous_element_path:
					if (unique_element_path is None) or (strategies[i_strategy] == core.Strategy.array_2D):
						i_strategy = i_strategy + 1
						if i_strategy >= len(strategies):
							i_strategy = len(strategies) - 1
				else:
					i_strategy = 0
					previous_element_path = element_path
					entry_list = core.get_entry_list(element_path)
					unique_candidate, elements = core.find_element(self.desktop, entry_list, window_candidates=[])
				strategy = strategies[i_strategy]
				unique_element_path = None
				if strategy == core.Strategy.unique_path:
					if unique_candidate is not None:
						unique_element_path = get_wrapper_path(unique_candidate)
						r = element_info.rectangle
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
					unique_array_1D = self.find_unique_element_array_1D(element_info, elements)
					if unique_array_1D is not None:
						unique_element_path = element_path + unique_array_1D
					else:
						strategy = core.Strategy.array_2D
				if strategy == core.Strategy.array_2D:
					unique_array_2D = self.find_unique_element_array_2D(element_info, elements)
					if unique_array_2D is not None:
						unique_element_path = element_path + unique_array_2D
				if event_list and unique_element_path is not None:
					event_list.append(ElementEvent(strategy, element_info.rectangle, unique_element_path))
				if event_list:
					self.main_overlay_add_record_icon()
				else:
					self.main_overlay_add_pause_icon()
				self.main_overlay_add_progress_icon(i)
				self.main_overlay_add_search_mode_icon()
				i = i + 1
				self.main_overlay.refresh()
				time.sleep(0.005)  # main_overlay.clear_all() doit attendre la fin de main_overlay.refresh()
			except Exception as e:
				exc_type, exc_value, exc_traceback = sys.exc_info()
				print(repr(traceback.format_exception(exc_type, exc_value, exc_traceback)))
		self.main_overlay.clear_all()
		self.main_overlay.refresh()
		if event_list:
			recorder.stop_recording()
		mouse.unhook_all()
		keyboard.unhook_all()
		print("Run end")

	def start_recording(self):
		global event_list
		x, y = win32api.GetCursorPos()
		event_list.append(mouse.MoveEvent(x, y, time.time()))
		self.main_overlay_add_record_icon()
		self.main_overlay.refresh()

	def stop_recording(self):
		global event_list
		if event_list:
			event_list_copy = list(event_list[0:-4])
			event_list = []
			clean_events(event_list_copy)
			process_events(event_list_copy)
			return write_in_file(event_list_copy)
		self.main_overlay.clear_all()
		self.main_overlay_add_pause_icon()
		self.main_overlay.refresh()
		return None

	def quit(self):
		print("Quit")
		self._is_running = False

