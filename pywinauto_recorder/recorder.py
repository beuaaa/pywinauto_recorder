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

reload(sys)
sys.setdefaultencoding('utf-8')

record_file = None
event_list = []


def get_double_click_time():
	""" Gets the Windows double click time in s """
	from ctypes import windll
	return windll.user32.GetDoubleClickTime() / 1000.0


class ElementEvent(object):
	def __init__(self, strategy, rectangle, path):
		self.strategy = strategy
		self.rectangle = rectangle
		self.path = path


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
	global record_file

	if e.name == 'r' and e.event_type == 'up' and 56 in keyboard._pressed_events:
		if record_file is None:
			recorder.start_recording()
		else:
			recorder.stop_recording()
	elif e.name == 'q' and e.event_type == 'up' and 56 in keyboard._pressed_events:
		recorder.quit()
	else:
		event_list.append(e)


old_common_path = ''


def find_common_path(element_event_before_click, i0, i1):
	while i0 > i1:
		if type(event_list[i0]) == ElementEvent:
			element_paths = core.get_entry_list(event_list[i0].path)
			_, _, y_x, _ = core.get_entry(element_paths[-1])
			if (y_x is not None) and not core.is_int(y_x[0]):
				element_paths = core.get_entry_list(y_x[0])[:-1]
			else:
				element_paths = element_paths[:-1]

			unique_element_paths = core.get_entry_list(element_event_before_click.path)
			unique_element_paths = unique_element_paths[:-1]
			n = 0
			try:
				while element_paths[n] == unique_element_paths[n]:
					n = n + 1
			except IndexError:
				common_path = core.path_separator.join(element_paths[0:n])
				return common_path
			common_path = core.path_separator.join(element_paths[0:n])
			return common_path
		i0 = i0 - 1
	return ''


# TODO: % is relative to the center of the element, A is absolute, P is proportional, ...
def record_click(i):
	global record_file
	global old_common_path
	t1 = event_list[i].time
	i1 = i

	i0 = i - 1
	while type(event_list[i0]) != mouse.MoveEvent:
		i0 = i0 - 1
	move_event_before_click = event_list[i0]

	i0 = i - 1
	while type(event_list[i0]) != ElementEvent:
		i0 = i0 - 1
	element_event_before_click = event_list[i0]

	i0 = i + 1
	double_click = False
	while (i0 < len(event_list)) and \
			((type(event_list[i0]) != mouse.ButtonEvent) or ((event_list[i0].time - t1) < get_double_click_time())):
		if (type(event_list[i0]) == mouse.ButtonEvent) and (event_list[i0].event_type == 'up'):
			double_click = True
			i1 = i0
			break
		i0 = i0 + 1
	i0 = i0 + 1
	triple_click = False
	while (i0 < len(event_list)) and \
			((type(event_list[i0]) != mouse.ButtonEvent) or (
					(event_list[i0].time - t1) < 2.0 * get_double_click_time())):
		if (type(event_list[i0]) == mouse.ButtonEvent) and (event_list[i0].event_type == 'up'):
			triple_click = True
			i1 = i0
			break
		i0 = i0 + 1

	common_path = ''
	i0 = i1 + 1
	while i0 < len(event_list):
		if (type(event_list[i0]) == mouse.ButtonEvent) and (event_list[i0].event_type == 'up'):
			common_path = find_common_path(element_event_before_click, i0, i1)
			break
		i0 = i0 + 1

	if old_common_path != common_path:
		record_file.write('in_region("""' + common_path + '""")\n')
	old_common_path = common_path
	if common_path:
		click_path = element_event_before_click.path[len(common_path) + len(core.path_separator):]
		entry_list = core.get_entry_list(click_path)
		str_name, str_type, y_x, dx_dy = core.get_entry(entry_list[-1])
		if (y_x is not None) and not core.is_int(y_x[0]):
			y_x[0] = y_x[0][len(common_path) + 2:]
			click_path = core.path_separator.join(entry_list[:-1]) + core.path_separator + str_name
			click_path = click_path + core.type_separator + str_type + "#[" + y_x[0] + "," + str(y_x[1]) + "]"
			if dx_dy is not None:
				click_path = click_path + "%(" + str(dx_dy[0]) + "," + str(dx_dy[1]) + ")"
	else:
		click_path = element_event_before_click.path

	if triple_click:
		record_file.write('triple_')
	elif double_click:
		record_file.write('double_')
	rx, ry = element_event_before_click.rectangle.mid_point()
	dx, dy = move_event_before_click.x - rx, move_event_before_click.y - ry
	record_file.write(event_list[i].button + "_")
	record_file.write('click("""' + click_path + '%(' + str(dx) + ',' + str(dy) + ')""")\n')
	return i1


def record_drag(i):
	global event_list
	global record_file
	i0 = i
	while (type(event_list[i0]) != mouse.ButtonEvent) or (event_list[i0].event_type != 'up'):
		i0 = i0 - 1
	move_event_end = event_list[i0]
	while (type(event_list[i0]) != mouse.ButtonEvent) or (event_list[i0].event_type != 'down'):
		i0 = i0 - 1
	while type(event_list[i0]) != ElementEvent:
		i0 = i0 - 1
	unique_element = event_list[i0]
	while type(event_list[i0]) != mouse.MoveEvent:
		i0 = i0 - 1
	move_event_start = event_list[i0]

	mouse_down_unique_rectangle = unique_element.rectangle
	x, y = move_event_start.x, move_event_start.y
	rx, ry = mouse_down_unique_rectangle.mid_point()
	dx, dy = x - rx, y - ry
	record_file.write('drag_and_drop("""' + unique_element.path + '%(' + str(dx) + ',' + str(dy))
	x2, y2 = move_event_end.x, move_event_end.y
	dx, dy = x2 - rx, y2 - ry
	record_file.write(')%(' + str(dx) + ',' + str(dy) + ')""")\n')


def record_wheel(i):
	global event_list
	global record_file
	record_file.write('mouse_wheel(' + str(event_list[i].delta) + ')\n')


def mouse_on(mouse_event):
	global record_file
	global event_list
	if record_file:
		if (type(mouse_event) == mouse.MoveEvent) and (len(event_list) > 0):
			if type(event_list[-1]) == mouse.MoveEvent:
				event_list = event_list[:-1]

		event_list.append(mouse_event)


mouse_down_time = 0
mouse_down_pos = (0, 0)


def record_mouse(i):
	global mouse_down_time
	global mouse_down_pos
	global event_list
	global record_file

	mouse_event = event_list[i]
	mouse_on_click = False
	mouse_on_drag = False
	mouse_on_wheel = False

	if type(mouse_event) == mouse.MoveEvent:
		a = 0  # TODO: recording with timings
	elif type(mouse_event) == mouse.ButtonEvent:
		if mouse_event.event_type == 'down':
			mouse_down_time = mouse_event.time
			mouse_down_pos = mouse.get_position()
		if mouse_event.event_type == 'up':
			if (mouse_event.time - mouse_down_time) < 0.2:
				mouse_on_click = True
			else:
				if mouse_down_pos != mouse.get_position():
					mouse_on_drag = True
				else:
					mouse_on_click = True
	elif type(mouse_event) == mouse.WheelEvent:
		mouse_on_wheel = True

	if (record_file is not None) and (mouse_on_click or mouse_on_drag or mouse_on_wheel):
		if mouse_on_click:
			i = record_click(i)
		if mouse_on_drag:
			record_drag(i)
		if mouse_on_wheel:
			record_wheel(i)

	return i


class Recorder(Thread):
	def __init__(self, path_separator='->', type_separator='||'):
		Thread.__init__(self)
		core.path_separator = path_separator
		core.type_separator = type_separator
		self.desktop = None
		self.main_overlay = None
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
		global record_file
		global event_list

		self.main_overlay = oaam.Overlay(transparency=100)
		self.desktop = pywinauto.Desktop(backend='uia', allow_magic_lookup=False)
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

				if record_file and unique_element_path is not None:
					if (len(event_list) > 0) and (type(event_list[-1]) == ElementEvent):
						if event_list[-1].path != unique_element_path:
							event_list.append(ElementEvent(strategy, element_info.rectangle, unique_element_path))
					else:
						event_list.append(ElementEvent(strategy, element_info.rectangle, unique_element_path))

				if record_file:
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
		if record_file:
			recorder.stop_recording()
		mouse.unhook_all()
		keyboard.unhook_all()
		print("Run end")

	def start_recording(self):
		global record_file

		new_path = 'Record files'
		if not os.path.exists(new_path):
			os.makedirs(new_path)
		record_file_name = './Record files/recorded ' + time.asctime() + '.py'
		record_file_name = record_file_name.replace(':', '_')
		print('Recording in file: ' + record_file_name)
		record_file = open(record_file_name, "w")
		record_file.write("# coding: utf-8\n")
		record_file.write("import sys, os\n")
		record_file.write("sys.path.append(os.path.realpath(os.path.dirname(__file__)+'/..'))\n")
		record_file.write("from player import *\n")
		record_file.write('send_keys("{LWIN down}""{DOWN}""{DOWN}""{LWIN up}")\n')
		record_file.write('time.sleep(0.5)\n')
		self.main_overlay_add_record_icon()
		self.main_overlay.refresh()
		return record_file_name

	def stop_recording(self):
		global event_list
		global record_file
		if event_list:
			event_list = event_list[:-1]
			event_list = event_list[:-1]
			i = 3
			while i < len(event_list):
				keyboard_events = []
				while i < len(event_list):
					if type(event_list[i]) == keyboard.KeyboardEvent:
						keyboard_events.append(event_list[i])
						i = i + 1
					elif type(event_list[i]) == ElementEvent:
						i = i + 1
					else:
						break
				line = get_send_keys_strings(keyboard_events)
				if line:
					record_file.write("send_keys(" + line + ")\n")
				while i < len(event_list):
					if type(event_list[i]) not in [mouse.ButtonEvent, mouse.WheelEvent, mouse.MoveEvent]:
						break
					elif type(event_list[i]) in [mouse.ButtonEvent, mouse.WheelEvent]:
						i = record_mouse(i)
					i = i + 1
		record_file.close()
		record_file = None
		self.main_overlay.clear_all()
		self.main_overlay_add_pause_icon()
		self.main_overlay.refresh()

	def quit(self):
		print("Quit")
		self._is_running = False


if __name__ == '__main__':
	global recorder
	recorder = Recorder()
	while recorder.is_alive():
		time.sleep(1.0)
	del recorder
	print("Exit")
	exit(0)
