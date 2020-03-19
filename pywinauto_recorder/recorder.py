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
main_overlay = None
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


def main_overlay_add_record_icon():
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=10, y=10, width=40, height=40,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
	main_overlay.add(
		geometry=oaam.Shape.ellipse, x=15, y=15, width=29, height=29,
		color=(255, 90, 90), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 0, 0))


def main_overlay_add_pause_icon():
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=10, y=10, width=40, height=40,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=15, y=15, width=12, height=30,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 0, 0))
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=32, y=15, width=12, height=30,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 0, 0))


def main_overlay_add_progress_icon(i):
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=60, y=10, width=40, height=40,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
	for b in range(i % 5):
		main_overlay.add(
			geometry=oaam.Shape.rectangle, x=65, y=15 + b * 8, width=30, height=6,
			color=(0, 255, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 200, 0))


def main_overlay_add_search_mode_icon():
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=110, y=10, width=40, height=40,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=115, y=15, width=30, height=30,
		color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 255, 0))
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=120, y=15 + 1 * 8, width=15, height=6,
		color=(0, 255, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 0, 0))
	main_overlay.add(
		geometry=oaam.Shape.rectangle, x=120, y=15 + 2 * 8, width=15, height=6,
		color=(0, 255, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 0, 0))


def get_element_path(w):
	try:
		path = ''
		wrapper_top_level_parent = w.top_level_parent()
		while w != wrapper_top_level_parent:
			if not w:
				break
			if w.window_text() != '':
				path = "->" + w.window_text() + '::' + w.element_info.control_type + path
			else:
				path = "->" + '::' + w.element_info.control_type + path
			w = w.parent()

		return w.window_text() + '::' + w.element_info.control_type + path
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
	global main_overlay
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


def find_comon_path(element_event_before_click, i0, i1):
	while i0 > i1:
		if type(event_list[i0]) == ElementEvent:
			element_paths = event_list[i0].path.split('->')
			element_paths = element_paths[:-1]
			unique_element_paths = element_event_before_click.path.split('->')
			unique_element_paths = unique_element_paths[:-1]
			n = 0
			try:
				while element_paths[n] == unique_element_paths[n]:
					n = n + 1
			except IndexError:
				common_path = "->".join(element_paths[0:n])
				return common_path
			common_path = "->".join(element_paths[0:n])
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
			common_path = find_comon_path(element_event_before_click, i0, i1)
			break
		i0 = i0 + 1

	if old_common_path != common_path:
		record_file.write('in_region("""' + common_path + '""")\n')
	old_common_path = common_path
	if common_path:
		click_path = element_event_before_click.path[len(common_path) + 2:]
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
	def __init__(self, **parameters):
		Thread.__init__(self)
		self._is_running = False
		self.daemon = True
		self.start()

	def run(self):
		global main_overlay
		global record_file
		global event_list

		main_overlay = oaam.Overlay(transparency=100)
		keyboard.hook(key_on)
		mouse.hook(mouse_on)
		unique_candidate = None
		elements = []
		desktop = pywinauto.Desktop(backend='uia', allow_magic_lookup=False)
		i = 0
		previous_element_path = None
		previous_element_path_found = False
		strategies = [core.Strategy.unique_path, core.Strategy.array_2D]
		i_strategy = 0
		self._is_running = True
		while self._is_running:
			try:
				main_overlay.clear_all()

				x, y = win32api.GetCursorPos()
				element_from_point = pywinauto.uia_defines.IUIA().iuia.ElementFromPoint(tagPOINT(x, y))
				element_info = pywinauto.uia_element_info.UIAElementInfo(element_from_point)
				wrapper = pywinauto.controls.uiawrapper.UIAWrapper(element_info)
				if wrapper is None:
					continue
				element_path = get_element_path(wrapper)
				if not element_path:
					continue
				if element_path == previous_element_path:
					if not previous_element_path_found:
						i_strategy = i_strategy + 1
						if i_strategy >= len(strategies):
							i_strategy = len(strategies) - 1
				else:
					i_strategy = 0
					previous_element_path = element_path
					entry_list = (element_path.decode('utf-8')).split("->")
					unique_candidate, elements = core.find_element(desktop, entry_list, window_candidates=[])

				strategy = strategies[i_strategy]
				previous_element_path_found = False
				if strategy == core.Strategy.unique_path:
					if unique_candidate is not None:
						unique_element_path = get_element_path(unique_candidate)
						unique_element_r = unique_candidate.rectangle()
						# unique_candidate.draw_outline(colour='green', thickness=2)
						r = unique_candidate.rectangle()
						main_overlay.add(
							geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(), height=r.height(),
							thickness=1, color=(0, 128, 0), brush=oaam.Brush.solid, brush_color=(0, 255, 0))
						previous_element_path_found = True
					else:
						for e in elements:
							r = e.rectangle()
							main_overlay.add(
								geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(), height=r.height(),
								thickness=1, color=(0, 128, 0), brush=oaam.Brush.solid, brush_color=(255, 0, 0))
				elif strategy == core.Strategy.array_2D:
					nb_y, nb_x, candidates = core.get_sorted_region(elements)
					for r_y in range(nb_y):
						for r_x in range(nb_x):
							try:
								r = candidates[r_y][r_x].rectangle()
							except IndexError:
								continue
							if r == element_info.rectangle:
								color = (255, 200, 0)
								unique_element_path = element_path + '#[' + str(r_y) + ',' + str(r_x) + ']'
								unique_element_r = r
								previous_element_path_found = True
							else:
								color = (255, 0, 0)
							main_overlay.add(
								geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(),
								height=r.height(),
								thickness=1, color=(255, 0, 0), brush=oaam.Brush.solid, brush_color=color)
				elif strategy == core.Strategy.array_1D:
					nb_y, nb_x, candidates = core.get_sorted_region(elements)
					for r_y in range(nb_y):
						for r_x in range(nb_x):
							try:
								r = candidates[r_y][r_x].rectangle()
							except IndexError:
								continue
							if r == element_info.rectangle:
								color = (255, 200, 0)
								unique_element_path = element_path + '#[' + str(r_y) + ',' + str(r_x) + ']'
								unique_element_r = r
								previous_element_path_found = True
							else:
								color = (255, 0, 0)
							main_overlay.add(
								geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(),
								height=r.height(),
								thickness=1, color=(255, 0, 0), brush=oaam.Brush.solid, brush_color=color)

				if record_file and (strategy is not core.Strategy.failed):
					if (len(event_list) > 0) and (type(event_list[-1]) == ElementEvent):
						if event_list[-1].path != unique_element_path:
							event_list.append(ElementEvent(strategy, unique_element_r, unique_element_path))
					else:
						event_list.append(ElementEvent(strategy, unique_element_r, unique_element_path))

				if record_file:
					main_overlay_add_record_icon()
				else:
					main_overlay_add_pause_icon()

				main_overlay_add_progress_icon(i)
				main_overlay_add_search_mode_icon()

				i = i + 1
				main_overlay.refresh()
				time.sleep(0.005)  # main_overlay.clear_all() doit attendre la fin de main_overlay.refresh()
			except Exception as e:
				exc_type, exc_value, exc_traceback = sys.exc_info()
				print(repr(traceback.format_exception(exc_type, exc_value, exc_traceback)))
		main_overlay.clear_all()
		main_overlay.refresh()
		if record_file:
			recorder.stop_recording()
		keyboard.unhook_all()
		mouse.unhook_all()
		del desktop
		del main_overlay

	def start_recording(self):
		global main_overlay
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
		main_overlay_add_record_icon()
		main_overlay.refresh()
		return record_file_name

	def stop_recording(self):
		global event_list
		global main_overlay
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
		main_overlay.clear_all()
		main_overlay_add_pause_icon()
		main_overlay.refresh()

	def quit(self):
		print("Quit")
		self._is_running = False


if __name__ == '__main__':
	global recorder
	recorder = Recorder()
	while recorder.is_alive():
		time.sleep(1.0)
	del recorder
