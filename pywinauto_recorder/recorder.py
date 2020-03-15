# -*- coding: utf-8 -*-

import os
from ctypes.wintypes import tagPOINT
import time
import win32api
from threading import Thread
import pywinauto
import overlay_arrows_and_more as oaam
import keyboard
import mouse
from recorder_fn import find_element

recorder = None
record_file = None
main_overlay = None
event_list = []


def get_double_click_time():
	""" Gets the Windows double click time in s """
	from ctypes import windll
	return windll.user32.GetDoubleClickTime() / 1000.0


class ElementEvent(object):
	def __init__(self, rectangle, path):
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


# TODO: % is relative to the center of the element, A is absolute, P is proportional, ...
def record_click(i):
	global record_file
	global old_common_path
	t1 = event_list[i].time
	i1 = i

	i0 = i - 1
	while type(event_list[i0]) != mouse.MoveEvent:
		i0 = i0 - 1
	up_event = event_list[i0]
	while type(event_list[i0]) != ElementEvent:
		i0 = i0 - 1
	unique_element = event_list[i0]

	# ne pas aller vers le passé mais vers le futur
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
	while i0 < len(event_list):
		if type(event_list[i0]) == ElementEvent:
			element_paths = event_list[i0].path.split('->')
			element_paths = element_paths[:-1]
			unique_element_paths = unique_element.path.split('->')
			unique_element_paths = unique_element_paths[:-1]
			n = 0
			try:
				while element_paths[n] == unique_element_paths[n]:
					n = n + 1
			except IndexError:
				common_path = "->".join(element_paths[0:n])
				break
			common_path = "->".join(element_paths[0:n])
			break
		i0 = i0 + 1

	if old_common_path != common_path:
		record_file.write('in_region("""' + common_path + '""")\n')
	old_common_path = common_path
	if common_path:
		click_path = unique_element.path[len(common_path) + 2:]
	else:
		click_path = unique_element.path

	if triple_click:
		record_file.write('triple_')
	elif double_click:
		record_file.write('double_')
	rx, ry = unique_element.rectangle.mid_point()
	dx, dy = up_event.x - rx, up_event.y - ry
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
		pywinauto_desktop = pywinauto.Desktop(backend='uia', allow_magic_lookup=False)
		i = 0
		self._is_running = True
		while self._is_running:
			try:
				main_overlay.clear_all()

				x, y = win32api.GetCursorPos()
				elem = pywinauto.uia_defines.IUIA().iuia.ElementFromPoint(tagPOINT(x, y))
				element = pywinauto.uia_element_info.UIAElementInfo(elem)
				wrapper = pywinauto.controls.uiawrapper.UIAWrapper(element)
				if wrapper is None:
					continue
				element_path = get_element_path(wrapper)
				if not element_path:
					continue

				entry_list = (element_path.decode('utf-8')).split("->")
				unique_candidate, elements = find_element(pywinauto_desktop, entry_list, window_candidates=[])
				if unique_candidate is not None:
					unique_element_path = get_element_path(unique_candidate)
					# unique_candidate.draw_outline(colour='green', thickness=2)
					r = unique_candidate.rectangle()
					if record_file:
						if (len(event_list) > 0) and (type(event_list[-1]) == ElementEvent):
							if event_list[-1].path != unique_element_path:
								event_list.append(ElementEvent(r, unique_element_path))
						else:
							event_list.append(ElementEvent(r, unique_element_path))
					main_overlay.add(
						geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(), height=r.height(),
						thickness=1, color=(0, 128, 0), brush=oaam.Brush.solid, brush_color=(0, 255, 0))

					for e in elements:
						r = e.rectangle()
						main_overlay.add(
							geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(), height=r.height(),
							thickness=1, color=(255, 0, 0), brush=oaam.Brush.solid, brush_color=(255, 0, 0))

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
				print('Exception raised in main loop: \n')
				print(type(e))
				print(e.args)
				print(e)
		main_overlay.clear_all()
		main_overlay.refresh()
		if record_file:
			recorder.stop_recording()
		keyboard.unhook_all()
		mouse.unhook_all()
		del pywinauto_desktop
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
		record_file.write("from recorder_fn import *\n")
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
