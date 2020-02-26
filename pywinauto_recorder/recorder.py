# -*- coding: utf-8 -*-

import sys, os
reload(sys)
sys.setdefaultencoding('utf-8')
sys.path.append(os.path.realpath(__file__))
from recorder_fn import *

from ctypes.wintypes import tagPOINT
import time, math
import win32api, win32con, win32gui, win32ui

import pywinauto
import overlay_arrows_and_more as oaam

import keyboard
import mouse
import traceback
from threading import Thread

record_file = None
unique_rectangle = None
unique_element_path = None
main_overlay = oaam.Overlay(transparency=100)

def main_overlay_add_record_icon():
	main_overlay.add(geometry=oaam.Shape.rectangle, x=10, y=10, width=40, height=40, color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid,  brush_color=(255, 255, 254))
	main_overlay.add(geometry=oaam.Shape.ellipse, x=15, y=15, width=29, height=29, color=(255, 90, 90), thickness=1,
					 brush=oaam.Brush.solid, brush_color=(255, 0, 0))


def main_overlay_add_pause_icon():
	main_overlay.add(geometry=oaam.Shape.rectangle, x=10, y=10, width=40, height=40, color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid,  brush_color=(255, 255, 254))
	main_overlay.add(geometry=oaam.Shape.rectangle, x=15, y=15, width=12, height=30, color=(0, 0, 0), thickness=1,
					 brush=oaam.Brush.solid, brush_color=(0, 0, 0))
	main_overlay.add(geometry=oaam.Shape.rectangle, x=32, y=15, width=12, height=30, color=(0, 0, 0), thickness=1,
					 brush=oaam.Brush.solid, brush_color=(0, 0, 0))


def get_element_path(w):
	try:
		textBW=''
		wrapper_top_level_parent = w.top_level_parent()
		while w != wrapper_top_level_parent:
			if w==None:
				break
			if w.window_text()!='':
				textBW = "->" + w.window_text() + '::' + w.element_info.control_type + textBW
			else:
				textBW = "->" + '::' + w.element_info.control_type + textBW
			w = w.parent()

		return w.window_text() + '::' + w.element_info.control_type + textBW
	except Exception:
		return ''


def get_type_strings(events, allow_backspace=True):
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
	for event in events:
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


def get_send_keys_strings(e):
	return ''.join(format(code) for code in get_type_strings(e))


def print_pressed_keys(e):
	global main_overlay
	global record_file

	if e.name == 'r' and e.event_type == 'up' and 56 in keyboard._pressed_events:
		if record_file is None:
			new_path = 'Record files'
			if not os.path.exists(new_path):
				os.makedirs(new_path)
			record_file_name = './Record files/recorded ' + time.asctime() + '.py'
			record_file_name = record_file_name.replace(':','_')
			print 'Recording in file: ' + record_file_name
			keyboard.start_recording()
			record_file = open(record_file_name, "w")
			record_file.write("# coding: utf-8\n")
			record_file.write("import sys, os\n")
			record_file.write("sys.path.append(os.path.realpath(os.path.dirname(__file__)+'/..'))\n")
			record_file.write("from recorder_fn import *\n")
			record_file.write('send_keys("{LWIN down}""{DOWN}""{DOWN}""{LWIN up}")\n')
			record_file.write('time.sleep(1.0)\n')
			main_overlay_add_record_icon()
			main_overlay.refresh()
		else:
			keyboard_events = keyboard.stop_recording()
			if keyboard_events:
				keyboard_events = keyboard_events[:-1]
			line = get_send_keys_strings(keyboard_events)
			if line:
				record_file.write("send_keys(" + line + ")\n")
			record_file.close()
			record_file = None
			main_overlay.clear_all()
			main_overlay_add_pause_icon()
			main_overlay.refresh()
	elif e.name == 'q' and e.event_type == 'up' and 56 in keyboard._pressed_events:
		exit_recorder()

# TODO: % is relative to the center of the element, A is absolute, P is proportional, ...
def mouse_on_click(mouse_event):
	global record_file
	global unique_rectangle
	global unique_element_path
	if record_file is not None:
		keyboard_events = keyboard.stop_recording()
		if keyboard_events:
			keyboard_events = keyboard_events[:-1]
		line = get_send_keys_strings(keyboard_events)
		if line:
			record_file.write("send_keys(" + line + ")\n")

		x, y = win32api.GetCursorPos()
		record_file.write(mouse_event.button + "_")
		rx, ry = unique_rectangle.mid_point()
		dx , dy = x - rx,  y - ry
		record_file.write('click("""' + unique_element_path + '%(' + str(dx) + ',' + str(dy) + ')""")\n')

		keyboard.start_recording()


def mouse_on_drag(mouse_down_pos, mouse_down_unique_rectangle):
	global record_file
	global unique_rectangle
	global unique_element_path
	if record_file is not None:
		keyboard_events = keyboard.stop_recording()
		if keyboard_events:
			keyboard_events = keyboard_events[:-1]
		line = get_send_keys_strings(keyboard_events)
		if line:
			record_file.write("send_keys(" + line + ")\n")

		x, y = mouse_down_pos[0], mouse_down_pos[1]
		rx, ry = mouse_down_unique_rectangle.mid_point()
		dx, dy = x - rx, y - ry
		record_file.write('drag_and_drop("""' + unique_element_path + '%(' + str(dx) + ',' + str(dy))

		x2, y2 = win32api.GetCursorPos()
		dx, dy = x2 - rx, y2 - ry
		record_file.write(')%(' + str(dx) + ',' + str(dy) + ')""")\n')

		keyboard.start_recording()

def mouse_on_wheel(mouse_event):
	print 'TODO'

mouse_down_unique_rectangle = None
mouse_down_time = 0
mouse_down_pos = (0, 0)


def mouse_on(mouse_event):
	global unique_rectangle
	global mouse_down_unique_rectangle
	global mouse_down_time
	global mouse_down_pos
	if type(mouse_event) == mouse.MoveEvent:
		a = 0 # TODO: recording with timings
	elif type(mouse_event) == mouse.ButtonEvent:
		if mouse_event.event_type=='down':
			mouse_down_time = mouse_event.time
			mouse_down_pos = mouse.get_position()
			mouse_down_unique_rectangle = unique_rectangle
		if mouse_event.event_type=='up':
			if (mouse_event.time - mouse_down_time) < 0.2:
				mouse_on_click(mouse_event)
			else:
				if mouse_down_pos != mouse.get_position():
					mouse_on_drag(mouse_down_pos, mouse_down_unique_rectangle)
				else:
					mouse_on_click(mouse_event)
	elif type(mouse_event) == mouse.WheelEvent:
			mouse_on_wheel(mouse_event)


def main():
	global main_overlay
	global record_file
	global unique_rectangle
	global unique_element_path

	send_keys("{LWIN down}""{DOWN}""{DOWN}""{LWIN up}")

	keyboard.hook(print_pressed_keys)
	mouse.hook(mouse_on)

	unique_candidate = None
	elements = []

	pywinauto_desktop = pywinauto.Desktop(backend='uia', allow_magic_lookup=False)
	wrapper_old = None

	try:
		i = 0
		_is_running = True
		while _is_running:
			main_overlay.clear_all()

			x, y = win32api.GetCursorPos()
			elem = pywinauto.uia_defines.IUIA().iuia.ElementFromPoint(tagPOINT(x, y))
			element = pywinauto.uia_element_info.UIAElementInfo(elem)
			wrapper = pywinauto.controls.uiawrapper.UIAWrapper(element)
			if wrapper is None:
				continue

			element_path = get_element_path(wrapper)
			entry_list = (element_path.decode('utf-8')).split("->")
			unique_candidate, elements = find_element(pywinauto_desktop, entry_list, window_candidates=[])
			if unique_candidate is not None:
				unique_element_path = get_element_path(unique_candidate)
				# unique_candidate.draw_outline(colour='green', thickness=2)
				r = unique_candidate.rectangle()
				unique_rectangle = r
				main_overlay.add(geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(), height=r.height(),
								 thickness=1, color=(0, 128, 0), brush=oaam.Brush.solid, brush_color=(0,255,0))

				for e in elements:
					r = e.rectangle()
					main_overlay.add(geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(), height=r.height(),
									 thickness=1, color=(255, 0, 0), brush=oaam.Brush.solid, brush_color=(255, 0, 0))

			if (i % 2 == 0):
				if record_file is not None:
					main_overlay_add_record_icon()
				else:
					main_overlay_add_pause_icon()
			i = i + 1
			main_overlay.refresh()
			time.sleep(0.005) # main_overlay.clear_all() doit attendre la fin de main_overlay.refresh()
	except Exception as e:
		print ('Exception raised in main loop: \n')
		print(type(e))
		print(e.args)
		print(e)
		# raise

def exit_recorder():
	global record_file
	print("Quit")
	if record_file != None:
		record_file.close()

	# kill this process with taskkill
	current_pid = os.getpid()
	os.system("taskkill /pid %s /f" % current_pid)
	exit(0)

if __name__ == '__main__':
	main()


