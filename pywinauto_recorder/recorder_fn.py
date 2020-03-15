# -*- coding: utf-8 -*-

import pywinauto
import win32api
import win32con
import time
import re

from enum import Enum


class MoveMode(Enum):
	linear = 0
	y_first = 1
	x_first = 2


def is_int(s):
	try:
		int(s)
		return True
	except ValueError:
		return False


def get_entry(entry):
	p = re.compile('^(.*)::([^:]|[^#].*?)?(#\[.+?,[-+]?[0-9]+?\])?(%\(.+?,.+?\))?$')
	m = p.match(entry)
	if not m:
		return None, None, None, None
	str_name = m.group(1)
	str_type = m.group(2)
	str_array = m.group(3)
	str_dx_dy = m.group(4)
	if str_array:
		words = str_array.split(',')
		y = words[0][2:]
		if is_int(y):
			y = int(y)
		x = int(words[1][:-1])
		y_x = [y, x]
	else:
		y_x = None
	if str_dx_dy:
		words = str_dx_dy.split(',')
		dx = int(words[0][2:])
		dy = int(words[1][:-1])
		dx_dy = (dx, dy)
	else:
		dx_dy = None
	return str_name, str_type, y_x, dx_dy


# A FAIRE: voir pourquoi ca ne marche pas bien avec le dialogue Ajouter variable


def same_entry_list(element, entry_list):
	try:
		i = len(entry_list) - 1
		top_level_parent = element.top_level_parent()
		current_element = element
		while True:
			current_element_text = current_element.window_text()
			current_element_type = current_element.element_info.control_type
			entry_text, entry_type, _, _ = get_entry(entry_list[i])
			if current_element_text == entry_text and current_element_type == entry_type:
				if current_element == top_level_parent:
					return True
				i -= 1
				current_element = current_element.parent()
			else:
				return False
			if i == -1:
				return False
	except Exception:
		return False


def find_element(desktop, entry_list, window_candidates=[], visible_only=True, enabled_only=True, active_only=True):
	if not window_candidates:
		title, control_type, _, _ = get_entry(entry_list[0])
		window_candidates = desktop.windows(
											title=title, control_type=control_type, visible_only=visible_only,
											enabled_only=enabled_only, active_only=active_only)
		if not window_candidates:
			if active_only:
				return find_element(
									desktop, entry_list, window_candidates=[], visible_only=True,
									enabled_only=False, active_only=False)
			else:
				print ("Fatal error: No window found!")
				return None, []

	if len(entry_list) == 1:
		if len(window_candidates) == 1:
			return window_candidates[0], []

	candidates = []
	for window in window_candidates:
		title, control_type, _, _ = get_entry(entry_list[-1])
		descendants = window.descendants(title=title, control_type=control_type)
		for descendant in descendants:
			if same_entry_list(descendant, entry_list):
				candidates.append(descendant)
			else:
				continue

	if not candidates:
		if active_only:
			return find_element(
								desktop, entry_list, window_candidates=[], visible_only=True,
								enabled_only=False, active_only=False)
		else:
			return None, []
	elif len(candidates) == 1:
		return candidates[0], []
	else:
		# We have several elements so we have to use the good strategy to select the good one.
		# Strategy 1: 1D array of elements beginning with an element having a unique path
		# Strategy 2: 2D array of elements
		# Strategy 3: we find a unique path in the ancestors
		# unique_candidate, elements = find_element(desktop, entry_list[0:-1], window_candidates=window_candidates)
		# return unique_candidate, candidates
		return candidates[0], candidates


element_path_start = ''


def get_element_path_start():
	global element_path_start
	return element_path_start


unique_element_old = None
element_path_old = ''
w_rOLD = None
click_desktop = None


def in_region(element_path):
	global element_path_start
	element_path_start = element_path


def find(element_path):
	global element_path_start
	global click_desktop
	if not click_desktop:
		click_desktop = pywinauto.Desktop(backend='uia', allow_magic_lookup=False)

	if element_path_start:
		element_path2 = element_path_start + "->" + element_path
	else:
		element_path2 = element_path

	entry_list = (element_path2.decode('utf-8')).split("->")
	i = 0
	unique_element = None
	while i < 99:
		try:
			unique_element, _ = find_element(click_desktop, entry_list, window_candidates=[])
		except Exception:
			time.sleep(0.1)
		i += 1

		if unique_element is not None:
			# Wait if element is not clickable (greyed, not still visible)
			_, control_type0, _, _ = get_entry(entry_list[0])
			_, control_type1, _, _ = get_entry(entry_list[-1])
			if control_type0 == 'Menu' or control_type1 == 'TreeItem':
				app = pywinauto.Application(backend='uia', allow_magic_lookup=False)
				app.connect(process=unique_element.element_info.element.CurrentProcessId)
				app.wait_cpu_usage_lower()
			if unique_element.is_enabled():
				break
	if not unique_element:
		raise Exception("unique element not found! ", element_path)
	return unique_element


def move(element_path, duration=0.5, mode=MoveMode.linear, button='left'):
	global unique_element_old
	global element_path_old
	global w_rOLD

	entry_list = (element_path.decode('utf-8')).split("->")
	if element_path == element_path_old:
		w_r = w_rOLD
		unique_element = unique_element_old
	else:
		unique_element = find(element_path)
		w_r = unique_element.rectangle()

	x, y = win32api.GetCursorPos()
	_, control_type, _, _ = get_entry(entry_list[0])
	if control_type == 'Menu':
		entry_list_old = (element_path_old.decode('utf-8')).split("->")
		_, control_type_old, _, _ = get_entry(entry_list_old[0])
		if control_type_old == 'Menu':
			mode = MoveMode.x_first
		else:
			mode = MoveMode.y_first
	
		xd, yd = w_r.mid_point()
	else:
		_, _, _, dx_dy = get_entry(entry_list[-1])
		if dx_dy:
			dx, dy = dx_dy[0], dx_dy[1]
		else:
			dx, dy = 0, 0
		xd, yd = w_r.mid_point()
		xd, yd = xd + dx, yd + dy

	if (x, y) != (xd, yd):
		dt = 0.01
		samples = duration/dt
		step_x = (xd-x)/samples
		step_y = (yd-y)/samples

		if mode == MoveMode.x_first:
			for i in range(int(samples)):
				x = x+step_x
				time.sleep(0.01)
				nx = int(x) * 65535 / win32api.GetSystemMetrics(0)
				ny = int(y) * 65535 / win32api.GetSystemMetrics(1)
				win32api.mouse_event(win32con.MOUSEEVENTF_MOVE|win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)
			step_x = 0

		if mode == MoveMode.y_first:
			for i in range(int(samples)):
				y = y+step_y
				time.sleep(0.01)
				nx = int(x) * 65535 / win32api.GetSystemMetrics(0)
				ny = int(y) * 65535 / win32api.GetSystemMetrics(1)
				win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)
			step_y = 0

		for i in range(int(samples)):
			x, y = x+step_x, y+step_y
			time.sleep(0.01)
			nx = int(x) * 65535 / win32api.GetSystemMetrics(0)
			ny = int(y) * 65535 / win32api.GetSystemMetrics(1)
			win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)

	nx = int(xd) * 65535 / win32api.GetSystemMetrics(0)
	ny = int(yd) * 65535 / win32api.GetSystemMetrics(1)
	win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)

	unique_element_old = unique_element
	element_path_old = element_path
	w_rOLD = w_r
	return unique_element


def mouse_wheel(steps):
	win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, win32con.WHEEL_DELTA * steps, 0)


def click(element_path, duration=0.5, mode=MoveMode.linear, button='left'):
	unique_element = move(element_path, duration=duration, mode=mode, button=button)

	if button == 'left' or button == 'double_left' or button == 'triple_left':
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
		time.sleep(.01)
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
		time.sleep(.1)

	if button == 'double_left' or button == 'triple_left':
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
		time.sleep(.01)
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
		time.sleep(.1)

	if button == 'triple_left':
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
		time.sleep(.01)
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

	if button == 'right':
		win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0)
		time.sleep(.01)
		win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0)
		time.sleep(.01)

	return unique_element


def left_click(element_path, duration=0.5, mode=MoveMode.linear):
	return click(element_path, duration=duration, mode=mode, button='left')


def right_click(element_path, duration=0.5, mode=MoveMode.linear):
	return click(element_path, duration=duration, mode=mode, button='right')


def double_left_click(element_path, duration=0.5, mode=MoveMode.linear):
	return click(element_path, duration=duration, mode=mode, button='double_left')


def triple_left_click(element_path, duration=0.5, mode=MoveMode.linear):
	return click(element_path, duration=duration, mode=mode, button='triple_left')


def drag_and_drop(element_path, duration=0.5, mode=MoveMode.linear):
	words = element_path.split("%(")
	element_path1 = words[0] + "%(" + words[1]
	unique_element = move(element_path1, duration=duration, mode=mode)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
	words = element_path.split("%(")
	last_word = words[-1]
	words = words[:-1]
	words = words[:-1]
	element_path2 = ''
	for w in words:
		element_path2 = element_path2 + w + "%("
	element_path2 = element_path2 + last_word
	move(element_path2, duration=duration, mode=mode)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
	return unique_element


def send_keys(str_keys):
	pywinauto.keyboard.send_keys(str_keys, with_spaces=True)
