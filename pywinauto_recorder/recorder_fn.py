# -*- coding: utf-8 -*-

import pywinauto, win32api, win32con, time

from enum import Enum


class MoveMode(Enum):
	linear = 0
	y_first = 1
	x_first = 2
	

def get_window_text(entry):
	if not entry:
		return ''
	if entry[0] == ':' and entry[1] == ':':
		return ''
	else:
		i = entry.rfind('::')
		return entry[0:i]


def get_control_type(entry):
	i = entry.rfind('::')
	words = [entry[0:i], entry[i+2:]]
	if len(words) == 2:
		control_type_dxy = words[1].split("%(")
		return control_type_dxy[0]
	else:
		return None


def get_dx_dy(entry):
	words = entry.split("%(")
	dx_dy_word = words[1].split(")")[0]
	dx_dy = dx_dy_word.split(",")
	return int(dx_dy[0]), int(dx_dy[1])

# A FAIRE: voir pourquoi ca ne marche pas bien avec le dialogue Ajouter variable


def same_entry_list(element, entry_list):
	try:
		i = len(entry_list) - 1
		top_level_parent = element.top_level_parent()
		current_element = element
		while True:
			current_element_text = current_element.window_text()
			current_element_type = current_element.element_info.control_type
			entry_text = get_window_text(entry_list[i])
			entry_type = get_control_type(entry_list[i])
			if current_element_text == entry_text and current_element_type == entry_type:
				if current_element == top_level_parent:
					return True
				i -= 1
				current_element = current_element.parent()
			else:
				return False
			if i == -1:
				return False

		return False
	except Exception:
		return False


def find_element(desktop, entry_list, window_candidates=[], visible_only=True, enabled_only=True, active_only=True):
	title = get_window_text(entry_list[0])
	control_type = get_control_type(entry_list[0])

	if not window_candidates:
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
		descendants = window.descendants(
											title=get_window_text(entry_list[-1]),
											control_type=get_control_type(entry_list[-1]))
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
		unique_candidate, elements = find_element(desktop, entry_list[0:-1], window_candidates=window_candidates)
		return unique_candidate, candidates


element_path_old = ''
w_rOLD = None
click_desktop = pywinauto.Desktop(backend='uia', allow_magic_lookup=False)


def find(element_path):
	entry_list = (element_path.decode('utf-8')).split("->")
	i = 0
	unique_element = None
	elements = []
	while i < 99:
		try:
			unique_element, elements = find_element(click_desktop, entry_list, window_candidates=[])
		except:
			time.sleep(0.1)
		i+=1

		if unique_element is not None:
			if get_control_type(entry_list[0]) == 'Menu' or get_control_type(entry_list[-1]) == 'TreeItem':
				app = pywinauto.Application(backend='uia', allow_magic_lookup=False)
				app.connect(process=unique_element.element_info.element.CurrentProcessId)
				app.wait_cpu_usage_lower()
			if unique_element.is_enabled():
				break
	return unique_element


def move(element_path, duration=0.5, mode=MoveMode.linear, button='left'):
	global click_desktop
	global element_path_old
	global w_rOLD

	entry_list = (element_path.decode('utf-8')).split("->")
	if element_path == element_path_old:
		w_r = w_rOLD
	else:
		unique_element = find(element_path)
		w_r = unique_element.rectangle()

	x, y = win32api.GetCursorPos()
	if get_control_type(entry_list[0]) == 'Menu':
		entry_list_old = (element_path_old.decode('utf-8')).split("->")
		if get_control_type(entry_list_old[0]) == 'Menu':
			mode = MoveMode.x_first
		else:
			mode = MoveMode.y_first
	
		xd, yd = w_r.mid_point()
	else:
		dx, dy = get_dx_dy(entry_list[-1])
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

	element_path_old = element_path
	w_rOLD = w_r
	return unique_element


def mouse_wheel(steps):
	win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, steps)


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
	unique_element = move(element_path, duration=duration, mode=mode)
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
