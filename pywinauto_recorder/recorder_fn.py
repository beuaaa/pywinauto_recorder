# -*- coding: utf-8 -*-

import pywinauto, win32api, win32con, time, anytree

from enum import Enum
class Move_mode(Enum):
	linear = 0
	y_first = 1
	x_first = 2
	

def get_window_text(entry):
	if len(entry)==0:
		return ''
	if entry[0]==':' and entry[1]==':':
		return ''
	else:
		i = entry.rfind('::')
		return entry[0:i]


def get_control_type(entry):
	i = entry.rfind('::')
	words = [entry[0:i], entry[i+2:]]
	if len(words)==2:
		control_type_dxy = words[1].split("%(")
		return control_type_dxy[0]
	else:
		return None


def get_dxdy(entry):
	words = entry.split("%(")
	dxVdy = words[1].split(")")[0]
	dxdy = dxVdy.split(",")
	return int(dxdy[0]), int(dxdy[1])

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


def find_element(desktop, entryList, window_candidates=[], visible_only=True, enabled_only=True, active_only=True):
	title = get_window_text(entryList[0])
	control_type = get_control_type(entryList[0])

	if len(window_candidates)==0:
		window_candidates = desktop.windows(title=title, control_type=control_type, visible_only=visible_only,
											enabled_only=enabled_only, active_only=active_only)
	if len(entryList) == 1:
		if len(window_candidates) == 1:
			return window_candidates[0], []

	if False: # A voir
		candidates = []
		for window in window_candidates:
			#descendants = window.descendants(title=get_window_text(entryList[-1]), control_type=get_control_type(entryList[-1]))
			try:
				descendants = window.child_window(title=get_window_text(entryList[-1]),
													control_type=get_control_type(entryList[-1]),
													depth=len(entryList) )
			except pywinauto.findwindows.ElementAmbiguousError as e:
				descendants = e.elements
			finally:
				candidates = candidates + descendants

	else:
		candidates = []
		for window in window_candidates:
			descendants = window.descendants(title=get_window_text(entryList[-1]), control_type=get_control_type(entryList[-1]))
			for descendant in descendants:
				if same_entry_list(descendant, entryList):
					candidates.append(descendant)
				else:
					continue

	if len(candidates)==0:
		if active_only==False:
			return None, []
		else:
			return find_element(desktop, entryList, window_candidates=[], visible_only=True, enabled_only=False, active_only=False)

	elif len(candidates)==1:
		#add_wrapper_cache(candidates[0], entryList)
		return candidates[0], []
	else:
		unique_candidate, elements = find_element(desktop, entryList[0:-1], window_candidates=window_candidates)
		#add_wrapper_cache(unique_candidate, entryList[0:-1])
		return unique_candidate, candidates


elementPathOLD = ''
w_rOLD = None
click_desktop = pywinauto.Desktop(backend='uia', allow_magic_lookup=False)


def move(element_path, duration=0.5, mode=Move_mode.linear, button='left'):
	global click_desktop
	global elementPathOLD
	global w_rOLD

	entryList = (element_path.decode('utf-8')).split("->")
	if element_path == elementPathOLD:
		w_r = w_rOLD
	else:
		i=0
		unique_element = None
		elements = []
		while i<99:
			try:
				unique_element, elements = find_element(click_desktop, entryList, window_candidates=[])
			except:
				time.sleep(0.1)
			i+=1

			if unique_element is not None:
				if get_control_type(entryList[0]) == 'Menu' or get_control_type(entryList[-1]) == 'TreeItem':
					app = pywinauto.Application(backend='uia', allow_magic_lookup=False)
					app.connect(process=unique_element.element_info.element.CurrentProcessId)
					app.wait_cpu_usage_lower()
				if unique_element.is_enabled():
					break
		w_r = unique_element.rectangle()

	x, y = win32api.GetCursorPos()
	if get_control_type(entryList[0])=='Menu':
		entryListOLD = (elementPathOLD.decode('utf-8')).split("->")
		if get_control_type(entryListOLD[0])=='Menu':
			mode = Move_mode.x_first
		else:
			mode = Move_mode.y_first
	
		xd, yd = w_r.mid_point()
	else:
		dx, dy = get_dxdy(entryList[-1])
		xd, yd = w_r.mid_point()
		xd, yd = xd + dx, yd + dy

	if (x,y) != (xd,yd):
		dt = 0.01
		samples = duration/dt
		step_x = (xd-x)/(samples)
		step_y = (yd-y)/(samples)

		if mode==Move_mode.x_first:
			for i in range(int(samples)):
				x = x+step_x
				time.sleep(0.01)
				#win32api.SetCursorPos((int(x),int(y)))
				#pywinauto.mouse.move(coords=(int(x),int(y)))
				nx = int(x) * 65535 / win32api.GetSystemMetrics(0)
				ny = int(y) * 65535 / win32api.GetSystemMetrics(1)
				win32api.mouse_event(win32con.MOUSEEVENTF_MOVE|win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)
			step_x = 0

		if mode==Move_mode.y_first:
			for i in range(int(samples)):
				y = y+step_y
				time.sleep(0.01)
				#win32api.SetCursorPos((int(x),int(y)))
				#pywinauto.mouse.move(coords=(int(x), int(y)))
				nx = int(x) * 65535 / win32api.GetSystemMetrics(0)
				ny = int(y) * 65535 / win32api.GetSystemMetrics(1)
				win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)
			step_y = 0

		for i in range(int(samples)):
			x,y = x+step_x, y+step_y
			time.sleep(0.01)
			#win32api.SetCursorPos((int(x),int(y)))
			#pywinauto.mouse.move(coords=(int(x), int(y)))
			nx = int(x) * 65535 / win32api.GetSystemMetrics(0)
			ny = int(y) * 65535 / win32api.GetSystemMetrics(1)
			win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)

	#win32api.SetCursorPos((int(xd), int(yd)))
	#pywinauto.mouse.move(coords=(int(xd), int(yd)))
	nx = int(xd) * 65535 / win32api.GetSystemMetrics(0)
	ny = int(yd) * 65535 / win32api.GetSystemMetrics(1)
	win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)

	elementPathOLD = element_path
	w_rOLD = w_r
	return unique_element

def mouse_wheel(steps):
	win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, steps)

def click(element_path, duration=0.5, mode=Move_mode.linear, button='left'):
	move(element_path, duration=duration, mode=mode, button=button)

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
		win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,0,0)
		time.sleep(.01)
		win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,0,0)
		time.sleep(.01)


def left_click(element_path, duration = 0.5, mode = Move_mode.linear):
	click(element_path, duration = duration, mode = mode, button='left')


def right_click(element_path, duration = 0.5, mode = Move_mode.linear):
	click(element_path, duration = duration, mode = mode, button='right')


def double_left_click(element_path, duration = 0.5, mode = Move_mode.linear):
	click(element_path, duration = duration, mode = mode, button='double_left')


def triple_left_click(element_path, duration = 0.5, mode = Move_mode.linear):
	click(element_path, duration = duration, mode = mode, button='triple_left')


def drag_and_drop(element_path, duration=0.5, mode=Move_mode.linear):
	move(element_path, duration=duration, mode=mode)
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


def send_keys(str_keys):
	pywinauto.keyboard.send_keys(str_keys, with_spaces=True)