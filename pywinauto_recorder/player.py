"""This module contains functions to replay a sequence of user actions automatically."""

# print('__file__={0:<35} | __name__={1:<20} | __package__={2:<20}'.format(__file__, __name__, str(__package__)))
from enum import Enum
from typing import Optional, Union, NewType
import time
import pywinauto
from win32api import GetCursorPos as win32api_GetCursorPos
from win32api import GetSystemMetrics as win32api_GetSystemMetrics
from win32api import mouse_event as win32api_mouse_event
from win32gui import LoadCursor as win32gui_LoadCursor
from win32gui import GetCursorInfo as win32gui_GetCursorInfo
from win32con import IDC_WAIT, MOUSEEVENTF_MOVE, MOUSEEVENTF_ABSOLUTE, MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP, \
	MOUSEEVENTF_MIDDLEDOWN, MOUSEEVENTF_MIDDLEUP, MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP, MOUSEEVENTF_WHEEL, \
	WHEEL_DELTA
from .core import type_separator, path_separator, get_entry, get_entry_list, find_element, get_sorted_region, \
	get_wrapper_path, is_int


UI_Coordinates = NewType('UI_Coordinates', (float, float))
UI_Path = str
PYWINAUTO_Wrapper = pywinauto.controls.uiawrapper.UIAWrapper
UI_Selector = Union[UI_Path, PYWINAUTO_Wrapper, UI_Coordinates]


class PywinautoRecorderException(Exception):
	"""Base class for other exceptions"""
	pass


class FailedSearch(PywinautoRecorderException):
	"""FailedSearch is a subclass of *PywinautoRecorderException* that is raised when a search for a control fails."""
	pass


__all__ = ['PlayerSettings', 'MoveMode', 'load_dictionary', 'shortcut', 'full_definition', 'UIPath', 'Window', 'Region',
           'find', 'move', 'click', 'left_click', 'right_click', 'double_left_click', 'triple_left_click',
           'drag_and_drop', 'middle_drag_and_drop', 'right_drag_and_drop', 'menu_click', 'mouse_wheel', 'send_keys',
           'set_combobox', 'set_text', 'exists', 'select_file']


# TODO special_char_array in core for recorder.py and player.py (check when to call escape & unescape)
def unescape_special_char(string):
	for r in (("\\\\", "\\"), ("\\t", "\t"), ("\\n", "\n"), ("\\r", "\r"), ("\\v", "\v"), ("\\f", "\f"), ('\\"', '"')):
		#for r in (("\\", "\\\\"), ("\t", "\\t"), ("\n", "\\n"), ("\r", "\\r"), ("\v", "\\v"), ("\f", "\\f"), ('"', '\\"')):
		string = string.replace(*r)
	return string


class PlayerSettings:
	mouse_move_duration = 0.5
	timeout = 10


class MoveMode(Enum):
	linear = 0
	y_first = 1
	x_first = 2


_dictionary = {}
unique_element_old = None
element_path_old = ''
w_rOLD = None


def load_dictionary(filename_key: str, filename_def: str,encoding: str = 'utf8') -> None:
	"""
	Loads a dictionary

	:param filename_key: filename of the key file
	:param filename_def: filename of the definition file
	:param encoding: encoding of the dictionary file
	"""
	abs_path = [x for x in range(99)]
	with open(filename_key, encoding=encoding) as fp_key, open(filename_def, encoding=encoding) as fp_def:
		for line_key, line_def in zip(fp_key, fp_def):
			words = line_key.split("\t")
			word = words[-1].translate(str.maketrans('', '', '\n\t\r'))
			
			words = line_def.split("\t")
			definition = words[-1].translate(str.maketrans('', '', '\n\t\r'))
			level = len(words) - 1
			#print(level)
			abs_path[level] = definition
			abs_definition = abs_path[0]
			for i in range(1, level):
				abs_definition += path_separator + abs_path[i]
			#print(abs_definition + path_separator + definition)
			_dictionary[word] = (abs_definition, definition)


def shortcut(str_shortcut: str) -> str:
	"""
	Returns the shortcut path associated to the shortcut defined in the previously loaded dictionary

	:param str_shortcut: shortcut
	"""
	return _dictionary[str_shortcut][1]


def full_definition(str_shortcut: str) -> str:
	"""
	Returns the full element path associated to the shortcut defined in the previously loaded dictionary

	:param str_shortcut: shortcut
	"""
	return _dictionary[str_shortcut][0] + path_separator + _dictionary[str_shortcut][1]


def wait_is_ready_try1(wrapper, timeout=120):
	"""
	Waits until element is ready (wait while greyed, not enabled, not visible, not ready, ...) :
	So far, I didn't find better than wait_cpu_usage_lower when greyed but must be enhanced
	"""
	t0 = time.time()
	while not wrapper.is_enabled() or not wrapper.is_visible():
		try:
			h_wait_cursor = win32gui_LoadCursor(0, IDC_WAIT)
			_, h_cursor, _ = win32gui_GetCursorInfo()
			app = pywinauto.Application(backend='uia', allow_magic_lookup=False)
			app.connect(process=wrapper.element_info.element.CurrentProcessId)
			while h_cursor == h_wait_cursor:
				app.wait_cpu_usage_lower()
			spec = app.window(handle=wrapper.handle, top_level_only=False)
			while not wrapper.is_enabled() or not wrapper.is_visible():
				spec.wait("exists enabled visible ready")
				if (time.time() - t0) > timeout:
					break
		except Exception:
			time.sleep(0.1)
		if (time.time() - t0) > timeout:
			msg = "Element " + get_wrapper_path(wrapper) + "  was not found after " + str(timeout) + " s of searching."
			raise TimeoutError("Time out! ", msg)


# > The class is used to keep track of the current path in the UI tree
class UIPath(object):
	_path_list = []
	_regex_list = []
	
	@staticmethod
	def get_full_path(element_path: Optional[UI_Path] = None) -> UI_Path:
		"""
		If the element_path is None, return the full path of the current UI element

		:param element_path: Optional[UI_Path] = None
		:type element_path: Optional[UI_Path]
		:return: The full path of the element.
		"""
		if element_path is None:
			return path_separator.join(UIPath._path_list)
		else:
			return path_separator.join(UIPath._path_list) + path_separator + element_path
	
	def __init__(self, relative_path=None, regex_title=False):
		self.relative_path = relative_path
		self.regex_title = regex_title
	
	def __enter__(self):
		if self.relative_path:
			UIPath._path_list.append(self.relative_path)
			UIPath._regex_list.append(self.regex_title)
		return self
	
	def __exit__(self, type, value, traceback):
		if self.relative_path:
			UIPath._path_list = UIPath._path_list[0:-1]
			UIPath._regex_list = UIPath._regex_list[0:-1]


Window = UIPath
Region = UIPath


def find(
		element_path: Optional[UI_Selector] = None,
		regex: bool=False,
		timeout: Optional[float] = None) -> PYWINAUTO_Wrapper:
	"""
	Finds an element

	:param element_path: element path
	:param timeout: period of time in seconds that will be allowed to find the element
	:return: Pywinauto wrapper of clicked element
	"""
	if regex:
		deprecated_msg = """
		The parameter 'regex' is deprecated. Please use the new RegEx syntax of UIPath!
		For example:
		
		with Window(".* - Notepad||Window", regex=True):
			edit = left_click("Text Editor||Edit")
		
		Must be coded with the new syntax:
		
		with UIPath("RegEx: .* - Notepad||Window"):
			edit = left_click("Text Editor||Edit")
			
		This parameter will be removed in the next release.
		"""
		print(deprecated_msg)
	if timeout is None:
		timeout = PlayerSettings.timeout
	
	full_element_path = UIPath.get_full_path(element_path)
	entry_list = get_entry_list(full_element_path)
	_, _, y_x, _ = get_entry(entry_list[-1])
	unique_element = None
	elements = None
	t0 = time.time()
	while (time.time() - t0) < timeout:
		while (not y_x and not unique_element and not elements) or (y_x and not elements):
			try:
				# print("find_element(...): ")
				unique_element, elements = find_element(entry_list)
				if (not y_x and not unique_element and not elements) or (y_x and not elements):
					time.sleep(2.0)
			except Exception:
				pass
			if (time.time() - t0) > timeout:
				msg = "No element found with the UIPath '" + full_element_path + "' after " + str(timeout) + " s of searching."
				raise FailedSearch(msg)
		
		if y_x is not None:
			if unique_element and int(y_x[0]) == int(y_x[1]) == 0:
				return unique_element
			else:
				nb_y, nb_x, candidates = get_sorted_region(elements)
				if is_int(y_x[0]):
					unique_element = candidates[int(y_x[0])][int(y_x[1])]
				else:
					# TODO: ref_unique_element, _ = find_element(get_entry_list(y_x[0]), window_candidates=[]) (dans tous le code)
					ref_entry_list = get_entry_list(full_element_path) + get_entry_list(y_x[0])
					ref_unique_element, _ = find_element(ref_entry_list)
					if not ref_unique_element:
						msg = "No element found with the UIPath '" + full_element_path + path_separator + y_x[0] + "' in the array line."
						raise FailedSearch(msg)
					ref_r = ref_unique_element.rectangle()
					r_y = 0
					while r_y < nb_y:
						for candidate in candidates[r_y]:
							y_candidate = candidate.rectangle().mid_point()[1]
							if ref_r.top < y_candidate < ref_r.bottom:
								unique_element = candidates[r_y][y_x[1]]
								return unique_element
						r_y = r_y + 1
		if unique_element is not None:
			break
		time.sleep(0.1)
	if not unique_element:
		full_element_path = UIPath.get_full_path(element_path)
		if elements:
			raise FailedSearch("There are " + str(len(elements)) + " elements that match the path '" + full_element_path + "'")
		raise FailedSearch("Unique element not found using path '", full_element_path + "'")
	return unique_element


def move(
		element_path: UI_Selector,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		timeout: float = 120) -> PYWINAUTO_Wrapper:
	"""
	Moves on element
	
	:param element_path: element path
	:param duration: duration in seconds of the mouse move (if duration is -1 the mouse cursor doesn't move)
	:param mode: move mouse mode
	:param timeout: period of time in seconds that will be allowed to find the element
	:return: Pywinauto wrapper of clicked element
	"""
	if duration is None:
		duration = PlayerSettings.mouse_move_duration
	
	if duration == -1:
		return
	
	global unique_element_old
	global element_path_old
	global w_rOLD
	
	x, y = win32api_GetCursorPos()
	if isinstance(element_path, str):
		element_path2 = UIPath.get_full_path(element_path)
		entry_list = get_entry_list(element_path2)
		if element_path2 == element_path_old:
			w_r = w_rOLD
			unique_element = unique_element_old
		else:
			unique_element = find(element_path, timeout=timeout)
			w_r = unique_element.rectangle()
		control_type = None
		for entry in entry_list:
			_, control_type, _, _ = get_entry(entry)
			if control_type == 'Menu':
				break
		if control_type == 'Menu':
			entry_list_old = get_entry_list(element_path_old)
			control_type_old = None
			for entry in entry_list_old:
				_, control_type_old, _, _ = get_entry(entry)
				if control_type_old == 'Menu':
					break
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
			xd, yd = xd + round(dx/100.0*(w_r.width()/2-1), 0), round(yd + dy/100.0*(w_r.height()/2-1), 0)
	elif issubclass(type(element_path), pywinauto.base_wrapper.BaseWrapper):
		unique_element = element_path
		element_path2 = get_wrapper_path(unique_element)
		w_r = unique_element.rectangle()
		xd, yd = w_r.mid_point()
	else:
		(xd, yd) = element_path
		unique_element = None
	x_max = win32api_GetSystemMetrics(0) - 1
	y_max = win32api_GetSystemMetrics(1) - 1
	if (x, y) != (xd, yd) and duration > 0:
		dt = 0.01
		samples = duration/dt
		step_x = (xd-x)/samples
		step_y = (yd-y)/samples
		if mode == MoveMode.x_first:
			for _ in range(int(samples)):
				x = x+step_x
				time.sleep(0.01)
				nx = int(x * 65535 / x_max)
				ny = int(y * 65535 / y_max)
				win32api_mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, nx, ny)
			step_x = 0
		if mode == MoveMode.y_first:
			for _ in range(int(samples)):
				y = y+step_y
				time.sleep(0.01)
				nx = int(x * 65535 / x_max)
				ny = int(y * 65535 / y_max)
				win32api_mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, nx, ny)
			step_y = 0
		for _ in range(int(samples)):
			x, y = x+step_x, y+step_y
			time.sleep(0.01)
			nx = int(x * 65535 / x_max)
			ny = int(y * 65535 / y_max)
			win32api_mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, nx, ny)
	nx = round(xd * 65535 / x_max)
	ny = round(yd * 65535 / y_max)
	win32api_mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, nx, ny)
	if unique_element is None:
		return None
	unique_element_old = unique_element
	element_path_old = element_path2
	w_rOLD = w_r
	return unique_element


def click(
		element_path: Optional[UI_Selector] = None,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		button: str = 'left',
		timeout: float = None,
		wait_ready: bool = True) -> PYWINAUTO_Wrapper:
	"""
	Clicks on element
	
	:param element_path: element path
	:param duration: duration in seconds of the mouse move
		(if duration is -1 the mouse cursor doesn't move, it just sends WM_CLICK window message,
		useful for minimized or non-active window).
	:param mode: move mouse mode
	:param button: mouse button: 'left','double_left', 'triple_left', 'right'
	:param timeout: period of time in seconds that will be allowed to find the element
	:param wait_ready: if True waits until the element is ready
	:return: Pywinauto wrapper of clicked element
	"""
	if duration is None:
		duration = PlayerSettings.mouse_move_duration
	
	if timeout is None:
		timeout = PlayerSettings.timeout
	
	if element_path:
		if duration == -1:
			wrapper = find(element_path)
			has_get_value = getattr(wrapper, "click", None)
			if callable(has_get_value):
				wrapper.click()
			else:
				wrapper.click_input()
			return wrapper
		else:
			wrapper = move(element_path, duration=duration, mode=mode, timeout=timeout)
			if wait_ready and isinstance(element_path, str):
				wait_is_ready_try1(wrapper, timeout=timeout)
			else:
				wrapper = None
	if button == 'left' or button == 'double_left' or button == 'triple_left':
		win32api_mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0)
		time.sleep(.01)
		win32api_mouse_event(MOUSEEVENTF_LEFTUP, 0, 0)
		time.sleep(.1)
	if button == 'double_left' or button == 'triple_left':
		win32api_mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0)
		time.sleep(.01)
		win32api_mouse_event(MOUSEEVENTF_LEFTUP, 0, 0)
		time.sleep(.1)
	if button == 'triple_left':
		win32api_mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0)
		time.sleep(.01)
		win32api_mouse_event(MOUSEEVENTF_LEFTUP, 0, 0)
	if button == 'right':
		win32api_mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0)
		time.sleep(.01)
		win32api_mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0)
		time.sleep(.01)
	if element_path:
		return wrapper
	else:
		return None


def left_click(
		element_path: UI_Selector,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		timeout: Optional[float] = None,
		wait_ready: bool = True) -> PYWINAUTO_Wrapper:
	"""
	Left clicks on element
	
	:param element_path: element path
	:param duration: duration in seconds of the mouse move (if duration is -1 the mouse cursor doesn't move)
	:param mode: move mouse mode
	:param timeout: period of time in seconds that will be allowed to find the element
	:param wait_ready: if True waits until the element is ready
	:return: Pywinauto wrapper of clicked element
	"""
	return click(element_path, duration=duration, mode=mode, button='left', timeout=timeout, wait_ready=wait_ready)


def right_click(
		element_path: UI_Selector,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		timeout: Optional[float] = None,
		wait_ready: bool = True) -> PYWINAUTO_Wrapper:
	"""
	Right clicks on element
	
	:param element_path: element path
	:param duration: duration in seconds of the mouse move (if duration is -1 the mouse cursor doesn't move)
	:param mode: move mouse mode
	:param timeout: period of time in seconds that will be allowed to find the element
	:param wait_ready: if True waits until the element is ready
	:return: Pywinauto wrapper of clicked element
	"""
	return click(element_path, duration=duration, mode=mode, button='right', timeout=timeout, wait_ready=wait_ready)


def double_left_click(
		element_path: UI_Selector,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		timeout: Optional[float] = None,
		wait_ready: bool = True) -> PYWINAUTO_Wrapper:
	"""
	Double left clicks on element
	
	:param element_path: element path
	:param duration: duration in seconds of the mouse move (if duration is -1 the mouse cursor doesn't move)
	:param mode: move mouse mode
	:param timeout: period of time in seconds that will be allowed to find the element
	:param wait_ready: if True waits until the element is ready
	:return: Pywinauto wrapper of clicked element
	"""
	return click(element_path, duration=duration, mode=mode, button='double_left', timeout=timeout, wait_ready=wait_ready)


def triple_left_click(
		element_path: UI_Selector,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		timeout: Optional[float] = None,
		wait_ready: bool = True) -> PYWINAUTO_Wrapper:
	"""
	Triple left clicks on element
	
	:param element_path: element path
	:param duration: duration in seconds of the mouse move (if duration is -1 the mouse cursor doesn't move)
	:param mode: move mouse mode
	:param timeout: period of time in seconds that will be allowed to find the element
	:param wait_ready: if True waits until the element is ready
	:return: Pywinauto wrapper of clicked element
	"""
	return click(element_path, duration=duration, mode=mode, button='triple_left', timeout=timeout, wait_ready=wait_ready)


def drag_and_drop(
		element_path1: UI_Selector,
		element_path2: UI_Selector,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		timeout: Optional[float] = None) -> PYWINAUTO_Wrapper:
	"""
	Drags and drop with left button pressed from element_path1 to element_path2.
	
	:param element_path1: element path
	:param element_path2: element path
	:param duration: duration in seconds of the mouse move (if duration is -1 the mouse cursor doesn't move)
	:param mode: move mouse mode
	:param timeout: period of time in seconds that will be allowed to find the element
	:return: Pywinauto wrapper with element_path2
	"""
	move(element_path1, duration=duration, mode=mode, timeout=timeout)
	win32api_mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0)
	unique_element = move(element_path2, duration=duration, timeout=timeout)
	win32api_mouse_event(MOUSEEVENTF_LEFTUP, 0, 0)
	return unique_element


def middle_drag_and_drop(
		element_path1: UI_Selector,
		element_path2: UI_Selector,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		timeout: Optional[float] = None) -> PYWINAUTO_Wrapper:
	"""
	Drags and drop with middle button pressed from element_path1 to element_path2.
	
	:param element_path1: element path
	:param element_path2: element path
	:param duration: duration in seconds of the mouse move (if duration is -1 the mouse cursor doesn't move)
	:param mode: move mouse mode
	:param timeout: period of time in seconds that will be allowed to find the element
	:return: Pywinauto wrapper with element_path2
	"""
	move(element_path1, duration=duration, mode=mode, timeout=timeout)
	win32api_mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0)
	unique_element = move(element_path2, duration=duration, mode=mode, timeout=timeout)
	win32api_mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0)
	return unique_element


def right_drag_and_drop(
		element_path1: UI_Selector,
		element_path2: UI_Selector,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		timeout: Optional[float] = None) -> PYWINAUTO_Wrapper:
	"""
	Drags and drop with right button pressed from element_path1 to element_path2.
	
	:param element_path1: element path
	:param element_path2: element path
	:param duration: duration in seconds of the mouse move (if duration is -1 the mouse cursor doesn't move)
	:param mode: move mouse mode
	:param timeout: period of time in seconds that will be allowed to find the element
	:return: Pywinauto wrapper with element_path2
	"""
	move(element_path1, duration=duration, mode=mode, timeout=timeout)
	win32api_mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0)
	unique_element = move(element_path2, duration=duration, mode=mode, timeout=timeout)
	win32api_mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0)
	return unique_element


def menu_click(
		element_path: UI_Selector,
		menu_path: str,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		menu_type: str = 'QT',
		timeout: Optional[float] = None) -> PYWINAUTO_Wrapper:
	"""
	Clicks on menu item.
	
	:param element_path: element path
	:param menu_path: menu path
	:param duration: duration in seconds of the mouse move (if duration is -1 the mouse cursor doesn't move)
	:param mode: move mouse mode
	:param menu_type: menu type ('QT', 'NPP')
	:param timeout: period of time in seconds that will be allowed to find the element
	:return: Pywinauto wrapper of the clicked item
	"""
	menu_entry_list = menu_path.split(path_separator)
	if menu_type == 'QT':
		menu_entry_list = [''] + menu_entry_list
	else:
		menu_entry_list = ['Application'] + menu_entry_list
	if element_path:
		element_path2 = element_path + path_separator
	else:
		element_path2 = ''
	left_click(element_path2 +
	           menu_entry_list[0] + type_separator + 'MenuBar' + path_separator +
	           menu_entry_list[1] + type_separator + 'MenuItem', duration=duration, mode=mode, timeout=timeout)
	w = None
	if menu_type == 'QT':
		path_list_old, regex_list_old = UIPath.path_list, UIPath.regex_list
		UIPath.path_list, UIPath.regex_list = [], []
		for entry in menu_entry_list[2:]:
			w = left_click(type_separator + 'Menu' + path_separator + entry + type_separator + 'MenuItem',
			               duration=duration, mode=mode, timeout=timeout)
		UIPath.path_list, UIPath.regex_list = path_list_old, regex_list_old
	else:
		for i, entry in enumerate(menu_entry_list[2:]):
			w = left_click(element_path +
			               menu_entry_list[i - 2] + type_separator + 'Menu' + path_separator +
			               unescape_special_char(entry) + type_separator + 'MenuItem',
			               duration=duration, mode=mode, timeout=timeout)
	return w


def mouse_wheel(steps: int, pause: float = 0) -> None:
	"""
	Turns the mouse wheel up or down.
	
	:param steps: number of wheel steps, if positive the mouse wheel is turned up else it is turned down
	:param pause: pause in seconds between each wheel step
	"""
	if pause == 0:
		win32api_mouse_event(MOUSEEVENTF_WHEEL, 0, 0, WHEEL_DELTA * steps, 0)
	else:
		for _ in range(abs(steps)):
			if steps > 0:
				win32api_mouse_event(MOUSEEVENTF_WHEEL, 0, 0, WHEEL_DELTA, 0)
			else:
				win32api_mouse_event(MOUSEEVENTF_WHEEL, 0, 0, -WHEEL_DELTA, 0)
			time.sleep(pause)


def send_keys(
		str_keys: str,
		pause: float = 0.1,
		with_spaces: bool = True,
		with_tabs: bool = True,
		with_newlines: bool = True,
		turn_off_numlock: bool = True,
		vk_packet: bool = True) -> None:
	"""
	Parses the keys and type them
	You can use any Unicode characters (on Windows) and some special keys.
	See https://pywinauto.readthedocs.io/en/latest/code/pywinauto.keyboard.html
	
	:param str_keys: string representing the keys to be typed
	:param pause: pause in seconds between each typed key
	:param with_spaces: if False spaces are not taken into account
	:param with_tabs: if False tabs are not taken into account
	:param with_newlines: if False newlines are not taken into account
	:param turn_off_numlock: if True numlock is turned off
	:param vk_packet: For Windows only, pywinauto defaults to sending a virtual key packet (VK_PACKET) for textual input
	"""
	for r in (('(', '{(}'),  (')', '{)}'), ('+', '{+}')):
		str_keys = str_keys.replace(*r)
	pywinauto.keyboard.send_keys(   # lgtm [py/call/wrong-named-argument]
		str_keys,
		pause=pause,
		with_spaces=with_spaces,
		with_tabs=with_tabs,
		with_newlines=with_newlines,
		turn_off_numlock=turn_off_numlock,
		vk_packet=vk_packet
	)


def set_combobox(
		element_path: UI_Selector,
		value: str,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		timeout: Optional[float] = None,
		wait_ready: bool = True) -> None:
	"""
	Sets the value of a combobox.
	
	:param element_path: element path
	:param value: value of the combobox
	:param duration: duration in seconds of the mouse move (if duration is -1 the mouse cursor doesn't move)
	:param mode: move mouse mode
	:param timeout: period of time in seconds that will be allowed to find the element
	"""
	left_click(element_path, duration=duration, mode=mode, timeout=timeout, wait_ready=wait_ready)
	time.sleep(0.9)
	send_keys(value + "{ENTER}")


def set_text(
		element_path: UI_Selector,
		value: str,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		timeout: Optional[float] = None,
		pause: float = 0.1) -> None:
	"""
	Sets the value of a text field.
	
	:param element_path: element path
	:param value: value of the combobox
	:param duration: duration in seconds of the mouse move (if duration is -1 the mouse cursor doesn't move)
	:param mode: move mouse mode
	:param timeout: period of time in seconds that will be allowed to find the element
	:param pause: pause in seconds between each typed key
	"""
	double_left_click(element_path, duration=duration, mode=mode, timeout=timeout)
	send_keys("{VK_CONTROL down}a{VK_CONTROL up}", pause=0)
	time.sleep(0.1)
	send_keys(value + "{ENTER}", pause=pause)


def exists(
		element_path: UI_Selector,
		timeout: Optional[float] = None) -> PYWINAUTO_Wrapper:
	"""
	Tests if en UI_Element exists.
	
	:param element_path: element path
	:param timeout: period of time in seconds that will be allowed to find the element
	:return: Pywinauto wrapper of the found element or None
	"""
	if timeout is None:
		timeout = PlayerSettings.timeout
	try:
		wrapper = find(element_path, timeout=timeout)
		return wrapper
	except TimeoutError:
		return None


def select_file(
		window_path: UI_Selector,
		full_path: str,
		force_slow_path_typing: bool = False) -> None:
	"""
	Selects a file in an already opened file dialog.
	
	:param window_path: window path of the file dialog (e.g. "Untitled - Paint||Window->Save As||Window"
	:param full_path: the full path of the file to select
	:param force_slow_path_typing: if True it will type the path even if the current path of the dialog box is the same
	as the file to select
	"""
	import pyperclip
	import pathlib
	p = pathlib.Path(full_path)
	folder = p.parent
	filename = p.name
	with UIPath(window_path):
		wrapper = find()
		wait_is_ready_try1(wrapper)
		left_click(wrapper.descendants(title="All locations", control_type="SplitButton")[0])
		if not force_slow_path_typing:
			send_keys("{VK_CONTROL down}c{VK_CONTROL up}")
		if force_slow_path_typing or pathlib.Path(pyperclip.paste()) != folder:
			send_keys(str(folder))
		send_keys("{ENTER}")
		double_left_click(u"*->File name:||ComboBox->File name:||Edit")
		send_keys(filename + "{ENTER}")

