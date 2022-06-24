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
from functools import partial, update_wrapper


UI_Coordinates = NewType('UI_Coordinates', (float, float))
UI_Path = str
PYWINAUTO_Wrapper = pywinauto.controls.uiawrapper.UIAWrapper
UI_Selector = Union[UI_Path, PYWINAUTO_Wrapper, UI_Coordinates]


class PywinautoRecorderException(Exception):
	"""Base class for other exceptions."""
	pass


class FailedSearch(PywinautoRecorderException):
	"""FailedSearch is a subclass of *PywinautoRecorderException* that is raised when a search for a control fails."""
	pass


__all__ = ['PlayerSettings', 'MoveMode', 'ButtonLocation', 'load_dictionary', 'shortcut', 'full_definition', 'UIPath',
           'Window', 'Region', 'find', 'find_all', 'move', 'click', 'left_click', 'right_click', 'double_left_click',
           'triple_left_click', 'drag_and_drop', 'middle_drag_and_drop', 'right_drag_and_drop', 'menu_click',
           'mouse_wheel', 'send_keys', 'set_combobox', 'set_text', 'exists', 'select_file']


# TODO special_char_array in core for recorder.py and player.py (check when to call escape & unescape)
def unescape_special_char(string):
	for r in (("\\\\", "\\"), ("\\t", "\t"), ("\\n", "\n"), ("\\r", "\r"), ("\\v", "\v"), ("\\f", "\f"), ('\\"', '"')):
		string = string.replace(*r)
	return string


class PlayerSettings:
	"""The player settings class contains the mouse move duration and timeout."""
	mouse_move_duration = 0.5
	timeout = 10


class MoveMode(Enum):
	"""The MoveMode class is an enumeration of the different ways that the mouse can move"""
	linear = 0
	y_first = 1
	x_first = 2


class ButtonLocation(Enum):
	"""The ButtonLocation class is an enumeration of the different locations of the mouse buttons"""
	left = 0
	middle = 1
	right = 2


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


class UIPath(object):
	"""
	UIPath is a context manager used to keep track of the current path in the UI tree.
	
	.. code-block:: python
		:caption: Example of code not using a 'UIPath' object::
		:emphasize-lines: 3,4
		
		from pywinauto_recorder.player import click
		
		click("Calculator||Window->*->Number pad||Group->One||Button")
		click("Calculator||Window->*->Number pad||Group->Two||Button")
		
	The code above clicks on two buttons. On each line that corresponds to a click operation, the whole path is repeated.
	A UIPath object will allow to factorize a common path where several operations will be performed.
	
	The following code does the same as the previous example:
	
	.. code-block:: python
		:caption: Example of code using a 'UIPath' object::
		:emphasize-lines: 3,3
		
		from pywinauto_recorder.player import UIPath, click
		
		with UIPath("Calculator||Window"):
			click("*->Number pad||Group->One||Button")
			click("*->Number pad||Group->Two||Button")
	
	UIPath objects can be nested.
	
	The following code does the same as the previous example:
	
	.. code-block:: python
		:caption: Example of code using nested 'UIPath' objects::
		:emphasize-lines: 3,4
		
		from pywinauto_recorder.player import UIPath, click
		
		with UIPath("Calculator||Window"):
			with UIPath("*->Number pad||Group"):
				click("One||Button")
				click("Two||Button")

	Methods:
	
	Methods
	-------
	get_full_path(element_path: Optional[UI_Path] = None) -> UI_Path
		Returns the full path of the element.
	"""
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
		if element_path is None or element_path == "":
			return path_separator.join(UIPath._path_list)
		elif UIPath._path_list:
			return path_separator.join(UIPath._path_list) + path_separator + element_path
		else:
			return element_path
	
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
	Finds the element matching element_path.

	.. code-block:: python
		:caption: Example of code using the 'find' function::
		:emphasize-lines: 3,3
		
		from pywinauto_recorder.player import UIPath, find
		with UIPath("RegEx: .* Google Chrome$||Pane"):
			find().set_focus()  # Set focus to the Google Chrome window.
	
	The code above will set focus to the Google Chrome window. The 'find' function is used to find the window.
	It will work only if the window is not minimized.
	
	:param element_path: element path
	:param timeout: period of time in seconds that will be allowed to find the element
	:return: Pywinauto wrapper of found element
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
	
	if element_path is None or isinstance(element_path, str):
		full_element_path = UIPath.get_full_path(element_path)
	else:
		full_element_path = get_wrapper_path(element_path)
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


def find_all(
		element_path: Optional[UI_Selector] = None,
		timeout: Optional[float] = None) -> PYWINAUTO_Wrapper:
	"""
	Finds all elements matching element_path.

	.. code-block:: python
		:caption: Example of code using the 'find_all' function::
		:emphasize-lines: 4,4
		
		from pywinauto_recorder.player import UIPath, find, find_all
		with UIPath("RegEx: .* Google Chrome$||Pane"):
			find().set_focus()
			wrapper_tab_list = find_all("*->RegEx: .*||TabItem") # Find all tabs.
			for wrapper_tab in wrapper_tab_list:
				wrapper_tab.click_input()
				wrapper_url = find("*->Address and search bar||Edit")
				print(wrapper_url.get_value())
		
	The code above will click on all tabs of Google Chrome and print the URL of each tab.
	It will work only if the Google Chrome window is not minimized.
	The 'find_all' function is used to find all tabs.
	
	
	:param element_path: element path
	:param timeout: period of time in seconds that will be allowed to find the element
	:return: Pywinauto wrapper list of found elements
	"""
	if timeout is None:
		timeout = PlayerSettings.timeout
	
	if element_path is None or type(element_path) is str:
		full_element_path = UIPath.get_full_path(element_path)
	else:
		full_element_path = get_wrapper_path(element_path)
	entry_list = get_entry_list(full_element_path)
	_, _, y_x, _ = get_entry(entry_list[-1])
	if y_x:
		return find(element_path, timeout=timeout)
	unique_element = None
	elements = None
	t0 = time.time()
	while (time.time() - t0) < timeout:
		try:
			unique_element, elements = find_element(entry_list)
			if unique_element or elements:
				break
			if not unique_element and not elements:
				time.sleep(2.0)
		except Exception:
			pass
		if (time.time() - t0) > timeout:
			msg = "No element found with the UIPath '" + full_element_path + "' after " + str(timeout) + " s of searching."
			raise FailedSearch(msg)
		time.sleep(0.1)

	if unique_element:
		return [unique_element]
	else:
		return elements


def __move(x, y, xd, yd, duration=1, refresh_rate=25):
	"""
	It moves the mouse from (x, y) to (xd, yd) in a straight line, with a duration of `duration` seconds.
	
	:param x: The x-coordinate of the mouse cursor before the move
	:param y: The y-coordinate of the mouse cursor before the move
	:param xd: The x-coordinate of the mouse cursor after the move
	:param yd: The y-coordinate of the mouse cursor after the move
	:param duration: The time (in seconds) it takes to move the mouse from (x, y) to (xd, yd)
	:param refresh_rate: 25 Hz is the default refresh rate of the mouse move
	"""
	x_max = win32api_GetSystemMetrics(0) - 1
	y_max = win32api_GetSystemMetrics(1) - 1
	if duration > 0:
		samples = duration * refresh_rate
		dt = 1 / refresh_rate
		step_x = (xd - x) / samples
		step_y = (yd - y) / samples
		t0 = time.time()
		for i in range(int(samples)):
			x, y = x+step_x, y+step_y
			t1 = time.time()
			if t1-t0 > i*dt:
				continue
			time.sleep(dt)
			nx = int(x * 65535 / x_max)
			ny = int(y * 65535 / y_max)
			win32api_mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, nx, ny)
	nx = round(xd * 65535 / x_max)
	ny = round(yd * 65535 / y_max)
	win32api_mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, nx, ny)


def move(
		element_path: UI_Selector,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		timeout: float = 120) -> PYWINAUTO_Wrapper:
	"""
	Moves on element
	
	:param element_path: element path
	:param duration: duration in seconds of the mouse move (it doesn't take into account the time it takes to find),
		if duration is -1 the mouse cursor doesn't move.
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
	if (x, y) != (xd, yd):
		if mode == MoveMode.linear:
			__move(x, y, xd, yd, duration)
		elif mode == MoveMode.x_first:
			__move(x, y, xd, y, duration/2)
			__move(xd, y, xd, yd, duration/2)
		elif mode == MoveMode.y_first:
			__move(x, y, x, yd, duration/2)
			__move(x, yd, xd, yd, duration/2)
	if unique_element is None:
		return None
	unique_element_old = unique_element
	element_path_old = element_path2
	w_rOLD = w_r
	return unique_element


def click(
		element_path: Optional[UI_Selector] = None,
		duration: Optional[float] = None,
		mode: MoveMode = MoveMode.linear,
		button: ButtonLocation = ButtonLocation.left,
		click_count: int = 1,
		timeout: float = None,
		wait_ready: bool = True) -> PYWINAUTO_Wrapper:
	"""
	Clicks on found element.
	
	.. code-block:: python
		:caption: Example of code using the 'click' function::
		:emphasize-lines: 3,3
		
		from pywinauto_recorder.player import click, MoveMode
		
		click("Calculator||Window->*->One||Button", mode=MoveMode.x_first, duration=4)
	
	:param element_path: element path
	:param duration: duration in seconds of the mouse move (it doesn't take into account the time it takes to find)
		(if duration is -1 the mouse cursor doesn't move, it just sends WM_CLICK window message,
		useful for minimized or non-active window).
	:param mode: move mouse mode: MoveMode.linear, MoveMode.x_first, MoveMode.y_first
	:param button: mouse button:  ButtonLocation.left, ButtonLocation.middle, ButtonLocation.right
	:param click_count: number of clicks
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
	if button == ButtonLocation.left:
		event_down = MOUSEEVENTF_LEFTDOWN
		event_up = MOUSEEVENTF_LEFTUP
	elif button == ButtonLocation.middle:
		event_down = MOUSEEVENTF_MIDDLEDOWN
		event_up = MOUSEEVENTF_MIDDLEUP
	elif button == ButtonLocation.right:
		event_down = MOUSEEVENTF_RIGHTDOWN
		event_up = MOUSEEVENTF_RIGHTUP
	for _ in range(click_count):
		win32api_mouse_event(event_down, 0, 0)
		time.sleep(.01)
		win32api_mouse_event(event_up, 0, 0)
		time.sleep(.1)
	if element_path:
		return wrapper
	else:
		return None


def wrapped_partial(func, *args, **kwargs):
	partial_func = partial(func, *args, **kwargs)
	update_wrapper(partial_func, func)
	partial_func.__doc__ = "This function is a partial function derived from the '" + func.__name__ + "' general function."
	partial_func.__doc__ += "\nThe parameters of the function are set with the following values:"
	for key in kwargs.keys():
		partial_func.__doc__ += "\n    - " + str(key) + "=" + str(kwargs[key])
	partial_func.__doc__ += "\n\n" + func.__doc__
	return partial_func


left_click = wrapped_partial(click, button=ButtonLocation.left)
right_click = wrapped_partial(click, button=ButtonLocation.right)
double_left_click = wrapped_partial(click, button=ButtonLocation.left, click_count=2)
triple_left_click = wrapped_partial(click, button=ButtonLocation.left, click_count=3)


def drag_and_drop(
		element_path1: UI_Selector,
		element_path2: UI_Selector,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		button: ButtonLocation = ButtonLocation.left,
		timeout: Optional[float] = None) -> PYWINAUTO_Wrapper:
	"""
	Drags and drop with left button pressed from element_path1 to element_path2.
	
	:param element_path1: element path
	:param element_path2: element path
	:param duration: duration in seconds of the mouse move (it doesn't take into account the time it takes to find)
		(if duration is -1 the mouse cursor doesn't move, it just sends WM_CLICK window message,
		useful for minimized or non-active window).
	:param mode: move mouse mode: MoveMode.linear, MoveMode.x_first, MoveMode.y_first
	:param button: mouse button:  ButtonLocation.left, ButtonLocation.middle, ButtonLocation.right
	:param timeout: period of time in seconds that will be allowed to find the element
	:return: Pywinauto wrapper with element_path2
	"""
	move(element_path1, duration=duration, mode=mode, timeout=timeout)
	if button == ButtonLocation.left:
		event_down = MOUSEEVENTF_LEFTDOWN
		event_up = MOUSEEVENTF_LEFTUP
	elif button == ButtonLocation.middle:
		event_down = MOUSEEVENTF_MIDDLEDOWN
		event_up = MOUSEEVENTF_MIDDLEUP
	elif button == ButtonLocation.right:
		event_down = MOUSEEVENTF_RIGHTDOWN
		event_up = MOUSEEVENTF_RIGHTUP
	win32api_mouse_event(event_down, 0, 0)
	unique_element = move(element_path2, duration=duration, timeout=timeout)
	win32api_mouse_event(event_up, 0, 0)
	return unique_element


left_drag_and_drop = wrapped_partial(drag_and_drop, button=ButtonLocation.left)
middle_drag_and_drop = wrapped_partial(drag_and_drop, button=ButtonLocation.middle)
right_drag_and_drop = wrapped_partial(drag_and_drop, button=ButtonLocation.right)


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
	:param duration: duration in seconds of the mouse move (it doesn't take into account the time it takes to find)
		(if duration is -1 the mouse cursor doesn't move, it just sends WM_CLICK window message,
		useful for minimized or non-active window).
	:param mode: move mouse mode: MoveMode.linear, MoveMode.x_first, MoveMode.y_first
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
	:param duration: duration in seconds of the mouse move (it doesn't take into account the time it takes to find)
		(if duration is -1 the mouse cursor doesn't move, it just sends WM_CLICK window message,
		useful for minimized or non-active window).
	:param mode: move mouse mode: MoveMode.linear, MoveMode.x_first, MoveMode.y_first
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
	:param duration: duration in seconds of the mouse move (it doesn't take into account the time it takes to find)
		(if duration is -1 the mouse cursor doesn't move, it just sends WM_CLICK window message,
		useful for minimized or non-active window).
	:param mode: move mouse mode: MoveMode.linear, MoveMode.x_first, MoveMode.y_first
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
	except FailedSearch:
		return None


def select_file(
		window_path: UI_Selector,
		full_path: str,
		force_slow_path_typing: bool = False) -> None:
	"""
	Selects a file in an already opened file dialog.
	
	.. code-block:: python
		:caption: Example of code using 'select_file'::
		:emphasize-lines: 3,3
		
		from pywinauto_recorder.player import select_file
		
		select_file("Document - WordPad||Window->Open||Window", "Documents/file.txt")
	
	To make this code work, you must first launch 'WordPad' and click on 'File->Open'.

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
		find().set_focus()
		click("*->All locations||SplitButton")
	if not force_slow_path_typing:
		send_keys("{VK_CONTROL down}c{VK_CONTROL up}")
	if force_slow_path_typing or pathlib.Path(pyperclip.paste()) != folder:
		send_keys(str(folder))
	send_keys("{ENTER}")
	with UIPath(window_path):
		double_left_click("*->File name:||ComboBox->File name:||Edit")
		send_keys(filename + "{ENTER}")

