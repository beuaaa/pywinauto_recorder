"""This module contains functions to replay a sequence of user actions automatically."""

# print('__file__={0:<35} | __name__={1:<20} | __package__={2:<20}'.format(__file__, __name__, str(__package__)))
from enum import Enum
from typing import Optional, Union, NewType
import time
import re
import pathlib
import pywinauto
from win32api import GetCursorPos as win32api_GetCursorPos
from win32api import GetSystemMetrics as win32api_GetSystemMetrics
from win32api import mouse_event as win32api_mouse_event
from win32gui import LoadCursor as win32gui_LoadCursor
from win32gui import GetCursorInfo as win32gui_GetCursorInfo
from win32con import IDC_WAIT, MOUSEEVENTF_MOVE, MOUSEEVENTF_ABSOLUTE, MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP, \
	MOUSEEVENTF_MIDDLEDOWN, MOUSEEVENTF_MIDDLEUP, MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP, MOUSEEVENTF_WHEEL, \
	WHEEL_DELTA
from .core import type_separator, path_separator, get_entry, get_entry_list, find_elements, get_sorted_region, \
	get_wrapper_path, is_int, is_absolute_path
from functools import partial, update_wrapper
from cachetools import func
import math
from .ocr_wrapper import OCRWrapper


UI_Coordinates = NewType('UI_Coordinates', (float, float))
UI_Path = str
PYWINAUTO_Wrapper = pywinauto.controls.uiawrapper.UIAWrapper
UI_Selector = Union[UI_Path, PYWINAUTO_Wrapper, UI_Coordinates]


class PywinautoRecorderException(Exception):
	"""Base class for other exceptions."""
	...


class FailedSearch(PywinautoRecorderException):
	"""FailedSearch is a subclass of *PywinautoRecorderException* that is raised when a search for a control fails."""
	...


__all__ = ['PlayerSettings', 'MoveMode', 'ButtonLocation', 'load_dictionary', 'shortcut', 'full_definition', 'UIPath',
           'Window', 'Region', 'find', 'find_all', 'move', 'click', 'left_click', 'right_click',
           'double_left_click', 'triple_left_click', 'double_click', 'triple_click',
           'drag_and_drop', 'middle_drag_and_drop', 'right_drag_and_drop', 'menu_click',
           'mouse_wheel', 'send_keys', 'set_combobox', 'set_text', 'exists', 'select_file', 'playback',
           'find_cache_clear']


# TODO special_char_array in core for recorder.py and player.py (check when to call escape & unescape)
def unescape_special_char(string):
	for r in (("\\\\", "\\"), ("\\t", "\t"), ("\\n", "\n"), ("\\r", "\r"), ("\\v", "\v"), ("\\f", "\f"), ('\\"', '"')):
		string = string.replace(*r)
	return string


class PlayerSettings:
	"""The player settings class contains the default settings."""
	
	typing_pause = 0.1
	"""The pause time between characters typed"""
	
	mouse_move_duration = 0.5
	"""Mouse move duration (in seconds)."""
	
	timeout = 10
	"""Maximum duration (in seconds) to wait for the :func:`find` function to search an element before giving up.
	If the element is not found after the given timeout, the search is interrupted."""

	use_cache = True
	"""If True, the :func:`find` function caches the results of the search.
	This is useful if the search is called several times on the same element."""
	
	@staticmethod
	def _apply_settings(
			typing_pause: Optional[float] = None,
			mouse_move_duration: Optional[float] = None,
			timeout: Optional[float] = None) -> dict:
		"""
		If the duration and timeout arguments are None, set them to the default values.
		
		:param typing_duration: The pause time between characters typed
		:param mouse_move_duration: The duration of the mouse movement
		:param timeout: The maximum duration to wait for the :func:`find` function to find an element before giving up
		:return: The duration and timeout are being returned.
		"""
		if typing_pause is None:
			typing_pause = PlayerSettings.typing_pause
		if mouse_move_duration is None:
			mouse_move_duration = PlayerSettings.mouse_move_duration
		if timeout is None:
			timeout = PlayerSettings.timeout
		return {"typing_pause": typing_pause, "mouse_move_duration": mouse_move_duration, "timeout": timeout}


class MoveMode(Enum):
	"""The MoveMode class is an enumeration of the different ways that the mouse can move."""
	
	linear = 0
	"""The mouse cursor moves in a straight line from the start point to the end point with a constant speed."""
	
	y_first = 1
	"""The mouse cursor moves at a right angle from the starting point to the ending point at a constant speed.
	The first segment of the right angle is vertical and the second is horizontal."""
	
	x_first = 2
	"""The mouse cursor moves at a right angle from the starting point to the ending point at a constant speed.
	The first segment of the right angle is horizontal and the second is vertical."""


class ButtonLocation(Enum):
	"""The ButtonLocation class is an enumeration of the different locations of the mouse buttons."""
	
	left = 0
	"""Left mouse button."""
	
	middle = 1
	"""Middle mouse button."""
	
	right = 2
	"""Right mouse button."""


_dictionary = {}
unique_element_old = None
element_path_old = ''
w_rOLD = None


def load_dictionary(filename_key: str, filename_def: str,encoding: str = 'utf8') -> None:
	"""
	Loads a dictionary.

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
	Returns the shortcut path associated to the shortcut defined in the previously loaded dictionary.

	:param str_shortcut: shortcut
	"""
	return _dictionary[str_shortcut][1]


def full_definition(str_shortcut: str) -> str:
	"""
	Returns the full element path associated to the shortcut defined in the previously loaded dictionary.

	:param str_shortcut: shortcut
	"""
	return _dictionary[str_shortcut][0] + path_separator + _dictionary[str_shortcut][1]


def wait_is_ready_try1(wrapper, timeout=120):
	"""
	Waits until element is ready (wait while greyed, not enabled, not visible, not ready, ...) :
	So far, I didn't find better than wait_cpu_usage_lower when greyed but must be enhanced.
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
	"""
	_path_list = []
	_regex_list = []  # UIPath._regex_list must be removed
	
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
		elif UIPath._path_list and not is_absolute_path(element_path):
			return path_separator.join(UIPath._path_list) + path_separator + element_path
		else:
			return element_path
	
	def __init__(self, relative_path=None, regex_title=False):
		self.relative_path = relative_path
		self.regex_title = regex_title  # UIPath._regex_list must be removed
	
	def __enter__(self):
		if self.relative_path:
			UIPath._path_list.append(self.relative_path)
			UIPath._regex_list.append(self.regex_title)  # UIPath._regex_list must be removed
		return self
	
	def __exit__(self, type, value, traceback):
		if self.relative_path:
			UIPath._path_list = UIPath._path_list[0:-1]
			UIPath._regex_list = UIPath._regex_list[0:-1]  # UIPath._regex_list must be removed


Window = UIPath
Region = UIPath


def find_cache_clear():
	"""
	Clears the cache of the :func:`find` function.
	"""
	_cached_find.cache_clear()


@func.ttl_cache(ttl=60)
def _cached_find(
		full_element_path: Optional[UI_Selector] = None,
		timeout: Optional[float] = None) -> PYWINAUTO_Wrapper:
	"""
	Finds the element defined by the full_element_path.
	
	full_element_path must not contain the relative coordinates.
	"""
	return _find(full_element_path, timeout)


def _find(
		full_element_path: Optional[UI_Selector] = None,
		timeout: Optional[float] = None) -> PYWINAUTO_Wrapper:
	"""
	Finds the element defined by the full_element_path.
	
	When the [] operator is used and only one element is found, the row and column indices are not tested and the element is returned.
	"""
	_, _, y_x, _ = get_entry(get_entry_list(full_element_path)[-1])
	elements = []
	t0 = time.time()
	while (time.time() - t0) < timeout:
		while not elements:
			if (time.time() - t0) > timeout:
				msg = "No element found with the UIPath '" + full_element_path + "' after " + str(timeout) + " s of searching."
				raise FailedSearch(msg)
			try:
				elements = find_elements(full_element_path)
				if not elements:
					#print("🔴", end="")
					time.sleep(2.0)
			except Exception:
				#print("🟢", end="")
				pass

		if len(elements) == 1:
				return elements[0]
		
		if y_x is not None:
			nb_y, _, candidates = get_sorted_region(elements)
			if is_int(y_x[0]):
				return candidates[int(y_x[0])][int(y_x[1])]
			else:
				full_smart_element_path = UIPath.get_full_path(y_x[0])
				ref_unique_element = find_elements(full_smart_element_path)
				if len(ref_unique_element) > 1:
					msg = "No element found with the UIPath '" + full_smart_element_path + "' in the array line."
					raise FailedSearch(msg)
				ref_r = ref_unique_element[0].rectangle()
				r_y = 0
				while r_y < nb_y:
					for candidate in candidates[r_y]:
						y_candidate = candidate.rectangle().mid_point()[1]
						if ref_r.top < y_candidate < ref_r.bottom:
							return candidates[r_y][y_x[1]]
					r_y = r_y + 1
		time.sleep(0.1)

	if len(elements) > 1:
		message = "There are " + str(len(elements)) + " undiscriminated elements that match the path '" + full_element_path + "':"
		for e in elements:
			if isinstance(e, OCRWrapper):
				message += "\n" + str(e.result)
			else:
				message += "\n" + get_wrapper_path(e)
		raise FailedSearch(message)
	raise FailedSearch("Unique element not found using path '", full_element_path + "'")


def find(
		element_path: Optional[UI_Selector] = None,
		regex: bool = False,
		timeout: Optional[float] = None) -> PYWINAUTO_Wrapper:
	"""
	Finds the element matching element_path.
	
	This function is called in all the other functions (:func:`click`, :func:`move`, ...) that require to search an element.
	To significantly increase search performance, the user can enable a cache with 'Player.Setting.use_cache = True'.
	When the cache is active, it is sometimes necessary to empty it with the :func:`find_cache_clear` function.

	When the [] operator is used and only one element is found, the row and column indices are not tested and the element is returned.

	.. code-block:: python
		:caption: Example of code using the 'find' function::
		:emphasize-lines: 3,3
		
		from pywinauto_recorder.player import UIPath, find
		with UIPath("RegEx: .* Google Chrome$||Pane"):
			find().set_focus()  # Set focus to the Google Chrome window.
	
	The code above will set focus to the Google Chrome window.
	The :func:`find` function is used to find the Pywinauto wrapper of the window.
	It will work only if the window is not minimized.
	
	:param element_path: element path
	:param regex: The parameter 'regex' is deprecated. Please use the new RegEx syntax of UIPath!
	:param timeout: duration in seconds that will be allowed to find the element
	:return: Pywinauto wrapper of found element
	:raises FailedSearch: if the element is not found
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
		
	timeout = PlayerSettings._apply_settings(timeout=timeout)["timeout"]
	
	if element_path is None or isinstance(element_path, str):
		if element_path is not None:
			element_path = re.sub(r"%\([+-]?\d*.?\d*,\s?[+-]?\d*.?\d*\)$", "", element_path)  # remove "%(?, ?)"
		full_element_path = UIPath.get_full_path(element_path)
	else:
		full_element_path = get_wrapper_path(element_path)
	if PlayerSettings.use_cache:
		return _cached_find(full_element_path, timeout=timeout)
	else:
		return _find(full_element_path, timeout=timeout)


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
	The :func:`find_all` function is used to find all tabs.
	
	
	:param element_path: element path
	:param timeout: period of time in seconds that will be allowed to find the element
	:return: Pywinauto wrapper list of found elements
	:raises FailedSearch: if no element found
	"""
	timeout = PlayerSettings._apply_settings(timeout=timeout)["timeout"]
	if element_path is None or isinstance(element_path, str):
		full_element_path = UIPath.get_full_path(element_path)
	else:
		full_element_path = get_wrapper_path(element_path)
	entry_list = get_entry_list(full_element_path)
	_, _, y_x, _ = get_entry(entry_list[-1])
	if y_x:
		return find(element_path, timeout=timeout)
	t0 = time.time()
	while (time.time() - t0) < timeout:
		try:
			elements = find_elements(full_element_path)
			if elements:
				return elements
			time.sleep(2.0)
		except Exception:
			pass
		time.sleep(0.1)
	msg = "No element found with the UIPath '" + full_element_path + "' after " + str(timeout) + " s of searching."
	raise FailedSearch(msg)


def _move(x, y, xd, yd, duration=1, refresh_rate=25):
	"""
	It moves the mouse from (x, y) to (xd, yd) in a straight line, with a duration of duration seconds.
	
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
		element_path: Optional[UI_Selector] = None,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		timeout: float = 120) -> PYWINAUTO_Wrapper:
	"""
	Moves the mouse cursor over the user interface element.
	
	:param element_path: element path
	:param duration: duration in seconds of the mouse move (it doesn't take into account the time it takes to find),
		if duration is -1 the mouse cursor doesn't move.
	:param mode: move mouse mode (see :class:`MoveMode`)
	:param timeout: period of time in seconds that will be allowed to find the element
	:return: Pywinauto wrapper of clicked element
	:raises FailedSearch: if the element is not found
	"""
	duration = PlayerSettings._apply_settings(mouse_move_duration=duration)["mouse_move_duration"]
	if duration == -1:
		return
	
	if element_path is None or isinstance(element_path, str):
		unique_element = find(element_path, timeout=timeout)
		try:
			w_r = unique_element.rectangle()
		except Exception:
			#with _cached_find.cache_lock:
			#	_cached_find.cache.pop(_cached_find.cache_key(element_path, timeout=timeout), None)
			find_cache_clear()
			unique_element = find(element_path, timeout=timeout)
			w_r = unique_element.rectangle()
		xd, yd = w_r.mid_point()
		if element_path:
			_, _, _, dx_dy = get_entry(get_entry_list(element_path)[-1])
			if dx_dy:
				dx, dy = dx_dy
				xd, yd = round(xd + dx/100.0*(w_r.width()/2-1), 0), round(yd + dy/100.0*(w_r.height()/2-1), 0)
	elif isinstance(element_path, pywinauto.base_wrapper.BaseWrapper):
		unique_element = element_path
		w_r = unique_element.rectangle()
		xd, yd = w_r.mid_point()
	else:
		(xd, yd) = element_path
		unique_element = None
	x, y = win32api_GetCursorPos()
	if (x, y) != (xd, yd):
		if mode == MoveMode.linear:
			_move(x, y, xd, yd, duration)
		elif mode == MoveMode.x_first:
			_move(x, y, xd, y, duration / 2)
			_move(xd, y, xd, yd, duration / 2)
		elif mode == MoveMode.y_first:
			_move(x, y, x, yd, duration / 2)
			_move(x, yd, xd, yd, duration / 2)
	return unique_element


def _win32api_mouse_click(button: ButtonLocation = ButtonLocation.left, click_count: int = 1):
	"""
	Clicks the mouse.
	
	:param button: The button to click
	:param click_count: How many times to click, defaults to 1
	"""
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
	:raises FailedSearch: if the element is not found
	"""
	settings = PlayerSettings._apply_settings(mouse_move_duration=duration, timeout=timeout)
	duration = settings["mouse_move_duration"]
	timeout = settings["timeout"]
	
	if duration == -1:
		wrapper = find(element_path)
		has_get_value = getattr(wrapper, "click", None)
		if callable(has_get_value):
			wrapper.click()
		else:
			wrapper.click_input()
		return wrapper
	else:
		
		
		if wait_ready and \
			(element_path is None or isinstance(element_path, str) or isinstance(element_path, pywinauto.base_wrapper.BaseWrapper)):
			use_cache_old = PlayerSettings.use_cache
			PlayerSettings.use_cache = False
			if isinstance(element_path, pywinauto.base_wrapper.BaseWrapper):
				wrapper = element_path
			else:
				wrapper = move(element_path, duration=duration, mode=mode, timeout=timeout)
			wait_is_ready_try1(wrapper, timeout=timeout)
			PlayerSettings.use_cache = use_cache_old
		else:
			wrapper = move(element_path, duration=duration, mode=mode, timeout=timeout)
	_win32api_mouse_click(button, click_count)
	return wrapper


def wrapped_partial(func, *args, **kwargs):
	partial_func = partial(func, *args, **kwargs)
	update_wrapper(partial_func, func)
	partial_func.__doc__ = "This function is a partial function derived from the :func:`" + func.__name__
	partial_func.__doc__ += "` general function.\nThe parameters of the function are set with the following values:"
	for key in kwargs.keys():
		partial_func.__doc__ += "\n    - " + str(key) + "=" + str(kwargs[key])
	partial_func.__doc__ += "\n\n" + func.__doc__
	return partial_func


left_click = wrapped_partial(click, button=ButtonLocation.left)
right_click = wrapped_partial(click, button=ButtonLocation.right)
double_left_click = wrapped_partial(click, button=ButtonLocation.left, click_count=2)
double_click = wrapped_partial(click, button=ButtonLocation.left, click_count=2)
triple_left_click = wrapped_partial(click, button=ButtonLocation.left, click_count=3)
triple_click = wrapped_partial(click, button=ButtonLocation.left, click_count=3)


def drag_and_drop(
		element_path1: UI_Selector,
		element_path2: UI_Selector,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		button: ButtonLocation = ButtonLocation.left,
		timeout: Optional[float] = None) -> PYWINAUTO_Wrapper:
	"""
	Drags and drops from element_path1 to element_path2.
	
	:param element_path1: source element path
	:param element_path2: destination element path
	:param duration: duration in seconds of the mouse move (it doesn't take into account the time it takes to find)
		(if duration is -1 the mouse cursor doesn't move, it just sends WM_CLICK window message,
		useful for minimized or non-active window).
	:param mode: move mouse mode: MoveMode.linear, MoveMode.x_first, MoveMode.y_first
	:param button: mouse button:  ButtonLocation.left, ButtonLocation.middle, ButtonLocation.right
	:param timeout: period of time in seconds that will be allowed to find the element
	:return: Pywinauto wrapper found with element_path2
	:raises FailedSearch: if the element is not found
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
		menu_path: str,
		duration: Optional[float] = None,
		timeout: Optional[float] = None,
		absolute_path = False) -> PYWINAUTO_Wrapper:
	"""
	Clicks on the menu items.
	
	.. code-block:: python
		:caption: Example of code using the 'menu_click' function to automate 'Notepad++'::
		:emphasize-lines: 5,5
		
		from pywinauto_recorder.player import UIPath, find, click, menu_click
	
		with UIPath("RegEx: .* - Notepad\+\+$||Window"):
			find().set_focus()
			menu_click("Language->P->Python")
		
	.. code-block:: python
		:caption: Example of code using the 'menu_click' function to automate 'WordPad'::
		:emphasize-lines: 6,6
		
		from pywinauto_recorder.player import UIPath, find, click, menu_click
		
		with UIPath("Document - WordPad||Window"):
			find().set_focus()
			click("*->File tab||Button")
			menu_click("New")
			
	In the above code, the 'File tab' element that opens the menu is not of type 'MenuItem', so it is not possible to call 'menu_click("File tab->New")'. In this case the menu must be opened with 'click("*->File tab||Button")' before calling :func:`menu_click`.
	
	:param menu_path: menu item path
	:param duration: duration in seconds of the mouse move (it doesn't take into account the time it takes to find)
		(if duration is -1 the mouse cursor doesn't move, it just sends WM_CLICK window message,
		useful for minimized or non-active window).
	:param timeout: period of time in seconds that will be allowed to find the element
	:return: Pywinauto wrapper of the clicked item
	:raises FailedSearch: if an element is not found
	"""
	
	def nearest_perimeter_point(rect, point):
		x1, y1, x2, y2 = rect.left, rect.top, rect.right, rect.bottom
		x, y = point
		if x1 <= x <= x2 and y1 <= y <= y2:
			return point
		if x < x1:
			x = x1
		elif x > x2:
			x = x2
		if y < y1:
			y = y1
		elif y > y2:
			y = y2
		return x, y
	
	def distance(pt_1, pt_2):
		return math.hypot(pt_2[0] - pt_1[0], pt_2[1] - pt_2[1])
	
	if duration not in [None, -1]:
		duration = float(duration)/2
		
	menu_entry_list = menu_path.split(path_separator)
	str_menu_item = 'MenuItem~Absolute_UIPath' if absolute_path else 'MenuItem'
	for i, menu_entry in enumerate(menu_entry_list):
		if i > 0:
			SAV_UIPath_path_list, SAV_UIPath_regex_list = UIPath._path_list, UIPath._regex_list
			UIPath._path_list = UIPath._regex_list = []
		mouse_cursor_pos = win32api_GetCursorPos()
		ws = find_all('*' + path_separator + menu_entry + type_separator + str_menu_item, timeout=timeout)
		ws.sort(key=lambda w: distance(w.rectangle().mid_point(), mouse_cursor_pos))
		w = ws[0]
		item_pos = w.rectangle().mid_point()
		r = w.parent().rectangle()
		nearest_p_point = nearest_perimeter_point(r, mouse_cursor_pos)
		move(nearest_p_point, duration=duration)
		click(item_pos, duration=duration)
		time.sleep(0.1)  # wait for the menu to open (it is not always instantaneous depending on the animation settings)
		if i>0:
			UIPath._path_list, UIPath._regex_list = SAV_UIPath_path_list, SAV_UIPath_regex_list
	return w


def mouse_wheel(steps: int, pause: float = 0.05) -> None:
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
		pause: Optional[float] = None,
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
	typing_pause = PlayerSettings._apply_settings(typing_pause=pause)["typing_pause"]
	
	for r in (('(', '{(}'),  (')', '{)}'), ('+', '{+}')):
		str_keys = str_keys.replace(*r)
	pywinauto.keyboard.send_keys(   # lgtm [py/call/wrong-named-argument]
		str_keys,
		pause=typing_pause,
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
	:raises FailedSearch: if the element is not found
	"""
	click(element_path, duration=duration, mode=mode, timeout=timeout, wait_ready=wait_ready)
	time.sleep(0.9)
	send_keys(value + "{ENTER}")


def set_text(
		element_path: UI_Selector,
		value: str,
		duration: Optional[float] = None,
		mode: Enum = MoveMode.linear,
		timeout: Optional[float] = None,
		pause: float = None) -> None:
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
	:raises FailedSearch: if the element is not found
	"""
	typing_pause = PlayerSettings._apply_settings(typing_pause=pause)["typing_pause"]
	double_left_click(element_path, duration=duration, mode=mode, timeout=timeout)
	send_keys("{VK_CONTROL down}a{VK_CONTROL up}", pause=0)
	time.sleep(0.1)
	send_keys(value + "{ENTER}", pause=typing_pause)


def exists(
		element_path: UI_Selector,
		timeout: Optional[float] = None) -> PYWINAUTO_Wrapper:
	"""
	Tests if the user interface element exists. It returns the element if it exists, None otherwise.
	In both cases no exception is thrown as for the :func:`find` function.
	
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
	:param force_slow_path_typing: if True it will type the path even if the current path of the dialog box is the same as the file to select
	:raises FailedSearch: if an element is not found
	"""
	p = pathlib.Path(full_path)
	folder = p.parent
	filename = p.name
	with UIPath(window_path):
		find().set_focus()
		click("*->All locations||SplitButton")
	if not force_slow_path_typing:
		with UIPath(window_path):
			try:
				old_folder = find("*->Address||ComboBox->Address||Edit").get_value()
			except Exception:
				find_cache_clear()
				#with _cached_find.cache_lock:
				#	_cached_find.cache.pop(_cached_find.cache_key("*->Address||ComboBox->Address||Edit"), None)
				old_folder = find("*->Address||ComboBox->Address||Edit").get_value()
			
	if force_slow_path_typing or old_folder != folder:
		send_keys(str(folder))
	send_keys("{ENTER}")
	with UIPath(window_path):
		double_left_click("*->File name:||ComboBox->File name:||Edit")
		send_keys(filename + "{ENTER}")


def playback(str_code='', filename=''):
	"""
	This function plays back a string of code or a Python file.

	:param str_code: The Python code to be played back
	:param filename: The name of the file corresponding to the Python code to be played back
	"""
	from ctypes import windll
	import traceback
	import os
	import sys
	import codecs
	
	if str_code == '' and os.path.isfile(filename):
		with codecs.open(filename, "r", encoding='utf-8') as python_file:
			str_code = python_file.read()
	try:
		script_dir = os.path.abspath(os.path.dirname(filename))
		os.chdir(os.path.abspath(script_dir))
		sys.path.append(script_dir)
		compiled_code = compile(str_code, filename, 'exec')
		exec(compiled_code)
	except Exception:
		windll.user32.ShowWindow(windll.kernel32.GetConsoleWindow(), 3)
		exc_type, exc_value, exc_traceback = sys.exc_info()
		output = traceback.format_exception(exc_type, exc_value, exc_traceback)
		i_line = d_line = 0
		full_traceback = False
		if not full_traceback:
			for line in output:
				i_line += 1
				if "pywinauto_recorder.py" in line:
					d_line = i_line
		
		for line in output[d_line:]:
			print(line, file=sys.stderr, end='')
		input("Press Enter to continue...")
