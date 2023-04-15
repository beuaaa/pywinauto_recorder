# -*- coding: utf-8 -*-

import re
from enum import Enum
import configparser
import ast
from pywinauto import Desktop as PywinautoDesktop
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto import findwindows
from thefuzz import fuzz
from .ocr_wrapper import find_all_ocr, OCRWrapper


__all__ = ['path_separator', 'type_separator', 'Strategy', 'is_int', 'is_absolute_path',
           'get_wrapper_path', 'get_entry_list', 'get_entry', 'match_entry_list', 'get_sorted_region',
           'find_elements', 'read_config_file']

desktop = PywinautoDesktop(backend='uia', allow_magic_lookup=False)


class CoreSettings:
	class window_filtering:
		mode = 'ignore_windows'
		admit_windows = []
		ignore_windows = []


path_separator = "->"
type_separator = "||"


# > The `Strategy` class is an enumeration of the different strategies that can be used to solve the problem
class Strategy(Enum):
	unique_path = 1
	array_1D = 2
	array_2D = 3


def get_wrapper_path(wrapper):
	"""
	It takes a UI Automation wrapper object and returns a string that represents the path to the element from the root of
	the UI Automation tree
	
	:param wrapper: The UIAutomation wrapper object
	:return: The path of the wrapper from the top level parent to the wrapper.
	"""
	try:
		path = ''
		wrapper_top_level_parent = wrapper.top_level_parent()
		while wrapper != wrapper_top_level_parent:
			path = path_separator + wrapper.element_info.name + type_separator + wrapper.element_info.control_type + path
			wrapper = wrapper.parent()
		return wrapper.element_info.name + type_separator + wrapper.element_info.control_type + path
	except Exception:
		return ''


def get_entry_list(path):
	"""
	It splits the path into a list of entries
	It only handles one #[y,x] at the end. It does not handle the #[y,x] in the middle.
	
	:param path: the path to the element
	:return: A list of the entries.
	"""
	i = path.rfind("#[")
	if i != -1:
		i = path.rfind(path_separator, 0, i)
	else:
		i = path.rfind(path_separator)
	if i == -1:
		return [path]
	last_entry = path[i + len(path_separator)::]
	start_entry = path[:i]
	return start_entry.split(path_separator) + [last_entry]


def is_int(s):
	"""
	"If the string can be converted to an integer, return True, otherwise return False."

	The try statement is used to catch exceptions. In this case, if the string can be converted to an integer, the int()
	function will return the integer. If the string cannot be converted to an integer, the int() function will raise a
	ValueError exception

	:param s: the string to check
	:return: True or False
	"""
	try:
		int(s)
		return True
	except ValueError:
		return False


def is_absolute_path(entry):
	return entry[-16:] == '~Absolute_UIPath'


def get_entry(entry):
	"""
	It takes a string of the form "name"#["y",x]%(dx,dy)~Absolute_UIPath and returns a tuple of the form (name, type, [y,x], (dx,dy))
	
	:param entry: the string that is the entry in the listbox
	:return: The name of the entry, the type of the entry, the size of the entry, and the dx and dy of the entry.
	"""
	if is_absolute_path(entry):
		entry = entry[:-16]
	if entry[0:7] == "RegEx: ":
		entry = entry[7:]
	i = entry.find(type_separator)
	if i == -1:
		return entry, None, None, None
	while i < len(entry) and entry[i] == type_separator[0]:
		i = i + 1
	i = i - len(type_separator)
	str_name = entry[0:i]
	
	entry2 = entry[i + len(type_separator):]
	i = entry2.rfind("""%(""")
	
	if i != -1:
		str_dx_dy = entry2[i + 2:]
		entry3 = entry2[0:i]
	else:
		str_dx_dy = None
		entry3 = entry2
	
	i = entry3.rfind("""#[""")
	if i != -1:
		str_array = entry3[i + 2:-1]
		str_type = entry3[0:i]
	else:
		str_array = None
		str_type = entry3
	
	if str_type == '':
		str_type = None
	if str_array:
		words = str_array.split(',')
		y = words[0]
		if is_int(y):
			y = int(y)
		x = int(words[1])
		y_x = [y, x]
	else:
		y_x = None
	if str_dx_dy:
		words = str_dx_dy.split(',')
		dx_dy = (float(words[0]), float(words[1].split(')')[0]))
	else:
		dx_dy = None
	return str_name, str_type, y_x, dx_dy


def is_regex_entry(entry):
	"""
	It returns True if the first seven characters of the string entry are "RegEx: " and False otherwise
	
	:param entry: The entry to check
	:return: A boolean value.
	"""
	return entry[0:7] == "RegEx: "


def match_entry(entry, template):
	"""
	It returns true if the entry matches the template, and false otherwise
	
	:param entry: the entry to be matched
	:param template: the template to match against
	:return: A boolean value.
	"""
	# print("match_entry: entry: " + entry + " template: " + template)
	template_name, template_type, _, _ = get_entry(template)
	entry_name, entry_type, _, _ = get_entry(entry)
	if is_regex_entry(template):
		return re.match(template_name, entry_name) and (template_type == entry_type or not template_type)
	else:
		return template_name == entry_name and (template_type == entry_type or not template_type)


def match_entry_list(l1, l2):
	"""
	It takes an entry list and a template list, and returns True if the entry list matches the template list

	:param l1: the entry list of an element
	:param l2: the template list (already completed with UIPath)
	:return: True if the entry list matches the template list.
	"""
	if (l1 == []):
		return (l2 == [] or l2 == ['*'])
	if (l2 == [] or l2[0] == '*'):
		return match_entry_list(l2, l1)
	if (l1[0] == '*'):
		return (match_entry_list(l1, l2[1:]) or match_entry_list(l1[1:], l2))
	if is_regex_entry(l1[0]):
		if match_entry(l2[0], l1[0]):
			return match_entry_list(l1[1:], l2[1:])
	elif match_entry(l1[0], l2[0]):
		return match_entry_list(l1[1:], l2[1:])
	else:
		return False


def is_filter_criteria_ok(child, min_height=8, max_height=200, min_width=8, max_width=800):
	"""
	"Return True if the child is visible and its height and width are within the specified ranges."
	
	The function is_filter_criteria_ok() takes four arguments:
	
	child: the child to check
	min_height: the minimum height of the child
	max_height: the maximum height of the child
	min_width: the minimum width of the child
	max_width: the maximum width of the child
	The function returns True if the child is visible and its height and width are within the specified ranges. Otherwise,
	it returns False
	
	:param child: the child element to check
	:param min_height: The minimum height of the element, defaults to 8 (optional)
	:param max_height: The maximum height of the element, defaults to 200 (optional)
	:param min_width: The minimum width of the element, defaults to 8 (optional)
	:param max_width: The maximum width of the element, defaults to 800 (optional)
	:return: A list of all the visible elements that meet the filter criteria.
	"""
	if child.is_visible():
		r = child.rectangle()
		h = r.height()
		w = r.width()
		if (min_height <= h <= max_height) and (min_width <= w <= max_width):
			return True
	return False


def all_height_equal(iterator):
	"""
	It returns True if all the elements in the iterator have the same height, and False otherwise
	
	:param iterator: an iterable object
	:return: A boolean value.
	"""
	iterator = iter(iterator)
	try:
		first = next(iterator)
		first_height = first.rectangle().height()
	except StopIteration:
		return True
	for x in iterator:
		if first_height != x.rectangle().height():
			return False
	return True


def get_sorted_region(elements, min_height=8, max_height=9999, min_width=8, max_width=9999, line_tolerance=20):
	"""
	It takes a list of elements, filters them based on the given criteria, sorts them by top and left, and then returns a
	list of lists of elements, where each list is a row of elements
	
	:param elements: the list of elements to be sorted
	:param min_height: The minimum height of the elements to be considered, defaults to 8 (optional)
	:param max_height: The maximum height of the elements in the region, defaults to 9999 (optional)
	:param min_width: The minimum width of the elements to be considered, defaults to 8 (optional)
	:param max_width: The maximum width of the elements in the region, defaults to 9999 (optional)
	:param line_tolerance: This is the maximum distance between two elements in the same row, defaults to 20 (optional)
	:return: The number of rows, the number of columns, and the sorted elements.
	"""
	filtered_elements = list(filter(
		lambda e: is_filter_criteria_ok(e, min_height, max_height, min_width, max_width), elements))
	filtered_elements.sort(key=lambda widget: (widget.rectangle().top, widget.rectangle().left))
	arrays = [[]]
	
	if all_height_equal(filtered_elements):
		line_tolerance = filtered_elements[0].rectangle().height()
	
	h = 0
	w = -1
	if len(filtered_elements) > 0:
		y = filtered_elements[0].rectangle().top
		for e in filtered_elements:
			if (e.rectangle().top - y) < line_tolerance:
				arrays[h].append(e)
				if len(arrays[h]) > w:
					w = len(arrays[h])
			else:
				if w > -1:
					arrays[h].sort(key=lambda widget: (widget.rectangle().left, -widget.rectangle().width()))
				arrays.append([])
				h = h + 1
				arrays[h].append(e)
				y = e.rectangle().top
		arrays[h].sort(key=lambda widget: (widget.rectangle().left, -widget.rectangle().width()))
	else:
		return 0, 0, []
	return h + 1, w, arrays


def find_window_candidates(root_entry, visible_only=True, enabled_only=True, active_only=True):
	"""
	It returns a list of windows that match the given root entry
	
	:param root_entry: The root entry of the window you want to find
	:param visible_only: If True, only visible windows are returned, defaults to True (optional)
	:param enabled_only: If True, only enabled controls are returned, defaults to True (optional)
	:param active_only: If True, only the active window will be returned, defaults to True (optional)
	:return: A list of window candidates.
	"""
	title, control_type, _, _ = get_entry(root_entry)
	if root_entry == "*":
		window_candidates = desktop.windows(
			visible_only=visible_only, enabled_only=enabled_only, active_only=active_only)
	elif is_regex_entry(root_entry):
		window_candidates = desktop.windows(
			title_re=title, control_type=control_type,
			visible_only=visible_only, enabled_only=enabled_only, active_only=active_only)
	else:
		window_candidates = desktop.windows(
			title=title, control_type=control_type,
			visible_only=visible_only, enabled_only=enabled_only, active_only=active_only)
	if not window_candidates:
		if active_only:
			return find_window_candidates(root_entry, visible_only=True, enabled_only=False, active_only=False)
		else:
			print("Warning: No window '" + title + "' with control type '" + control_type + "' found! ")
			return None
	return window_candidates


def filter_window_candidates(window_candidates):
	"""
	It filters out windows that are not in the list of admitted windows, or that are in the list of ignored windows
	
	:param window_candidates: a list of Window objects
	:return: A list of window candidates that have been filtered.
	"""
	global core_settings
	if core_settings.window_filtering.mode == 'admit_windows':
		window_candidates = list(filter(
			lambda w: any(substr in w.element_info.name for substr in core_settings.window_filtering.admit_windows),
			window_candidates))
	else:
		window_candidates = list(filter(
			lambda w: not any(substr in w.element_info.name for substr in core_settings.window_filtering.ignore_windows),
			window_candidates))
	return window_candidates


def find_ocr_elements(ocr_text, window, entry_list):
	entry_list = entry_list[:-1]
	title, control_type, _, _ = get_entry(entry_list[-1])
	while entry_list[-1] == '*':
		entry_list = entry_list[:-1]
		title, control_type, _, _ = get_entry(entry_list[-1])
	descendants = window.descendants(title=title, control_type=control_type)
	candidates = list(filter(lambda e: match_entry_list(get_entry_list(get_wrapper_path(e)), entry_list), descendants))
	ocr_candidates = []
	if not candidates:
		candidates = [window]
	for wrapper in candidates:
		results = find_all_ocr(wrapper)
		for r in results:
			if fuzz.partial_ratio(ocr_text, r[1]) > 90:
				ocr_candidates.append(OCRWrapper(r))
	perfect_candidate = [ocr_candidate for ocr_candidate in ocr_candidates if ocr_candidate.result[1] == ocr_text]
	if len(perfect_candidate) == 1:
		return perfect_candidate
	return ocr_candidates


def find_elements(full_element_path=None, visible_only=True, enabled_only=True, active_only=True):
	"""
	It takes an entry list and returns the elements that matche the entry list

	:param full_element_path: full path of the element(s) to find
	:param visible_only: If True, only visible elements are returned, defaults to True (optional)
	:param enabled_only: If True, only enabled controls are returned, defaults to True (optional)
	:param active_only: If True, only active windows are considered, defaults to True (optional)
	:return: The elements found
	"""
	entry_list = get_entry_list(full_element_path)
	window_candidates = find_window_candidates(entry_list[0], visible_only=visible_only, enabled_only=enabled_only,
	                                           active_only=active_only)
	if window_candidates is None:
		return []
	window_candidates = filter_window_candidates(window_candidates)
	if len(entry_list) == 1 and len(window_candidates) == 1:
		return [window_candidates[0]]

	candidates = []
	title, control_type, _, _ = get_entry(entry_list[-1])
	for window in window_candidates:
		if is_regex_entry(entry_list[-1]):
			eis = findwindows.find_elements(title_re=title, control_type=control_type, backend="uia", parent=window, top_level_only=False)
			descendants = [UIAWrapper(ei) for ei in eis]
			candidates += filter(lambda e: match_entry_list(get_entry_list(get_wrapper_path(e)), entry_list), descendants)
		else:
			if control_type == "OCR_Text":
				candidates += find_ocr_elements(title, window, entry_list)
			else:
				descendants = window.descendants(title=title, control_type=control_type)  # , depth=max(1, len(entry_list)-2)
				candidates += filter(lambda e: match_entry_list(get_entry_list(get_wrapper_path(e)), entry_list), descendants)
	if not candidates:
		if active_only:
			return find_elements(full_element_path, visible_only=True, enabled_only=False, active_only=False)
		else:
			return []
	else:
		return candidates


# The following code should be called in recoder.py because it's useful only when recording
def read_config_file():
	"""
	It reads the configuration file and stores the settings in the global variable `core_settings`
	"""
	global core_settings
	config_file_content = '''[window_filtering]
    mode = ignore_windows
    admit_windows = []
    ignore_windows = []
    '''
	from pathlib import Path
	config_file = Path.home() / 'Pywinauto recorder' / 'config.ini'
	if config_file.exists and config_file.is_file():
		print("Reading configuration file: " + str(config_file))
		config = configparser.RawConfigParser()
		config.read(config_file)
		
		# config.read_string(ini_config)
		
		class Dict2Obj:
			def __init__(self, in_dict: dict):
				assert isinstance(in_dict, dict)
				for key, val in in_dict.items():
					if isinstance(val, (list, tuple)):
						setattr(self, key, [Dict2Obj(x) if isinstance(x, dict) else x for x in val])
					else:
						setattr(self, key, Dict2Obj(val) if isinstance(val, dict) else val)
		
		core_settings = Dict2Obj({s: dict(config.items(s)) for s in config.sections()})
		print("Window filtering mode: " + core_settings.window_filtering.mode)
		if core_settings.window_filtering.mode == "admit_windows":
			print("Admitted windows: " + core_settings.window_filtering.admit_windows)
			core_settings.window_filtering.admit_windows = ast.literal_eval(core_settings.window_filtering.admit_windows)
		else:
			print("Ignored windows: " + core_settings.window_filtering.ignore_windows)
			core_settings.window_filtering.ignore_windows = ast.literal_eval(core_settings.window_filtering.ignore_windows)
	else:
		print("Warning: '" + str(config_file) + "' not found! This file is created with a default configuration!")
		home_dir = Path.home() / 'Pywinauto recorder'
		home_dir.mkdir(parents=True, exist_ok=True)
		with open(config_file, 'w') as out_file:
			out_file.write(config_file_content)


core_settings = CoreSettings()
