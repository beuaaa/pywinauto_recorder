# -*- coding: utf-8 -*-

import re
import traceback
from enum import Enum
import configparser
import ast
from typing import Optional, Union, NewType
from pywinauto import Desktop as PywinautoDesktop
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto import findwindows

__all__ = ['path_separator', 'type_separator', 'Strategy', 'is_int',
           'get_wrapper_path', 'get_entry_list', 'get_entry', 'match_entry_list', 'get_sorted_region',
           'find_element', 'read_config_file']

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
	except Exception as e:
		traceback.print_exc()
		print(e.message)
		return ''


def get_entry_list(path):
	"""
	It splits the path into a list of entries
	
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


def get_entry(entry):
	"""
	It takes a string of the form `"name"#["y",x]%(dx,dy)` and returns a tuple of the form `(name, type, [y,x], (dx,dy))`
	
	:param entry: the string that is the entry in the listbox
	:return: The name of the entry, the type of the entry, the size of the entry, and the dx and dy of the entry.
	"""
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
		str_array = entry3[i + 2:]
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
		x = int(words[1][:-1])
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
		return (re.match(template_name, entry_name) or not entry_name) \
		       and (template_type == entry_type or not entry_type)
	else:
		return (template_name == entry_name or not entry_name) \
		       and (template_type == entry_type or not entry_type)


def match_entry_sequence(i_e, entry_list, sequence_list):
	"""
	If the entry at index `i_e` in `entry_list` matches the first entry in `sequence_list`, and the entry at index `i_e+1`
	in `entry_list` matches the second entry in `sequence_list`, and so on, then return `True`. Otherwise, return `False`
	
	:param i_e: the index of the entry in the entry list that we're currently looking at
	:param entry_list: a list of entries, each entry is a list of words
	:param sequence_list: a list of strings, each of which is a sequence of characters
	:return: A boolean value.
	"""
	for i_s in range(len(sequence_list)):
		if not match_entry(entry_list[i_e + i_s], sequence_list[i_s]):
			return False
	return True


def find_entry_sequence_after(i_e, entry_list, sequence_list):
	"""
	> Find the index of the first entry in the entry list that matches the sequence list, starting at index i_e
	
	:param i_e: the index of the entry in the entry list
	:param entry_list: a list of entries
	:param sequence_list: a list of sequences, each sequence is a list of entries
	:return: The index of the first entry in the entry list that matches the sequence list.
	"""
	while i_e < len(entry_list):
		if match_entry_sequence(i_e, entry_list, sequence_list):
			return i_e
		else:
			i_e += 1
	return -1


def match_entry_list(entry_list, template_list):
	"""
	It takes an entry list and a template list, and returns True if the entry list matches the template list
	
	:param entry_list: entry list of one of the elements found in descendants
	:param template_list: list of the entries of the element pattern (already completed with UIPath)
	:return: True if the entry list matches the template list.
	"""
	
	try:
		# print("match_entry_list: entry_list: " + str(entry_list) + " template_list: " + str(template_list))
		# using list comprehension + zip() + slicing + enumerate()
		# Split list into lists by particular value
		size = len(template_list)
		idx_list = [idx + 1 for idx, val in enumerate(template_list) if val == '*']
		if idx_list:
			res = [template_list[i: j] for i, j in zip([0] + idx_list, idx_list + ([size] if idx_list[-1] != size else []))]
			# Debug trace: print("The index list after splitting by a value : " + str(idx_list))
			# Debug trace: print("The list after splitting by a value : " + str(res))
			for sequence_list in res:
				if sequence_list[-1] == '*':
					del sequence_list[-1]
		else:
			res = [template_list]
		# print("The template list after splitting by ->*-> : " + str(res))
		i_e = 0
		if not match_entry_sequence(i_e, entry_list, res[0]):
			return False
		# print("The first sequence is matched: " + str(res[0]))
		i_e += len(res[0])
		res.pop(0)
		for sequence_t in res:
			i_e = find_entry_sequence_after(i_e, entry_list, sequence_t)
			if i_e == -1:
				# print("The sequence is not matched: " + str(sequence_t))
				return False
			# print("The sequence is matched: " + str(sequence_t))
			i_e = i_e + len(sequence_t)
		if i_e != len(entry_list):
			return False
		
		return True
	except Exception as e:
		print("match_entry_list exception: " + str(e))
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


def find_element(entry_list=None, visible_only=True, enabled_only=True, active_only=True):
	"""
	It takes an entry list and returns the unique element that matches the entry list

	:param entry_list: full path list
	:param visible_only: If True, only visible elements are returned, defaults to True (optional)
	:param enabled_only: If True, only enabled controls are returned, defaults to True (optional)
	:param active_only: If True, only active windows are considered, defaults to True (optional)
	:return: The element and the list of candidates
	"""
	window_candidates = find_window_candidates(entry_list[0],
	                                           visible_only=visible_only, enabled_only=enabled_only,
	                                           active_only=active_only)
	if window_candidates is None:
		return None, []
	
	window_candidates = filter_window_candidates(window_candidates)
	
	if len(entry_list) == 1 and len(window_candidates) == 1:
		return window_candidates[0], []
	
	candidates = []
	for window in window_candidates:
		title, control_type, _, _ = get_entry(entry_list[-1])
		if is_regex_entry(entry_list[-1]):
			eis = findwindows.find_elements(title_re=title, control_type=control_type,
			                                backend="uia", parent=window, top_level_only=False)
			descendants = [UIAWrapper(ei) for ei in eis]
			candidates += filter(lambda e: match_entry_list(get_entry_list(get_wrapper_path(e)),
			                                                entry_list), descendants)
		else:
			descendants = window.descendants(title=title,
			                                 control_type=control_type)  # , depth=max(1, len(entry_list)-2)
			candidates += filter(lambda e: match_entry_list(get_entry_list(get_wrapper_path(e)),
			                                                entry_list), descendants)
	
	if not candidates:
		if active_only:
			return find_element(entry_list, visible_only=True, enabled_only=False, active_only=False)
		else:
			return None, []
	elif len(candidates) == 1:
		return candidates[0], []
	else:
		return None, candidates


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
		print("Warning: '" + str(config_file) + "' not found! Default C file is created!")
		home_dir = Path.home() / 'Pywinauto recorder'
		home_dir.mkdir(parents=True, exist_ok=True)
		with open(config_file, 'w') as out_file:
			out_file.write(config_file_content)


core_settings = CoreSettings()
