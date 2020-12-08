# -*- coding: utf-8 -*-

import re
import traceback
from enum import Enum

path_separator = "->"
type_separator = "||"


class Strategy(Enum):
    unique_path = 1
    array_1D = 2
    array_2D = 3


def is_int(s):
    try:
        int(s)
        return True
    except ValueError:
        return False


def get_wrapper_path(wrapper):
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
    #path = path.decode('utf-8')
    i = path.rfind("#[")
    if i != -1:
        i = path.rfind(path_separator, 0, i)
    else:
        i = path.rfind(path_separator)
    if i == -1:
        return [path]
    last_entry = path[i+len(path_separator)::]
    start_entry = path[:i]
    return start_entry.split(path_separator) + [last_entry]



def get_entry(entry):
    i = entry.find(type_separator)
    if i == -1:
        return entry, None, None, None
    while i < len(entry) and entry[i] == type_separator[0]:
        i = i + 1
    i = i - len(type_separator)
    str_name = entry[0:i]

    entry2 = entry[i+len(type_separator):]
    i = entry2.rfind("""%(""")

    if i != -1:
        str_dx_dy = entry2[i+2:]
        entry3 = entry2[0:i]
    else:
        str_dx_dy = None
        entry3 = entry2

    i = entry3.rfind("""#[""")
    if i != -1:
        str_array = entry3[i+2:]
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
        dx = float(words[0])
        dy = float(words[1][:-1])
        dx_dy = (dx, dy)
    else:
        dx_dy = None
    return str_name, str_type, y_x, dx_dy


# TODO windowName||windowType->*->||->name||->||type
def same_entry_list(element, entry_list, regex_title=False):
    try:
        i = len(entry_list) - 1
        top_level_parent = element.top_level_parent()
        current_element = element
        while i >= 0:
            if entry_list[i] == '*':    # TODO: make possible to use more than one asterisk in second position
                current_element = top_level_parent
                i = 0
            current_element_name = current_element.element_info.name
            current_element_type = current_element.element_info.control_type
            entry_name, entry_type, _, _ = get_entry(entry_list[i])
            if i == 0 and current_element == top_level_parent:
                if regex_title:
                    return (re.match(entry_list[0], entry_name) or not entry_name)\
                           and (current_element_type == entry_type or not entry_type)
                else:
                    return (current_element_name == entry_name or not entry_name)\
                            and (current_element_type == entry_type or not entry_type)
            elif (current_element_name == entry_name or not entry_name)\
                    and (current_element_type == entry_type or not entry_type):
                i -= 1
                current_element = current_element.parent()
            else:
                return False
        return False
    except Exception:
        return False


def is_filter_criteria_ok(child, min_height=8, max_height=200, min_width=8, max_width=800):
    if child.is_visible():
        h = child.rectangle().height()
        if (min_height <= h) and (h <= max_height):
            w = child.rectangle().width()
            if (min_width <= w) and (w <= max_width):
                if child.rectangle().top > 0:
                    return True
    return False


def get_sorted_region(elements, min_height=8, max_height=9999, min_width=8, max_width=9999, line_tolerance=20):
    filtered_elements = []
    for e in elements:
        if is_filter_criteria_ok(e, min_height, max_height, min_width, max_width):
            filtered_elements.append(e)

    filtered_elements.sort(key=lambda widget: (widget.rectangle().top, widget.rectangle().left))
    arrays = [[]]

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


def find_element(desktop, entry_list, window_candidates=[], visible_only=True, enabled_only=True, active_only=True, regex_title=False):
    if not window_candidates:
        title, control_type, _, _ = get_entry(entry_list[0])
        if regex_title:
            window_candidates = desktop.windows(
                title_re=title, control_type=control_type, visible_only=visible_only,
                enabled_only=enabled_only, active_only=active_only)
        else:
            window_candidates = desktop.windows(
                title=title, control_type=control_type, visible_only=visible_only,
                enabled_only=enabled_only, active_only=active_only)
        if not window_candidates:
            if active_only:
                return find_element(
                    desktop, entry_list, window_candidates=[], visible_only=True,
                    enabled_only=False, active_only=False, regex_title=regex_title)
            else:
                print("Warning: No window '" + title + "' with control type '" + control_type + "' found! ")
                return None, []

    if len(entry_list) == 1 and len(window_candidates) == 1:
        return window_candidates[0], []

    candidates = []
    for window in window_candidates:
        title, control_type, _, _ = get_entry(entry_list[-1])
        descendants = window.descendants(title=title, control_type=control_type) # , depth=max(1, len(entry_list)-2)
        for descendant in descendants:
            if same_entry_list(descendant, entry_list, regex_title=regex_title):
                candidates.append(descendant)
            else:
                continue

    if not candidates:
        if active_only:
            return find_element(
                desktop, entry_list, window_candidates=[], visible_only=True,
                enabled_only=False, active_only=False, regex_title=regex_title)
        else:
            return None, []
    elif len(candidates) == 1:
        return candidates[0], []
    else:
        return None, candidates
