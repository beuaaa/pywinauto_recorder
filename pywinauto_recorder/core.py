# -*- coding: utf-8 -*-

import re
from enum import Enum


class Strategy(Enum):
    failed = 0
    unique_path = 1
    array_1D = 2
    array_2D = 3
    ancestor_unique_path = 4


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
                print ("Warning: No window found!")
                return Strategy.failed, None, []

    if len(entry_list) == 1 and len(window_candidates) == 1:
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
        # Strategy 1: unique path
        # Strategy 2: 2D array of elements
        # Strategy 3: 1D array of elements beginning with an element having a unique path
        # Strategy 4: we find a unique path in the ancestors
        #_, unique_candidate, candidates = find_element(desktop, entry_list[0:-1], window_candidates=window_candidates)
        return None, candidates
