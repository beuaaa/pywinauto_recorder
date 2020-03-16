# -*- coding: utf-8 -*-

import re


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
                print ("Warning: No window found!")
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
        # return (Strategy 3, unique_candidate, candidates)
        return candidates[0], candidates
