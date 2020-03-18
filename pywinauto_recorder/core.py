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


def isFilterCriteriaOk(child, minHeight=8, maxHeight=200, minWidth=8, maxWidth=800):
    if child.is_visible():
        h = child.rectangle().height()
        if (minHeight <= h) and (h <= maxHeight):
            w = child.rectangle().width()
            if (minWidth <= w) and (w <= maxWidth):
                if child.rectangle().top > 0:
                    return True
    return False


def getSortedRegion(children, minHeight=8, maxHeight=9999, minWidth=8, maxWidth=9999, lineTolerance=20):
    widgetList = []
    for child in children:
        if isFilterCriteriaOk(child, minHeight, maxHeight, minWidth, maxWidth):
            widgetList.append(child)

    widgetList.sort(key=lambda widget: (widget.rectangle().top, widget.rectangle().left))
    widgetLists = []
    widgetLists.append([])

    h = 0
    w = -1
    if len(widgetList) > 0:
        y = widgetList[0].rectangle().top
        for child in widgetList:
            if (child.rectangle().top - y) < lineTolerance:
                widgetLists[h].append(child)
                if len(widgetLists[h]) > w:
                    w = len(widgetLists[h])
            else:
                if w > -1:
                    widgetLists[h].sort(key=lambda widget: (widget.rectangle().left, -widget.rectangle().width()))
                widgetLists.append([])
                h = h + 1
                widgetLists[h].append(child)
                y = child.rectangle().top
        widgetLists[h].sort(key=lambda widget: (widget.rectangle().left, -widget.rectangle().width()))
    else:
        return 0, 0, []
    return h + 1, w, widgetLists


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

    if len(entry_list) == 1:
        if len(window_candidates) == 1:
            return Strategy.unique_path, window_candidates[0], []

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
            return Strategy.failed, None, []
    elif len(candidates) == 1:
        return Strategy.unique_path, candidates[0], []
    else:
        # We have several elements so we have to use the good strategy to select the good one.
        # Strategy 1: 1D array of elements beginning with an element having a unique path
        # Strategy 2: 2D array of elements
        # Strategy 3: we find a unique path in the ancestors
        #_, unique_candidate, candidates = find_element(desktop, entry_list[0:-1], window_candidates=window_candidates)
        return Strategy.array_2D, None, candidates
