# -*- coding: utf-8 -*-

import core
import pywinauto
import win32api
import win32con
import time
from enum import Enum


class MoveMode(Enum):
    linear = 0
    y_first = 1
    x_first = 2


unique_element_old = None
element_path_old = ''
w_rOLD = None


class Region(object):
    common_path = ''

    def __init__(self, relative_path):
        self.click_desktop = None
        self.relative_path = relative_path

    def __enter__(self):
        if Region.common_path:
            Region.common_path = Region.common_path + core.path_separator + self.relative_path
        else:
            Region.common_path = self.relative_path
        return self

    def __exit__(self, type, value, traceback):
        i = Region.common_path.find(self.relative_path)
        if i != -1:
            if i == 0:
                Region.common_path = ''
            else:
                Region.common_path = Region.common_path[0:i-len(core.path_separator)]

    def find(self, element_path):
        if not self.click_desktop:
            self.click_desktop = pywinauto.Desktop(backend='uia', allow_magic_lookup=False)
        if Region.common_path:
            element_path2 = Region.common_path + core.path_separator + element_path
        else:
            element_path2 = element_path
        entry_list = core.get_entry_list(element_path2)
        i = 0
        unique_element = None
        elements = None
        strategy = None
        while i < 99:
            try:
                unique_element, elements = core.find_element(self.click_desktop, entry_list, window_candidates=[])
            except Exception:
                pass
            i += 1
            _, _, y_x, _ = core.get_entry(entry_list[-1])
            if y_x is not None:
                nb_y, nb_x, candidates = core.get_sorted_region(elements)
                if core.is_int(y_x[0]):
                    unique_element = candidates[int(y_x[0])][int(y_x[1])]
                else:
                    ref_entry_list = core.get_entry_list(Region.common_path) + [y_x[0]]
                    ref_unique_element, _ = core.find_element(self.click_desktop, ref_entry_list, window_candidates=[])
                    ref_r = ref_unique_element.rectangle()
                    r_y = 0
                    while r_y < nb_y:
                        y_candidate = candidates[r_y][0].rectangle().mid_point()[1]
                        if ref_r.top < y_candidate < ref_r.bottom:
                            unique_element = candidates[r_y][y_x[1]]
                            break
                        r_y = r_y + 1
                    else:
                        unique_element = None
            if unique_element is not None:
                # Wait if element is not clickable (greyed, not still visible) :
                # So far, I didn't find better than wait_cpu_usage_lower but must be enhanced
                _, control_type0, _, _ = core.get_entry(entry_list[-1])
                if control_type0 in ('Menu', 'MenuItem', 'TreeItem'):
                    app = pywinauto.Application(backend='uia', allow_magic_lookup=False)
                    app.connect(process=unique_element.element_info.element.CurrentProcessId)
                    app.wait_cpu_usage_lower()
                if unique_element.is_enabled():
                    break
            time.sleep(0.1)
        if not unique_element:
            raise Exception("Unique element not found! ", element_path)
        return unique_element

    def move(self, element_path, duration=0.5, mode=MoveMode.linear):
        global unique_element_old
        global element_path_old
        global w_rOLD

        x, y = win32api.GetCursorPos()
        if isinstance(element_path, basestring):
            element_path2 = element_path
            if Region.common_path:
                if Region.common_path != element_path[0:len(Region.common_path)]:
                    element_path2 = Region.common_path + core.path_separator + element_path
            entry_list = core.get_entry_list(element_path2)
            if element_path2 == element_path_old:
                w_r = w_rOLD
                unique_element = unique_element_old
            else:
                unique_element = self.find(element_path)
                w_r = unique_element.rectangle()
            control_type = None
            for entry in entry_list:
                _, control_type, _, _ = core.get_entry(entry)
                if control_type == 'Menu':
                    break
            if control_type == 'Menu':
                entry_list_old = core.get_entry_list(element_path_old)
                control_type_old = None
                for entry in entry_list_old:
                    _, control_type_old, _, _ = core.get_entry(entry)
                    if control_type_old == 'Menu':
                        break
                if control_type_old == 'Menu':
                    mode = MoveMode.x_first
                else:
                    mode = MoveMode.y_first
                xd, yd = w_r.mid_point()
            else:
                _, _, _, dx_dy = core.get_entry(entry_list[-1])
                if dx_dy:
                    dx, dy = dx_dy[0], dx_dy[1]
                else:
                    dx, dy = 0, 0
                xd, yd = w_r.mid_point()
                xd, yd = xd + dx, yd + dy
        else:
            (xd, yd) = element_path
            unique_element = None
        if (x, y) != (xd, yd):
            dt = 0.01
            samples = duration/dt
            step_x = (xd-x)/samples
            step_y = (yd-y)/samples
            if mode == MoveMode.x_first:
                for i in range(int(samples)):
                    x = x+step_x
                    time.sleep(0.01)
                    nx = int(x) * 65535 / win32api.GetSystemMetrics(0)
                    ny = int(y) * 65535 / win32api.GetSystemMetrics(1)
                    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)
                step_x = 0
            if mode == MoveMode.y_first:
                for i in range(int(samples)):
                    y = y+step_y
                    time.sleep(0.01)
                    nx = int(x) * 65535 / win32api.GetSystemMetrics(0)
                    ny = int(y) * 65535 / win32api.GetSystemMetrics(1)
                    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)
                step_y = 0
            for i in range(int(samples)):
                x, y = x+step_x, y+step_y
                time.sleep(0.01)
                nx = int(x) * 65535 / win32api.GetSystemMetrics(0)
                ny = int(y) * 65535 / win32api.GetSystemMetrics(1)
                win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)
        nx = int(xd) * 65535 / win32api.GetSystemMetrics(0)
        ny = int(yd) * 65535 / win32api.GetSystemMetrics(1)
        win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)
        if unique_element is None:
            return None
        unique_element_old = unique_element
        element_path_old = element_path2
        w_rOLD = w_r
        return unique_element

    def click(self, element_path, duration=0.5, mode=MoveMode.linear, button='left'):
        unique_element = self.move(element_path, duration=duration, mode=mode)
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
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0)
            time.sleep(.01)
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0)
            time.sleep(.01)
        return unique_element

    def left_click(self, element_path, duration=0.5, mode=MoveMode.linear):
        return self.click(element_path, duration=duration, mode=mode, button='left')

    def right_click(self, element_path, duration=0.5, mode=MoveMode.linear):
        return self.click(element_path, duration=duration, mode=mode, button='right')

    def double_left_click(self, element_path, duration=0.5, mode=MoveMode.linear):
        return self.click(element_path, duration=duration, mode=mode, button='double_left')

    def triple_left_click(self, element_path, duration=0.5, mode=MoveMode.linear):
        return self.click(element_path, duration=duration, mode=mode, button='triple_left')

    def drag_and_drop(self, element_path, duration=0.5, mode=MoveMode.linear):
        words = element_path.split("%(")
        element_path1 = words[0] + "%(" + words[1]
        unique_element = self.move(element_path1, duration=duration, mode=mode)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        words = element_path.split("%(")
        last_word = words[-1]
        words = words[:-1]
        words = words[:-1]
        element_path2 = ''
        for w in words:
            element_path2 = element_path2 + w + "%("
        element_path2 = element_path2 + last_word
        self.move(element_path2, duration=duration, mode=mode)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
        return unique_element

    def menu_click(self, element_path, menu_path, duration=0.5, mode=MoveMode.linear, menu_type='QT'):
        menu_entry_list = menu_path.split(core.path_separator)
        self.left_click(
            element_path +
            menu_entry_list[0] + core.type_separator + 'MenuBar' + core.path_separator +
            menu_entry_list[1] + core.type_separator + 'MenuItem', duration=duration, mode=mode)

        if menu_type is 'QT':
            common_path_old = Region.common_path
            Region.common_path = ''
            for entry in menu_entry_list[2:]:
                w = self.left_click(
                    core.type_separator + 'Menu' + core.path_separator +
                    entry + core.type_separator + 'MenuItem', duration=duration, mode=mode)
            Region.common_path = common_path_old
        else:
            for i, entry in enumerate(menu_entry_list[2:]):
                w = self.left_click(
                    element_path +
                    menu_entry_list[i - 3] + core.type_separator + 'Menu' + core.path_separator +
                    entry + core.type_separator + 'MenuItem', duration=duration, mode=mode)
        return w


def mouse_wheel(steps):
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, win32con.WHEEL_DELTA * steps, 0)


def send_keys(str_keys):
    pywinauto.keyboard.send_keys(str_keys, with_spaces=True)
