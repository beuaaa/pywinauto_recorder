# -*- coding: utf-8 -*-

import core
import pywinauto
import win32api
import win32con
import time
from ctypes.wintypes import tagPOINT
from enum import Enum


class MoveMode(Enum):
    linear = 0
    y_first = 1
    x_first = 2


element_path_start = ''
unique_element_old = None
element_path_old = ''
w_rOLD = None
click_desktop = None


def get_element_path_start():
    global element_path_start
    return element_path_start


def in_region(element_path):
    global element_path_start
    element_path_start = element_path


def find(element_path):
    global element_path_start
    global click_desktop
    if not click_desktop:
        click_desktop = pywinauto.Desktop(backend='uia', allow_magic_lookup=False)

    if element_path_start:
        element_path2 = element_path_start + "->" + element_path
    else:
        element_path2 = element_path

    entry_list = core.get_entry_list(element_path2.decode('utf-8'))
    i = 0
    unique_element = None
    elements = None
    strategy = None
    while i < 99:
        try:
            unique_element, elements = core.find_element(click_desktop, entry_list, window_candidates=[])
        except Exception:
            time.sleep(0.1)
        i += 1

        _, _, y_x, _ = core.get_entry(entry_list[-1])
        if y_x is not None:
            nb_y, nb_x, candidates = core.get_sorted_region(elements)
            if core.is_int(y_x[0]):
                unique_element = candidates[int(y_x[0])][int(y_x[1])]
            else:
                ref_entry_list = core.get_entry_list(element_path_start.decode('utf-8')) + [y_x[0]]
                ref_unique_element, _ = core.find_element(click_desktop, ref_entry_list, window_candidates=[])
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
            _, control_type0, _, _ = core.get_entry(entry_list[0])
            _, control_type1, _, _ = core.get_entry(entry_list[-1])
            if control_type0 == 'Menu' or control_type1 == 'TreeItem':
                # Wait if element is not clickable (greyed, not still visible) :
                # So far, I didn't find better than wait_cpu_usage_lower but must be enhanced
                app = pywinauto.Application(backend='uia', allow_magic_lookup=False)
                app.connect(process=unique_element.element_info.element.CurrentProcessId)
                app.wait_cpu_usage_lower()
            if unique_element.is_enabled():
                break
    if not unique_element:
        raise Exception("Unique element not found! ", element_path)

    return unique_element


def move(element_path, duration=0.5, mode=MoveMode.linear):
    global unique_element_old
    global element_path_old
    global w_rOLD

    entry_list = core.get_entry_list(element_path.decode('utf-8'))
    if element_path == element_path_old:
        w_r = w_rOLD
        unique_element = unique_element_old
    else:
        unique_element = find(element_path)
        w_r = unique_element.rectangle()

    x, y = win32api.GetCursorPos()
    _, control_type, _, _ = core.get_entry(entry_list[0])
    if control_type == 'Menu':
        entry_list_old = core.get_entry_list(element_path_old.decode('utf-8'))
        _, control_type_old, _, _ = core.get_entry(entry_list_old[0])
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

    unique_element_old = unique_element
    element_path_old = element_path
    w_rOLD = w_r
    return unique_element


def mouse_wheel(steps):
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, win32con.WHEEL_DELTA * steps, 0)


def click(element_path, duration=0.5, mode=MoveMode.linear, button='left'):
    unique_element = move(element_path, duration=duration, mode=mode)

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


def left_click(element_path, duration=0.5, mode=MoveMode.linear):
    return click(element_path, duration=duration, mode=mode, button='left')


def right_click(element_path, duration=0.5, mode=MoveMode.linear):
    return click(element_path, duration=duration, mode=mode, button='right')


def double_left_click(element_path, duration=0.5, mode=MoveMode.linear):
    return click(element_path, duration=duration, mode=mode, button='double_left')


def triple_left_click(element_path, duration=0.5, mode=MoveMode.linear):
    return click(element_path, duration=duration, mode=mode, button='triple_left')


def drag_and_drop(element_path, duration=0.5, mode=MoveMode.linear):
    words = element_path.split("%(")
    element_path1 = words[0] + "%(" + words[1]
    unique_element = move(element_path1, duration=duration, mode=mode)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    words = element_path.split("%(")
    last_word = words[-1]
    words = words[:-1]
    words = words[:-1]
    element_path2 = ''
    for w in words:
        element_path2 = element_path2 + w + "%("
    element_path2 = element_path2 + last_word
    move(element_path2, duration=duration, mode=mode)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    return unique_element


def send_keys(str_keys):
    pywinauto.keyboard.send_keys(str_keys, with_spaces=True)
