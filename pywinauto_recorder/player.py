# -*- coding: utf-8 -*-

# print('__file__={0:<35} | __name__={1:<20} | __package__={2:<20}'.format(__file__, __name__, str(__package__)))
from six import string_types
from .core import *
import pywinauto
import win32api
import win32con
import win32gui
import time
from enum import Enum
import copy


class MoveMode(Enum):
    linear = 0
    y_first = 1
    x_first = 2


unique_element_old = None
element_path_old = ''
w_rOLD = None


def wait_is_ready_try1(wrapper, timeout=120):
    """
    Wait until element is ready (wait while greyed, not enabled, not visible, not ready, ...) :
    So far, I didn't find better than wait_cpu_usage_lower when greyed but must be enhanced
    """
    t0 = time.time()
    while not wrapper.is_enabled() or not wrapper.is_visible():
        try:
            h_wait_cursor = win32gui.LoadCursor(0, win32con.IDC_WAIT)
            _, h_cursor, _ = win32gui.GetCursorInfo()
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
            pass
        if (time.time() - t0) > timeout:
            raise Exception("Time out! ", wrapper)


class Region(object):
    wait_element_is_ready = wait_is_ready_try1
    common_path = ''
    list_path = []
    regex_title = False
    click_desktop = None
    current = None

    def __init__(self, relative_path=None, regex_title=False):
        self.relative_path = relative_path
        self.regex_title = regex_title

    def __enter__(self):
        Region.current = self
        if not Region.list_path:
            Region.regex_title = self.regex_title
        else:
            self.regex_title = Region.regex_title
        if self.relative_path:
            Region.list_path.append(self.relative_path)
        Region.common_path = path_separator.join(self.list_path)
        return self

    def __exit__(self, type, value, traceback):
        Region.list_path = Region.list_path[0:-1]
        Region.common_path = path_separator.join(self.list_path)


def find(element_path, timeout=60*5):
    if not Region.click_desktop:
        Region.click_desktop = pywinauto.Desktop(backend='uia', allow_magic_lookup=False)
    if Region.common_path:
        if element_path:
            element_path2 = Region.common_path + path_separator + element_path
        else:
            element_path2 = Region.common_path
    else:
        element_path2 = element_path
    entry_list = get_entry_list(element_path2)
    unique_element = None
    elements = None
    strategy = None
    t0 = time.time()
    while (time.time() - t0) < timeout:
        while unique_element is None and not elements:
            try:
                unique_element, elements = find_element(
                    Region.click_desktop, entry_list, window_candidates=[], regex_title=Region.current.regex_title)
                if unique_element is None and not elements:
                    time.sleep(2.0)
            except Exception:
                pass
            if (time.time() - t0) > timeout:
                raise Exception("Time out! ", element_path2)

        _, _, y_x, _ = get_entry(entry_list[-1])
        if y_x is not None:
            nb_y, nb_x, candidates = get_sorted_region(elements)
            if is_int(y_x[0]):
                unique_element = candidates[int(y_x[0])][int(y_x[1])]
            else:
                ref_entry_list = get_entry_list(Region.common_path) + get_entry_list(y_x[0])
                ref_unique_element, _ = find_element(
                    Region.click_desktop, ref_entry_list, window_candidates=[], regex_title=Region.current.regex_title)
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
            break
        time.sleep(0.1)
    if not unique_element:
        raise Exception("Unique element not found! ", element_path2)
    return unique_element


def move(element_path, duration=0.5, mode=MoveMode.linear):
    global unique_element_old
    global element_path_old
    global w_rOLD

    x, y = win32api.GetCursorPos()
    if isinstance(element_path, string_types):
        element_path2 = element_path
        if Region.common_path:
            if Region.common_path != element_path[0:len(Region.common_path)]:
                element_path2 = Region.common_path + path_separator + element_path
        entry_list = get_entry_list(element_path2)
        if element_path2 == element_path_old:
            w_r = w_rOLD
            unique_element = unique_element_old
        else:
            unique_element = find(element_path)
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
            xd, yd = xd + round(dx/100.0*w_r.width(), 0), round(yd + dy/100.0*w_r.height(), 0)
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
                nx = int(x * 65535 / win32api.GetSystemMetrics(0))
                ny = int(y * 65535 / win32api.GetSystemMetrics(1))
                win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)
            step_x = 0
        if mode == MoveMode.y_first:
            for i in range(int(samples)):
                y = y+step_y
                time.sleep(0.01)
                nx = int(x * 65535 / win32api.GetSystemMetrics(0))
                ny = int(y * 65535 / win32api.GetSystemMetrics(1))
                win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)
            step_y = 0
        for i in range(int(samples)):
            x, y = x+step_x, y+step_y
            time.sleep(0.01)
            nx = int(x * 65535 / win32api.GetSystemMetrics(0))
            ny = int(y * 65535 / win32api.GetSystemMetrics(1))
            win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)
    nx = int(xd * 65535 / win32api.GetSystemMetrics(0))
    ny = int(yd * 65535 / win32api.GetSystemMetrics(1))
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE, nx, ny)
    if unique_element is None:
        return None
    unique_element_old = unique_element
    element_path_old = element_path2
    w_rOLD = w_r
    return unique_element


def click(element_path, duration=0.5, mode=MoveMode.linear, button='left'):
    unique_element = move(element_path, duration=duration, mode=mode)
    if isinstance(element_path, string_types):
        wait_is_ready_try1(unique_element, timeout=60*5)
    else:
        unique_element = None
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


def drag_and_drop(element_path1, element_path2, duration=0.5, mode=MoveMode.linear):
    unique_element = move(element_path1, duration=duration, mode=mode)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    move(element_path2, duration=duration, mode=mode)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    return unique_element


def menu_click(element_path, menu_path, duration=0.5, mode=MoveMode.linear, menu_type='QT'):
    menu_entry_list = menu_path.split(path_separator)
    if menu_type == 'QT':
        menu_entry_list = [''] + menu_entry_list
    else:
        menu_entry_list = ['Application'] + menu_entry_list
    left_click(
        element_path +
        menu_entry_list[0] + type_separator + 'MenuBar' + path_separator +
        menu_entry_list[1] + type_separator + 'MenuItem', duration=duration, mode=mode)
    w = None
    if menu_type == 'QT':
        common_path_old = Region.common_path
        Region.common_path = ''
        for entry in menu_entry_list[2:]:
            w = left_click(
                type_separator + 'Menu' + path_separator +
                entry + type_separator + 'MenuItem', duration=duration, mode=mode)
        Region.common_path = common_path_old
    else:
        for i, entry in enumerate(menu_entry_list[2:]):
            w = left_click(
                element_path +
                menu_entry_list[i - 2] + type_separator + 'Menu' + path_separator +
                entry + type_separator + 'MenuItem', duration=duration, mode=mode)
    return w


Window = Region


def mouse_wheel(steps):
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, win32con.WHEEL_DELTA * steps, 0)


def send_keys(
    str_keys,
    pause=0.1,
    with_spaces=True,
    with_tabs=True,
    with_newlines=True,
    turn_off_numlock=True,
    vk_packet=True):
    """Parse the keys and type them"""
    pywinauto.keyboard.send_keys(
        str_keys,
        pause=pause,
        with_spaces=with_spaces,
        with_tabs=with_tabs,
        with_newlines=with_newlines,
        turn_off_numlock=turn_off_numlock,
        vk_packet=vk_packet
    )
