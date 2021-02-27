# -*- coding: utf-8 -*-

# print('__file__={0:<35} | __name__={1:<20} | __package__={2:<20}'.format(__file__, __name__, str(__package__)))
from six import string_types
from .core import *
import pywinauto
from win32api import GetCursorPos as win32api_GetCursorPos
from win32api import GetSystemMetrics as win32api_GetSystemMetrics
from win32api import mouse_event as win32api_mouse_event
from win32gui import LoadCursor as win32gui_LoadCursor
from win32gui import GetCursorInfo as win32gui_GetCursorInfo
from win32con import IDC_WAIT, MOUSEEVENTF_MOVE, MOUSEEVENTF_ABSOLUTE, MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP, \
    MOUSEEVENTF_MIDDLEDOWN, MOUSEEVENTF_MIDDLEUP, MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP, MOUSEEVENTF_WHEEL, WHEEL_DELTA
import time
from enum import Enum
from typing import Any, Optional, Union

UI_Element = Union[str, pywinauto.controls.uiawrapper.UIAWrapper]

# TODO special_char_array in core for recorder.py and player.py (check when to call escape & unescape)
def unescape_special_char(string):
    for r in (("\\\\", "\\"), ("\\t", "\t"), ("\\n", "\n"), ("\\r", "\r"), ("\\v", "\v"), ("\\f", "\f"), ('\\"', '"')):
    #for r in (("\\", "\\\\"), ("\t", "\\t"), ("\n", "\\n"), ("\r", "\\r"), ("\v", "\\v"), ("\f", "\\f"), ('"', '\\"')):
        string = string.replace(*r)
    return string


class PlayerSettings:
    mouse_move_duration = 0.5


class MoveMode(Enum):
    linear = 0
    y_first = 1
    x_first = 2


_dictionary = {}
unique_element_old = None
element_path_old = ''
w_rOLD = None


def load_dictionary(filename):
    
    with open(filename) as fp:
        for line in fp:
            words = line.split("\t")
            i = 0
            while words[i] == '':
                i += 1
            variable = words[-1].translate(str.maketrans('', '', '\n\t\r'))
            value = words[i]
            _dictionary[variable] = value


def shortcut(variable):
    return _dictionary[variable]


def wait_is_ready_try1(wrapper, timeout=120):
    """
    Waits until element is ready (wait while greyed, not enabled, not visible, not ready, ...) :
    So far, I didn't find better than wait_cpu_usage_lower when greyed but must be enhanced
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
            pass
        if (time.time() - t0) > timeout:
            msg = "Element " + get_wrapper_path(wrapper) + "  was not found after " + str(timeout) + " s of searching."
            raise TimeoutError("Time out! ", msg)


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
        if self.relative_path:
            Region.list_path = Region.list_path[0:-1]
        Region.common_path = path_separator.join(self.list_path)


Window = Region


def find(element_path=None, timeout=120):
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
                if Region.current:
                    regex_title = Region.current.regex_title
                else:
                    regex_title = False
                unique_element, elements = find_element(
                    Region.click_desktop, entry_list, window_candidates=[], regex_title=regex_title)
                if unique_element is None and not elements:
                    time.sleep(2.0)
            except Exception:
                pass
            if (time.time() - t0) > timeout:
                msg = "Element " + element_path2 + "  was not found after " + str(timeout) + " s of searching."
                raise TimeoutError("Time out! ", msg)

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


def move(
        element_path: UI_Element,
        duration: Optional[float] = 0.5,
        mode: Enum = MoveMode.linear,
        timeout: float = 120) -> UI_Element:
    """
    Moves on element
    
    :param element_path: element path
    :param duration: duration in seconds of the mouse move
    :param mode: move mouse mode
    :param timeout: period of time in seconds that will be allowed to find the element
    :return: Pywinauto wrapper of clicked element
    """
    
    global unique_element_old
    global element_path_old
    global w_rOLD
    
    x, y = win32api_GetCursorPos()
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
            unique_element = find(element_path, timeout=timeout)
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
            xd, yd = xd + round(dx/100.0*(w_r.width()/2-1), 0), round(yd + dy/100.0*(w_r.height()/2-1), 0)
    elif issubclass(type(element_path), pywinauto.base_wrapper.BaseWrapper):
        unique_element = element_path
        element_path2 = get_wrapper_path(unique_element)
        w_r = unique_element.rectangle()
        xd, yd = w_r.mid_point()
    else:
        (xd, yd) = element_path
        unique_element = None
    x_max = win32api_GetSystemMetrics(0) - 1
    y_max = win32api_GetSystemMetrics(1) - 1
    if (x, y) != (xd, yd) and duration > 0:
        dt = 0.01
        samples = duration/dt
        step_x = (xd-x)/samples
        step_y = (yd-y)/samples
        if mode == MoveMode.x_first:
            for i in range(int(samples)):
                x = x+step_x
                time.sleep(0.01)
                nx = int(x * 65535 / x_max)
                ny = int(y * 65535 / y_max)
                win32api_mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, nx, ny)
            step_x = 0
        if mode == MoveMode.y_first:
            for i in range(int(samples)):
                y = y+step_y
                time.sleep(0.01)
                nx = int(x * 65535 / x_max)
                ny = int(y * 65535 / y_max)
                win32api_mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, nx, ny)
            step_y = 0
        for i in range(int(samples)):
            x, y = x+step_x, y+step_y
            time.sleep(0.01)
            nx = int(x * 65535 / x_max)
            ny = int(y * 65535 / y_max)
            win32api_mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, nx, ny)
    nx = round(xd * 65535 / x_max)
    ny = round(yd * 65535 / y_max)
    win32api_mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, nx, ny)
    if unique_element is None:
        return None
    unique_element_old = unique_element
    element_path_old = element_path2
    w_rOLD = w_r
    return unique_element


def click(
        element_path: UI_Element,
        duration: Optional[float] = 0.5,
        mode: Enum = MoveMode.linear,
        button: str = 'left',
        timeout: float = 120,
        wait_ready: bool = True) -> UI_Element:
    """
    Clicks on element
    
    :param element_path: element path
    :param duration: duration in seconds of the mouse move
    :param mode: move mouse mode
    :param button: mouse button: 'left','double_left', 'triple_left', 'right'
    :param timeout: period of time in seconds that will be allowed to find the element
    :param wait_ready: if True waits until the element is ready
    :return: Pywinauto wrapper of clicked element
    """
    
    unique_element = move(element_path, duration=duration, mode=mode, timeout=timeout)
    if wait_ready and isinstance(element_path, string_types):
        wait_is_ready_try1(unique_element, timeout=timeout)
    else:
        unique_element = None
    if button == 'left' or button == 'double_left' or button == 'triple_left':
        win32api_mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0)
        time.sleep(.01)
        win32api_mouse_event(MOUSEEVENTF_LEFTUP, 0, 0)
        time.sleep(.1)
    if button == 'double_left' or button == 'triple_left':
        win32api_mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0)
        time.sleep(.01)
        win32api_mouse_event(MOUSEEVENTF_LEFTUP, 0, 0)
        time.sleep(.1)
    if button == 'triple_left':
        win32api_mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0)
        time.sleep(.01)
        win32api_mouse_event(MOUSEEVENTF_LEFTUP, 0, 0)
    if button == 'right':
        win32api_mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0)
        time.sleep(.01)
        win32api_mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0)
        time.sleep(.01)
    return unique_element


def left_click(
        element_path: UI_Element,
        duration: Optional[float] = None,
        mode: Enum = MoveMode.linear,
        timeout: float = 120,
        wait_ready: bool = True) -> UI_Element:
    """
    Left clicks on element
    
    :param element_path: element path
    :param duration: duration in seconds of the mouse move
    :param mode: move mouse mode
    :param timeout: period of time in seconds that will be allowed to find the element
    :param wait_ready: if True waits until the element is ready
    :return: Pywinauto wrapper of clicked element
    """
    if not duration:
        duration = PlayerSettings.mouse_move_duration
    return click(element_path, duration=duration, mode=mode, button='left', timeout=timeout, wait_ready=wait_ready)


def right_click(
        element_path: UI_Element,
        duration: Optional[float] = None,
        mode: Enum = MoveMode.linear,
        timeout: float = 120,
        wait_ready: bool = True) -> UI_Element:
    """
    Right clicks on element
    
    :param element_path: element path
    :param duration: duration in seconds of the mouse move
    :param mode: move mouse mode
    :param timeout: period of time in seconds that will be allowed to find the element
    :param wait_ready: if True waits until the element is ready
    :return: Pywinauto wrapper of clicked element
    """
    if not duration:
        duration = PlayerSettings.mouse_move_duration
    return click(element_path, duration=duration, mode=mode, button='right', timeout=timeout, wait_ready=wait_ready)


def double_left_click(
        element_path: UI_Element,
        duration: Optional[float] = None,
        mode: Enum = MoveMode.linear,
        timeout: float = 120,
        wait_ready: bool = True) -> UI_Element:
    """
    Double left clicks on element
    
    :param element_path: element path
    :param duration: duration in seconds of the mouse move
    :param mode: move mouse mode
    :param timeout: period of time in seconds that will be allowed to find the element
    :param wait_ready: if True waits until the element is ready
    :return: Pywinauto wrapper of clicked element
    """
    
    if not duration:
        duration = PlayerSettings.mouse_move_duration
    return click(element_path, duration=duration, mode=mode, button='double_left', timeout=timeout, wait_ready=wait_ready)


def triple_left_click(
        element_path: UI_Element,
        duration: Optional[float] = None,
        mode: Enum = MoveMode.linear,
        timeout: float = 120,
        wait_ready: bool = True) -> UI_Element:
    """
    Triple left clicks on element
    
    :param element_path: element path
    :param duration: duration in seconds of the mouse move
    :param mode: move mouse mode
    :param timeout: period of time in seconds that will be allowed to find the element
    :param wait_ready: if True waits until the element is ready
    :return: Pywinauto wrapper of clicked element
    """
    if not duration:
        duration = PlayerSettings.mouse_move_duration
    return click(element_path, duration=duration, mode=mode, button='triple_left', timeout=timeout, wait_ready=wait_ready)


def drag_and_drop(
        element_path1: UI_Element,
        element_path2: UI_Element,
        duration: Optional[float] = None,
        mode: Enum = MoveMode.linear,
        timeout: float = 120) -> UI_Element:
    """
    Drags and drop with left button pressed from element_path1 to element_path2.
    
    :param element_path1: element path
    :param element_path2: element path
    :param duration: duration in seconds of the mouse move
    :param mode: move mouse mode
    :param timeout: period of time in seconds that will be allowed to find the element
    :return: Pywinauto wrapper with element_path2
    """
    
    if not duration:
        duration = PlayerSettings.mouse_move_duration
    move(element_path1, duration=duration, mode=mode, timeout=timeout)
    win32api_mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0)
    unique_element = move(element_path2, duration=duration, timeout=timeout)
    win32api_mouse_event(MOUSEEVENTF_LEFTUP, 0, 0)
    return unique_element


def middle_drag_and_drop(
        element_path1: UI_Element,
        element_path2: UI_Element,
        duration: Optional[float] = None,
        mode: Enum = MoveMode.linear,
        timeout: float = 120) -> UI_Element:
    """
    Drags and drop with middle button pressed from element_path1 to element_path2.
    
    :param element_path1: element path
    :param element_path2: element path
    :param duration: duration in seconds of the mouse move
    :param mode: move mouse mode
    :param timeout: period of time in seconds that will be allowed to find the element
    :return: Pywinauto wrapper with element_path2
    """
    if not duration:
        duration = PlayerSettings.mouse_move_duration
    move(element_path1, duration=duration, mode=mode, timeout=timeout)
    win32api_mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0)
    unique_element = move(element_path2, duration=duration, mode=mode, timeout=timeout)
    win32api_mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0)
    return unique_element


def right_drag_and_drop(
        element_path1: UI_Element,
        element_path2: UI_Element,
        duration: Optional[float] = None,
        mode: Enum = MoveMode.linear,
        timeout: float = 120) -> UI_Element:
    """
    Drags and drop with right button pressed from element_path1 to element_path2.
    
    :param element_path1: element path
    :param element_path2: element path
    :param duration: duration in seconds of the mouse move
    :param mode: move mouse mode
    :param timeout: period of time in seconds that will be allowed to find the element
    :return: Pywinauto wrapper with element_path2
    """
    if not duration:
        duration = PlayerSettings.mouse_move_duration
    move(element_path1, duration=duration, mode=mode, timeout=timeout)
    win32api_mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0)
    unique_element = move(element_path2, duration=duration, mode=mode, timeout=timeout)
    win32api_mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0)
    return unique_element


def menu_click(
        element_path: UI_Element,
        menu_path: str,
        duration: Optional[float] = None,
        mode: Enum = MoveMode.linear,
        menu_type: str = 'QT',
        timeout: float = 120) -> UI_Element:
    """
    Clicks on menu item.
    
    :param element_path: element path
    :param menu_path: menu path
    :param duration: duration in seconds of the mouse move
    :param mode: move mouse mode
    :param menu_type: menu type ('QT', 'NPP')
    :param timeout: period of time in seconds that will be allowed to find the element
    :return: Pywinauto wrapper of the clicked item
    """
    if not duration:
        duration = PlayerSettings.mouse_move_duration
    menu_entry_list = menu_path.split(path_separator)
    if menu_type == 'QT':
        menu_entry_list = [''] + menu_entry_list
    else:
        menu_entry_list = ['Application'] + menu_entry_list
    if element_path:
        element_path2 = element_path + path_separator
    else:
        element_path2 = ''
    left_click(element_path2 +
               menu_entry_list[0] + type_separator + 'MenuBar' + path_separator +
               menu_entry_list[1] + type_separator + 'MenuItem', duration=duration, mode=mode, timeout=timeout)
    w = None
    if menu_type == 'QT':
        common_path_old = Region.common_path
        Region.common_path = ''
        for entry in menu_entry_list[2:]:
            w = left_click(type_separator + 'Menu' + path_separator + entry + type_separator + 'MenuItem',
                           duration=duration, mode=mode, timeout=timeout)
        Region.common_path = common_path_old
    else:
        for i, entry in enumerate(menu_entry_list[2:]):
            w = left_click(element_path +
                           menu_entry_list[i - 2] + type_separator + 'Menu' + path_separator +
                           unescape_special_char(entry) + type_separator + 'MenuItem',
                           duration=duration, mode=mode, timeout=timeout)
    return w


def mouse_wheel(steps: int, pause: float = 0) -> None:
    """
    Turns the mouse wheel up or down.
    
    :param steps: number of wheel steps, if positive the mouse wheel is turned up else it is turned down
    :param pause: pause in seconds between each wheel step
    """
    if pause == 0:
        win32api_mouse_event(MOUSEEVENTF_WHEEL, 0, 0, WHEEL_DELTA * steps, 0)
    else:
        for i in range(abs(steps)):
            if steps > 0:
                win32api_mouse_event(MOUSEEVENTF_WHEEL, 0, 0, WHEEL_DELTA, 0)
            else:
                win32api_mouse_event(MOUSEEVENTF_WHEEL, 0, 0, -WHEEL_DELTA, 0)
            time.sleep(pause)


def send_keys(
        str_keys: str,
        pause: float = 0.1,
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
    for r in (('(', '{(}'),  (')', '{)}'), ('+', '{+}')):
        str_keys = str_keys.replace(*r)
    pywinauto.keyboard.send_keys(
        str_keys,
        pause=pause,
        with_spaces=with_spaces,
        with_tabs=with_tabs,
        with_newlines=with_newlines,
        turn_off_numlock=turn_off_numlock,
        vk_packet=vk_packet
    )


def set_combobox(
        element_path: UI_Element,
        value: str,
        duration: Optional[float] = None,
        mode: Enum = MoveMode.linear,
        timeout: float = 120) -> None:
    """
    Sets the value of a combobox.
    
    :param element_path: element path
    :param value: value of the combobox
    :param duration: duration in seconds of the mouse move
    :param mode: move mouse mode
    :param timeout: period of time in seconds that will be allowed to find the element
    """
    if not duration:
        duration = PlayerSettings.mouse_move_duration
    left_click(element_path, duration=duration, mode=mode, timeout=timeout)
    time.sleep(0.9)
    send_keys(value + "{ENTER}")


def set_text(
        element_path: UI_Element,
        value: str,
        duration: Optional[float] = None,
        mode: Enum = MoveMode.linear,
        timeout: float = 120,
        pause: float = 0.1) -> None:
    """
    Sets the value of a combobox.
    
    :param element_path: element path
    :param value: value of the combobox
    :param duration: duration in seconds of the mouse move
    :param mode: move mouse mode
    :param timeout: period of time in seconds that will be allowed to find the element
    :param pause: pause in seconds between each typed key
    """
    if not duration:
        duration = PlayerSettings.mouse_move_duration
    double_left_click(element_path, duration=duration, mode=mode, timeout=timeout)
    send_keys("{VK_CONTROL down}a{VK_CONTROL up}", pause=0)
    time.sleep(0.1)
    send_keys(value + "{ENTER}", pause=pause)


def exists(
        element_path: UI_Element,
        timeout: float = 120) -> UI_Element:
    """
    Tests if en UI_Element exists.
    
    :param element_path: element path
    :param timeout: period of time in seconds that will be allowed to find the element
    :return: Pywinauto wrapper of the found element or None
    """
    try:
        wrapper = find(element_path, timeout=timeout)
        return wrapper
    except TimeoutError as e:
        return None


def select_file(
        element_path: UI_Element,
        full_path: str,
        force_slow_path_typing: bool = False) -> None:
    """
    Opens a dialog box and select a file.
    
    :param element_path: element path
    :param full_path: the full path of the file to select
    :param force_slow_path_typing: if True it will type the path even if the current path of the dialog box is the same
    than the file to select
    """
    import pyperclip
    import pathlib
    p = pathlib.Path(full_path)
    folder = p.parent
    filename = p.name
    with Region(element_path, regex_title=True):
        left_click(find().descendants(title="All locations", control_type="SplitButton")[0])
        if not force_slow_path_typing:
            send_keys("{VK_CONTROL down}""c""{VK_CONTROL up}")
        if force_slow_path_typing or pathlib.Path(pyperclip.paste()) != folder:
            send_keys(str(folder))
        send_keys("{ENTER}")
        double_left_click(u"File name:||ComboBox->File name:||Edit")
        send_keys(filename + "{ENTER}")

