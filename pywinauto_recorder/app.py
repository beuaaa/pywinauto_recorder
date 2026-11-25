# -*- coding: utf-8 -*-
"""
The Pywinauto Recorder application: the tray icon and its menu, the splash screen, and the replay of a Python script
given on the command line.

It is started by main(), with the command "pywinauto-recorder", with "python -m pywinauto_recorder",
or with pywinauto_recorder.py, the launcher compiled by Nuitka into pywinauto_recorder.exe.
"""

from ctypes import wintypes, windll, Structure, POINTER, pointer, c_ubyte
import argparse
import functools
import re
import urllib.request
import overlay_arrows_and_more as oaam
from pathlib import Path
import os
import sys
import psutil
import win32con
import win32gui_struct
import win32gui
import win32ui
import pyperclip
import threading
import webbrowser
from .recorder import Recorder, RecorderMode, IconSet, _overlay_add_mode_icon, _overlay_add_progress_icon
from .recorder import HOTKEY_RECORD, HOTKEY_SMART_MODE, HOTKEY_ELEMENT_INFO, HOTKEY_COPY_CODE
from win32api import GetSystemMetrics, GetCursorPos
import time


EXIT_OK = 0
"""Exit code: the script was replayed without error, or the recorder was closed."""
EXIT_REPLAY_FAILED = 1
"""Exit code: the replayed script raised an exception."""
EXIT_FILE_NOT_FOUND = 2
"""Exit code: the script to replay was not found (2 is also the exit code of argparse for a wrong command line)."""

ICONS_FOLDER = Path(__file__).parent.absolute() / "Icons"
"""Folder of the icons of the package."""
OUTPUT_FOLDER = Path.home() / "Pywinauto recorder"
"""Folder where the recorder writes the recorded scripts."""
WEB_SITE = "https://pywinauto-recorder.readthedocs.io"
"""Web site of Pywinauto Recorder, opened by the 'Web site' entry of the tray menu."""

AC_LINE_OFFLINE = 0
"""ACLineStatus of SYSTEM_POWER_STATUS when the computer runs on battery (1: on AC power, 255: unknown)."""

_COMMAND_INTERPRETERS = frozenset((
	"cmd.exe", "powershell.exe", "pwsh.exe", "bash.exe", "sh.exe", "wsl.exe", "nu.exe", "tcc.exe", "far.exe"))
"""Programs whose console belongs to the user: the recorder does not minimize it."""


class _SYSTEM_POWER_STATUS(Structure):
	"""The SYSTEM_POWER_STATUS structure filled by GetSystemPowerStatus."""
	# The fields are unsigned: with wintypes.BYTE (signed), the unknown status 255 would be read as -1
	_fields_ = [
		('ACLineStatus', c_ubyte),
		('BatteryFlag', c_ubyte),
		('BatteryLifePercent', c_ubyte),
		('SystemStatusFlag', c_ubyte),
		('BatteryLifeTime', wintypes.DWORD),
		('BatteryFullLifeTime', wintypes.DWORD),
	]


def get_ac_line_status():
	"""
	Returns the AC power status of the computer.

	:return: 0 if the computer runs on battery, 1 if it is on AC power, 255 if the status is unknown,
		or None if Windows cannot give the status
	"""
	get_system_power_status = windll.kernel32.GetSystemPowerStatus
	get_system_power_status.argtypes = [POINTER(_SYSTEM_POWER_STATUS)]
	get_system_power_status.restype = wintypes.BOOL
	status = _SYSTEM_POWER_STATUS()
	if not get_system_power_status(pointer(status)):
		return None
	return status.ACLineStatus


def is_on_battery(ac_line_status):
	"""
	Returns True if the computer runs on battery, according to its AC power status.

	:param ac_line_status: The AC power status returned by get_ac_line_status()
	:return: True only if the status is 'on battery': False when it is unknown
	"""
	return ac_line_status == AC_LINE_OFFLINE


def warn_if_on_battery():
	"""It displays a modal message box if the computer runs on battery, because it can slow down the recorder."""
	if is_on_battery(get_ac_line_status()):
		windll.user32.MessageBoxW(
			0, "Not plugging your laptop in the AC Power can dramatically slow down Pywinauto Recorder.",
			"AC line status is off!", win32con.MB_OK | win32con.MB_ICONEXCLAMATION)


def _console_process_ids():
	"""Returns the IDs of the processes attached to the console of Pywinauto Recorder."""
	size = 16
	while True:
		process_ids = (wintypes.DWORD * size)()
		count = int(windll.kernel32.GetConsoleProcessList(process_ids, size))
		if count <= size:
			return list(process_ids[:count])
		size = count  # the array was too small: try again with the right size


def _other_console_process_names():
	"""Returns the names of the other processes attached to the console of Pywinauto Recorder."""
	names = []
	for process_id in _console_process_ids():
		if process_id == os.getpid():
			continue
		try:
			names.append(psutil.Process(process_id).name())
		except psutil.Error:
			pass
	return names


def should_minimize_console(other_process_names):
	"""
	Returns True if the console can be minimized: no command interpreter is attached to it.

	The console was then created for Pywinauto Recorder (double click on the executable, a shortcut...). When the
	recorder is started from a command interpreter (cmd, PowerShell...), the console belongs to the user and stays as it is.

	:param other_process_names: The names of the other processes attached to the console
	:return: True if the console can be minimized
	"""
	return not any(name.lower() in _COMMAND_INTERPRETERS for name in other_process_names)


def minimize_console_if_owned():
	"""It minimizes the console of Pywinauto Recorder, unless it belongs to a command interpreter."""
	console_window = windll.kernel32.GetConsoleWindow()
	if console_window and should_minimize_console(_other_console_process_names()):
		windll.user32.ShowWindow(console_window, win32con.SW_MINIMIZE)


def _wait_for_enter_in_console():
	"""It waits for the Enter key, so that a message can be read, only if someone can press it in a console."""
	if sys.stdin is not None and sys.stdin.isatty():
		input("Press Enter to continue...")


MENU_SEPARATOR = None
"""A separator line between the entries of a menu."""


class MenuItem(object):
	"""
	An entry of the menu of the tray icon.

	'text', 'icon' and 'enabled' are values, or functions without argument that return the value. They are evaluated
	each time the menu is displayed, so that the menu shows the current state of the recorder.
	"""
	def __init__(self, text, icon=None, action=None, submenu=None, enabled=True):
		"""
		:param text: The text of the entry, or a function that returns it
		:param icon: The path of the .ico file of the entry, or a function that returns it, or None
		:param action: The function called with the SysTrayIcon when the entry is clicked
		:param submenu: The entries of the submenu (a list of MenuItem and MENU_SEPARATOR), or None
		:param enabled: True if the entry can be clicked, or a function that returns it. A disabled entry is grayed.
		"""
		self.text = text
		self.icon = icon
		self.action = action
		self.submenu = submenu
		self.enabled = enabled

	@staticmethod
	def _value(value):
		return value() if callable(value) else value

	def current_text(self):
		"""Returns the text of the entry."""
		return self._value(self.text)

	def current_icon(self):
		"""Returns the path of the icon of the entry, or None."""
		return self._value(self.icon)

	def is_enabled(self):
		"""Returns True if the entry can be clicked."""
		return bool(self._value(self.enabled))


class SysTrayIcon(object):
	"""
	The icon of Pywinauto Recorder in the notification area of the taskbar, and its menu.

	It creates a hidden window that receives the messages of the icon. run() processes these messages until quit() is called.
	"""
	FIRST_ID = 1023
	_WM_NOTIFY_ICON = win32con.WM_USER + 20

	def __init__(self, icon, hover_text, menu, on_quit=None, default_action=None, window_class_name="SysTrayIconPy"):
		"""
		:param icon: The path of the .ico file of the icon
		:param hover_text: The text displayed when the mouse cursor is over the icon
		:param menu: The entries of the menu: a list of MenuItem and MENU_SEPARATOR
		:param on_quit: The function called with the SysTrayIcon when it is destroyed
		:param default_action: The function called with the SysTrayIcon on a double click on the icon
		:param window_class_name: The name of the class of the hidden window
		"""
		self.icon = icon
		self.hover_text = hover_text
		self.menu = menu
		self.on_quit = on_quit
		self.default_action = default_action
		self._actions_by_id = {}
		self._menu_bitmaps = {}

		# The handlers return 0, the result of the window procedure: recent versions of pywin32 refuse None, and print
		# "WNDPROC return value cannot be converted to LRESULT"
		message_map = {win32gui.RegisterWindowMessage("TaskbarCreated"): self.restart,
		               win32con.WM_DESTROY: self.destroy,
		               win32con.WM_COMMAND: self.command,
		               self._WM_NOTIFY_ICON: self.notify}
		# Register the Window class.
		window_class = win32gui.WNDCLASS()
		hinst = window_class.hInstance = win32gui.GetModuleHandle(None)
		window_class.lpszClassName = window_class_name
		window_class.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW
		window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
		window_class.hbrBackground = win32con.COLOR_WINDOW
		window_class.lpfnWndProc = message_map  # could also specify a wndproc.
		class_atom = win32gui.RegisterClass(window_class)
		# Create the Window.
		style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
		self.hwnd = win32gui.CreateWindow(class_atom, window_class_name, style, 0, 0,
		                                  win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, 0, 0, hinst, None)
		win32gui.UpdateWindow(self.hwnd)
		self.notify_id = None
		self.refresh_icon()

	def run(self):
		"""It processes the messages of the icon and of its menu until quit() is called."""
		win32gui.PumpMessages()

	def quit(self):
		"""It removes the icon, calls on_quit and stops run()."""
		win32gui.DestroyWindow(self.hwnd)

	def refresh_icon(self):
		"""It adds the icon to the notification area, or updates it."""
		hinst = win32gui.GetModuleHandle(None)
		if os.path.isfile(self.icon):
			icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
			hicon = win32gui.LoadImage(hinst, self.icon, win32con.IMAGE_ICON, 0, 0, icon_flags)
		else:
			print("Can't find icon file - using default.")
			hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
		if self.notify_id:
			message = win32gui.NIM_MODIFY
		else:
			message = win32gui.NIM_ADD
		self.notify_id = (self.hwnd, 0, win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
		                  self._WM_NOTIFY_ICON, hicon, self.hover_text)
		win32gui.Shell_NotifyIcon(message, self.notify_id)

	def restart(self, hwnd, msg, wparam, lparam):
		"""Called when the taskbar is created again (Windows Explorer restarted): it adds the icon again."""
		self.refresh_icon()
		return 0

	def destroy(self, hwnd, msg, wparam, lparam):
		"""
		Called when the hidden window is destroyed: it calls on_quit, removes the icon and stops run().
		The icon is removed and run() is stopped even if on_quit raises an exception.
		"""
		try:
			if self.on_quit:
				self.on_quit(self)
		finally:
			nid = (self.hwnd, 0)
			win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
			win32gui.PostQuitMessage(0)  # Terminate the app.
		return 0

	def notify(self, hwnd, msg, wparam, lparam):
		"""Called on a mouse event on the icon: a double click calls default_action, a right click shows the menu."""
		if lparam == win32con.WM_LBUTTONDBLCLK:
			if self.default_action:
				self.default_action(self)
		elif lparam == win32con.WM_RBUTTONUP:
			self.show_menu()
		return 0

	def show_menu(self):
		"""It shows the menu at the position of the mouse cursor."""
		self._actions_by_id = {}
		menu = win32gui.CreatePopupMenu()
		self._fill_menu(menu, self.menu)
		pos = win32gui.GetCursorPos()
		win32gui.SetForegroundWindow(self.hwnd)
		win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, self.hwnd, None)
		win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
		# The command of the clicked entry is posted: the menu, and its submenus, are no longer needed
		win32gui.DestroyMenu(menu)

	def _fill_menu(self, menu, items):
		"""
		It adds the entries to a native menu, with their current text, icon and state.

		:param menu: The handle of the native menu
		:param items: A list of MenuItem and MENU_SEPARATOR
		"""
		for item in reversed(items):  # each entry is inserted at the top of the menu
			if item is MENU_SEPARATOR:
				win32gui.InsertMenu(menu, 0, win32con.MF_SEPARATOR | win32con.MF_BYPOSITION, 0, None)
				continue
			icon = item.current_icon()
			bitmap = self._menu_bitmap(icon) if icon else None
			state = None if item.is_enabled() else win32con.MFS_GRAYED
			if item.submenu is not None:
				submenu = win32gui.CreatePopupMenu()
				self._fill_menu(submenu, item.submenu)
				menu_item_info, extras = win32gui_struct.PackMENUITEMINFO(
					text=item.current_text(), hbmpItem=bitmap, hSubMenu=submenu, fState=state)
			else:
				item_id = self.FIRST_ID + len(self._actions_by_id)
				self._actions_by_id[item_id] = item.action
				menu_item_info, extras = win32gui_struct.PackMENUITEMINFO(
					text=item.current_text(), hbmpItem=bitmap, wID=item_id, fState=state)
			win32gui.InsertMenuItem(menu, 0, 1, menu_item_info)

	def _menu_bitmap(self, icon):
		"""
		Returns the handle of the bitmap of an icon of the menu. Each icon is loaded once, then kept for the next menus.

		:param icon: The path of the .ico file
		"""
		bitmap = self._menu_bitmaps.get(icon)
		if bitmap is None:
			bitmap = self._load_menu_bitmap(icon)
			self._menu_bitmaps[icon] = bitmap  # keeps the bitmap object, so that its handle stays valid
		return bitmap.GetHandle()

	def _load_menu_bitmap(self, icon):
		"""Returns a bitmap object with the icon drawn on the background color of the menus."""
		ico_x = GetSystemMetrics(win32con.SM_CXSMICON)
		ico_y = GetSystemMetrics(win32con.SM_CYSMICON)
		hicon = win32gui.LoadImage(0, icon, win32con.IMAGE_ICON, ico_x, ico_y, win32con.LR_LOADFROMFILE)
		hwnd_dc = win32gui.GetWindowDC(self.hwnd)
		dc = win32ui.CreateDCFromHandle(hwnd_dc)
		mem_dc = dc.CreateCompatibleDC()
		icon_bitmap = win32ui.CreateBitmap()
		icon_bitmap.CreateCompatibleBitmap(dc, ico_x, ico_y)
		old_bitmap = mem_dc.SelectObject(icon_bitmap)
		brush = win32gui.GetSysColorBrush(win32con.COLOR_MENU)
		win32gui.FillRect(mem_dc.GetSafeHdc(), (0, 0, ico_x, ico_y), brush)
		win32gui.DrawIconEx(mem_dc.GetSafeHdc(), 0, 0, hicon, ico_x, ico_y, 0, 0, win32con.DI_NORMAL)
		mem_dc.SelectObject(old_bitmap)
		mem_dc.DeleteDC()
		win32gui.ReleaseDC(self.hwnd, hwnd_dc)
		win32gui.DestroyIcon(hicon)
		return icon_bitmap

	def command(self, hwnd, msg, wparam, lparam):
		"""Called when an entry of the menu is clicked: it calls its action."""
		action = self._actions_by_id.get(win32gui.LOWORD(wparam))
		if action is not None:
			action(self)
		return 0


class TrayIcons(object):
	"""The paths of the icons of the tray icon and of its menu."""
	def __init__(self, folder=ICONS_FOLDER):
		"""
		:param folder: The folder of the .ico files
		"""
		self.pywinauto_recorder = str(folder / "IconPyRec.ico")
		self.record = str(folder / "record.ico")
		self.stop = str(folder / "stop.ico")
		self.play = str(folder / "play.ico")
		self.folder = str(folder / "folder.ico")
		self.search = str(folder / "search.ico")
		self.light_on = str(folder / "light-on.ico")
		self.settings = str(folder / "settings.ico")
		self.check = str(folder / "check.ico")
		self.cross = str(folder / "cross.ico")
		self.power = str(folder / "power.ico")
		self.help = str(folder / "help.ico")
		self.comments = str(folder / "comments.ico")
		self.favourite = str(folder / "favourite.ico")


class ClipboardReplay(object):
	"""
	It replays the Python code of the clipboard in a thread, so that the tray menu keeps working during the replay.
	"""
	def __init__(self, recorder):
		"""
		:param recorder: The Recorder that replays the code: it displays the 'play' icon during the replay
		"""
		self.recorder = recorder
		self._thread = None

	def is_running(self):
		"""Returns True while a replay is in progress."""
		return self._thread is not None and self._thread.is_alive()

	def start(self, code):
		"""
		It starts replaying Python code, unless a replay is already in progress.

		:param code: The Python code to replay
		:return: True if the replay was started, False if a replay is already in progress
		"""
		if self.is_running():
			return False
		self._thread = threading.Thread(
			target=self.recorder.playback, kwargs={"str_code": code}, name="Pywinauto Recorder replay", daemon=True)
		self._thread.start()
		return True


def toggle_recording(recorder):
	"""It starts recording, or stops the recording in progress and writes the recorded script."""
	if recorder.mode == RecorderMode.RECORD:
		recorder.stop_recording()
	else:
		recorder.start_recording()


def show_about(recorder):
	"""It stops the recorder, after writing the recording in progress, and displays the splash screen."""
	if recorder.mode == RecorderMode.RECORD:
		recorder.stop_recording()
	recorder.mode = RecorderMode.STOP
	display_splash_screen()


def open_output_folder():
	"""It opens the folder of the recorded scripts in the Windows Explorer, after creating it if needed."""
	OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
	os.startfile(str(OUTPUT_FOLDER))


def open_web_site():
	"""It opens the web site of Pywinauto Recorder in the default browser."""
	webbrowser.open(WEB_SITE)


def build_menu(recorder, clipboard_replay, icons):
	"""
	Returns the entries of the tray menu.

	Their texts, icons and states are evaluated each time the menu is displayed. During a replay of the clipboard,
	the entries that would interfere with the replay are grayed, "Quit" included.

	:param recorder: The Recorder
	:param clipboard_replay: The ClipboardReplay
	:param icons: The TrayIcons
	:return: A list of MenuItem and MENU_SEPARATOR
	"""
	def no_replay():
		return not clipboard_replay.is_running()

	def recording():
		return recorder.mode == RecorderMode.RECORD

	def displaying_info():
		return recorder.mode == RecorderMode.INFO

	return [
		MenuItem(
			lambda: ("Stop recording" if recording() else "Start recording") + "\t\t" + HOTKEY_RECORD,
			lambda: icons.stop if recording() else icons.record,
			action=lambda tray: toggle_recording(recorder), enabled=no_replay),
		MenuItem(
			lambda: "Start replaying clipboard" if no_replay() else "Wait end of replay",
			lambda: icons.play if no_replay() else icons.stop,
			action=lambda tray: clipboard_replay.start(pyperclip.paste()), enabled=no_replay),
		MenuItem(
			lambda: ("Stop" if displaying_info() else "Start") + " displaying element information\t" + HOTKEY_ELEMENT_INFO,
			lambda: icons.stop if displaying_info() else icons.search,
			action=lambda tray: recorder.toggle_element_info(), enabled=no_replay),
		MenuItem(
			lambda: ("Stop" if recorder.smart_mode else "Start") + " Smart mode\t\t" + HOTKEY_SMART_MODE,
			lambda: icons.stop if recorder.smart_mode else icons.light_on,
			action=lambda tray: setattr(recorder, "smart_mode", not recorder.smart_mode), enabled=no_replay),
		MENU_SEPARATOR,
		MenuItem("Open output folder", icons.folder, action=lambda tray: open_output_folder()),
		MenuItem("Process events", icons.settings, submenu=[
			MenuItem(
				"% relative coordinates", lambda: icons.check if recorder.relative_coordinate_mode else icons.cross,
				action=lambda tray: setattr(recorder, "relative_coordinate_mode", not recorder.relative_coordinate_mode)),
			MenuItem(
				"menu_click", lambda: icons.check if recorder.process_menu_click_mode else icons.cross,
				action=lambda tray: setattr(recorder, "process_menu_click_mode", not recorder.process_menu_click_mode)),
			MenuItem("set_text (not implemented yet)", icons.cross, enabled=False),
			MenuItem("set_combobox (not implemented yet)", icons.cross, enabled=False)]),
		MENU_SEPARATOR,
		MenuItem("Help", icons.help, submenu=[
			MenuItem("About", icons.comments, action=lambda tray: show_about(recorder), enabled=no_replay),
			MenuItem("Web site", icons.favourite, action=lambda tray: open_web_site())]),
		MenuItem("Quit", icons.power, action=lambda tray: tray.quit(), enabled=no_replay),
	]


LATEST_VERSION_URL = "https://raw.githubusercontent.com/beuaaa/pywinauto_recorder/master/bin/VersionInfo.rc"
"""bin/VersionInfo.rc of the branch 'master', which receives only the published versions: it gives the latest one."""
VERSION_CHECK_TIMEOUT = 5
"""Maximum duration, in seconds, of the request for the latest published version."""

_VERSION_PATTERN = re.compile(r"v?(\d+(?:\.\d+)*)[._-]?([a-z]*)[._-]?(\d*)")
_PRE_RELEASE_RANKS = {"dev": 0, "a": 1, "alpha": 1, "b": 2, "beta": 2, "c": 3, "rc": 3}
_FINAL_RELEASE_RANK = 4


def get_latest_version(timeout=VERSION_CHECK_TIMEOUT):
	"""
	Returns the latest published version of Pywinauto Recorder.

	:param timeout: The maximum duration of the request, in seconds
	:return: The version read in LATEST_VERSION_URL, or None if it cannot be read (no network, timeout, proxy, unexpected
		content...)
	"""
	try:
		with urllib.request.urlopen(LATEST_VERSION_URL, timeout=timeout) as response:
			text = response.read().decode('utf-8')
		version = text.split('ProductVersion')[1].split('"')[2].strip()
	except Exception:
		return None
	return version if _VERSION_PATTERN.match(version.lower()) else None


def version_key(version):
	"""
	Returns a key to compare versions: "0.6.9" < "0.10.0" < "1.0.0.dev0" < "1.0.0b1" < "1.0.0rc1" < "1.0.0" == "1.0".

	:param version: A version, for example "1.0.0b1"
	:return: A tuple: the numbers of the version without the trailing zeros, the rank of the pre-release (a final
		release comes after its pre-releases), and the number of the pre-release
	"""
	match = _VERSION_PATTERN.match(version.strip().lower())
	if not match:
		return (), -1, 0
	numbers = [int(number) for number in match.group(1).split(".")]
	while len(numbers) > 1 and numbers[-1] == 0:
		numbers.pop()  # 1.0.0 is the same version as 1.0
	label = match.group(2)
	rank = _PRE_RELEASE_RANKS.get(label, _FINAL_RELEASE_RANK) if label else _FINAL_RELEASE_RANK
	return tuple(numbers), rank, int(match.group(3) or 0)


def version_messages(installed_version, latest_version):
	"""
	Returns the messages about the version of Pywinauto Recorder.

	:param installed_version: The version of this Pywinauto Recorder
	:param latest_version: The latest published version, or None if it is unknown
	:return: The message printed in the console, and the line displayed by the splash screen
	"""
	version_line = 'Pywinauto recorder ' + installed_version
	if latest_version is None:
		return "Cannot check whether a newer version of Pywinauto recorder is available.", version_line
	if version_key(installed_version) == version_key(latest_version):
		return "Your Pywinauto recorder is up to date.", version_line
	if version_key(installed_version) < version_key(latest_version):
		return ("Pywinauto recorder v" + latest_version + " is available!",
		        'Your version of Pywinauto recorder (' + installed_version + ') is not up to date (' +
		        latest_version + ' is available!)')
	return ("Your Pywinauto recorder " + installed_version + " is newer than the latest published version (" +
	        latest_version + ")."), version_line


@functools.lru_cache(maxsize=None)
def check_version():
	"""
	It checks whether a newer version of Pywinauto Recorder is available, and prints the result in the console.
	The check is done only once: the next calls return the same line, without printing it again.

	:return: The line about the version displayed by the splash screen
	"""
	from . import __version__
	console_message, version_line = version_messages(__version__, get_latest_version())
	print(console_message)
	return version_line


_SPLASH_WIDTH = 640
_SPLASH_HEIGHT = 540
_SPLASH_NB_LINES = 24
_SPLASH_LINE_HEIGHT = 22.4


def _splash_screen_text_lines(version_line):
	"""
	Returns the lines of text of the splash screen, with the keyboard shortcuts.

	:param version_line: The line about the version of Pywinauto Recorder
	:return: A list of 24 lines. The last one, the message to continue, is set by display_splash_screen.
	"""
	text_lines = [''] * _SPLASH_NB_LINES
	text_lines[6] = version_line
	text_lines[7] = 'by David Pratmarty'
	text_lines[9] = 'Record / Stop                            Smart mode On / Off'
	text_lines[10] = HOTKEY_RECORD + ' ' * 38 + HOTKEY_SMART_MODE
	text_lines[12] = 'Search algorithm speed        Display / Hide information'
	text_lines[13] = ' ' * 55 + HOTKEY_ELEMENT_INFO
	text_lines[15] = 'Copy python code in clipboard          Click on "Quit" in'
	text_lines[16] = HOTKEY_COPY_CODE + ' ' * 32 + 'tray menu to quit'
	text_lines[18] = 'To replay your recorded file you can:                          '
	text_lines[19] = '- Drag and drop it on pywinauto_recorder.exe            '
	text_lines[20] = '- Click on "Start replaying clipboard" in tray menu       '
	text_lines[21] = '- Run "python.exe recorded_file.py"                           '
	return text_lines


def _splash_rectangle():
	"""Returns the rectangle (left, top, right, bottom) of the splash screen, centered on the screen."""
	left = GetSystemMetrics(0) / 2 - _SPLASH_WIDTH / 2
	top = GetSystemMetrics(1) / 2 - _SPLASH_HEIGHT / 2
	return left, top, left + _SPLASH_WIDTH, top + _SPLASH_HEIGHT


def _is_mouse_cursor_in(rectangle):
	"""Returns True if the mouse cursor is in the rectangle (left, top, right, bottom)."""
	left, top, right, bottom = rectangle
	x, y = GetCursorPos()
	return left < x < right and top < y < bottom


def overlay_add_pywinauto_recorder_icon(overlay, x, y):
	"""It adds the logo of Pywinauto Recorder, 200x100 pixels, to the overlay at (x, y)."""
	overlay.add(
		geometry=oaam.Shape.rectangle, x=x, y=y, width=200, height=100, thickness=5, color=(200, 66, 66),
		brush=oaam.Brush.solid, brush_color=(255, 254, 255))
	overlay.add(
		geometry=oaam.Shape.rectangle, x=x + 5, y=y + 15, width=190, height=38, thickness=0,
		brush=oaam.Brush.solid, brush_color=(255, 254, 255), text_color=(66, 66, 66),
		text=u'PYWINAUTO', font_size=44, font_name='Impact')
	overlay.add(
		geometry=oaam.Shape.rectangle, x=x + 20, y=y + 50, width=160, height=38, thickness=0,
		brush=oaam.Brush.solid, brush_color=(255, 254, 255), text_color=(200, 40, 40),
		text=u'recorder', font_size=48, font_name='Arial Black')


def _draw_splash_background(overlay, rectangle):
	"""It draws the background of the splash screen, and the logo of Pywinauto Recorder."""
	x, y, right, bottom = rectangle
	w, h = right - x, bottom - y
	overlay.clear_all()
	overlay.add(
		geometry=oaam.Shape.triangle, thickness=0,
		xyrgb_array=((x, y, 200, 0, 0), (x + w, y, 0, 128, 0), (x + w, y + h, 0, 0, 155)))
	overlay.add(
		geometry=oaam.Shape.triangle, thickness=0,
		xyrgb_array=((x + w, y + h, 0, 0, 155), (x, y + h, 22, 128, 66), (x, y, 200, 0, 0)))
	overlay.refresh()
	overlay_add_pywinauto_recorder_icon(overlay, x + w / 2 - 200 / 2, y + 30)


def _draw_splash_foreground(overlay, rectangle, text_lines, n):
	"""
	It draws the texts and the icons of the splash screen.

	:param n: The number of the frame: the record/stop icon and the progress icon are animated
	"""
	splash_left, splash_top, splash_right, _ = rectangle
	line_height = _SPLASH_LINE_HEIGHT
	overlay.clear_all()
	for i, text_line in enumerate(text_lines):
		overlay.add(
			x=splash_left, y=splash_top + i * line_height, width=_SPLASH_WIDTH, height=line_height,
			text=text_line, font_size=44 if i < 2 else 20, text_color=(254, 255, 255), color=(255, 255, 255),
			geometry=oaam.Shape.rectangle, thickness=0
		)
	if n % 6 in [0, 1, 2]:
		_overlay_add_mode_icon(overlay, IconSet.hicon_record, splash_left + 50, splash_top + line_height * 8.7)
	else:
		_overlay_add_mode_icon(overlay, IconSet.hicon_stop, splash_left + 50, splash_top + line_height * 8.7)
	_overlay_add_mode_icon(overlay, IconSet.hicon_light_on, int(splash_right - (52 + 50)), int(splash_top + line_height * 8.7))
	_overlay_add_progress_icon(overlay, n % 6, splash_left + 50, splash_top + line_height * 11.7)
	_overlay_add_mode_icon(overlay, IconSet.hicon_search, int(splash_right - (52 + 50)), splash_top + line_height * 11.7)
	_overlay_add_mode_icon(overlay, IconSet.hicon_clipboard, splash_left + 50, splash_top + line_height * 14.7)
	_overlay_add_mode_icon(overlay, IconSet.hicon_power, int(splash_right - (52 + 50)), splash_top + line_height * 14.7)
	_overlay_add_mode_icon(overlay, IconSet.hicon_play, splash_left + 50, int(splash_top + line_height * 19.3))
	overlay.refresh()


def display_splash_screen():
	"""
	It displays the splash screen, with the version and the keyboard shortcuts, until the mouse cursor moves out of it.
	If the mouse cursor is not over the splash screen when it is displayed, it waits until the mouse cursor moves over it.
	"""
	text_lines = _splash_screen_text_lines(check_version())
	splash_foreground = oaam.Overlay(transparency=0.0)
	time.sleep(0.2)
	splash_background = oaam.Overlay(transparency=0.1)
	try:
		rectangle = _splash_rectangle()
		_draw_splash_background(splash_background, rectangle)
		mouse_was_in_splash_screen = _is_mouse_cursor_in(rectangle)
		if mouse_was_in_splash_screen:
			message_to_continue = 'To continue: move the mouse cursor out of this window'
		else:
			message_to_continue = 'To continue: move the mouse cursor over this window'
		n = 0
		while True:
			_draw_splash_foreground(splash_foreground, rectangle, text_lines, n)
			time.sleep(0.4)
			text_lines[_SPLASH_NB_LINES - 1] = message_to_continue if n % 2 == 0 else ''  # the message blinks
			n = n + 1
			if _is_mouse_cursor_in(rectangle) != mouse_was_in_splash_screen:
				break
	finally:
		# The overlays are closed: before, they were only cleared, and two more stayed open at each display ("About").
		# quit() is reliable since overlay_arrows_and_more 0.6.0, even just after the creation of an overlay.
		splash_foreground.quit()
		splash_background.quit()


def _program_name():
	"""Returns the name of the program displayed by the help of the command line."""
	name = os.path.basename(sys.argv[0]) if sys.argv and sys.argv[0] else ""
	if not name or name == "__main__.py":
		return "python -m pywinauto_recorder"
	return name


def parse_arguments(argv=None):
	"""
	It parses the command line.

	:param argv: The command line arguments, without the program name. If None, sys.argv[1:] is used.
	:return: The parsed arguments: 'filename' (the script to replay, or '') and 'no_splash_screen'
	"""
	parser = argparse.ArgumentParser(
		prog=_program_name(),
		description="Records user interface actions in a Python script, or replays a Python script.")
	parser.add_argument(
		"filename", metavar='path', help="replay a python script", type=str, action='store', nargs='?', default='')
	parser.add_argument(
		"--no_splash_screen", help="Does not display the splash screen", action='store_true')
	return parser.parse_args(argv)


def replay_file(filename):
	"""
	It replays a Python script. The recorder displays the 'play' icon during the replay.

	:param filename: The path of the Python script
	:return: EXIT_OK if the script was replayed without error, EXIT_REPLAY_FAILED if it raised an exception,
		EXIT_FILE_NOT_FOUND if the file does not exist
	"""
	if not os.path.isfile(filename):
		print("Error: file '" + filename + "' not found.")
		_wait_for_enter_in_console()
		return EXIT_FILE_NOT_FOUND
	recorder = Recorder()
	print("Replaying: " + filename)
	try:
		succeeded = recorder.playback(filename=filename)
	finally:
		recorder.quit()  # before, the process of the element information tooltip kept the application running
	print("Exit")
	return EXIT_OK if succeeded else EXIT_REPLAY_FAILED


def run_interactive(args):
	"""
	It starts the recorder with its tray icon, and waits until the user clicks on "Quit" in the tray menu.

	:param args: The parsed command line arguments
	:return: EXIT_OK
	"""
	recorder = Recorder()
	if not args.no_splash_screen:
		display_splash_screen()
	icons = TrayIcons()
	clipboard_replay = ClipboardReplay(recorder)

	def double_click(tray):
		if not clipboard_replay.is_running():
			recorder.toggle_element_info()

	tray = SysTrayIcon(
		icons.pywinauto_recorder, "Pywinauto recorder", build_menu(recorder, clipboard_replay, icons),
		on_quit=lambda tray: recorder.quit(), default_action=double_click)
	tray.run()
	return EXIT_OK


def main(argv=None):
	"""
	Starts Pywinauto Recorder: it replays the Python script given on the command line, or else it starts the recorder.

	:param argv: The command line arguments, without the program name. If None, sys.argv[1:] is used.
	:return: The exit code of the program: EXIT_OK, EXIT_REPLAY_FAILED or EXIT_FILE_NOT_FOUND
	"""
	args = parse_arguments(argv)
	warn_if_on_battery()
	minimize_console_if_owned()
	if args.filename:
		return replay_file(args.filename)
	return run_interactive(args)
