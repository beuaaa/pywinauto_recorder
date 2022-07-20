#!/usr/bin/env python3

# print('__file__={0:<35} | __name__={1:<20} | __package__={2:<20}'.format(__file__, __name__, str(__package__)))

# import comtypes.client
# comtypes.client.gen_dir = None
# import pywinauto

from ctypes import wintypes, windll, Structure, POINTER, pointer, WinError
import argparse
import overlay_arrows_and_more as oaam
from os.path import isfile as os_path_isfile
from sys import exit as sys_exit
from pathlib import Path
import os
import win32con
import win32gui_struct
import win32gui
import win32ui
import pyperclip
from pywinauto_recorder.recorder import Recorder, IconSet, _overlay_add_mode_icon, _overlay_add_progress_icon
from win32api import GetSystemMetrics, GetCursorPos
import time


def get_AC_line_satus():
	class SYSTEM_POWER_STATUS(Structure):
		_fields_ = [
			('ACLineStatus', wintypes.BYTE),
			('BatteryFlag', wintypes.BYTE),
			('BatteryLifePercent', wintypes.BYTE),
			('Reserved1', wintypes.BYTE),
			('BatteryLifeTime', wintypes.DWORD),
			('BatteryFullLifeTime', wintypes.DWORD),
		]
	SYSTEM_POWER_STATUS_P = POINTER(SYSTEM_POWER_STATUS)
	GetSystemPowerStatus = windll.kernel32.GetSystemPowerStatus
	GetSystemPowerStatus.argtypes = [SYSTEM_POWER_STATUS_P]
	GetSystemPowerStatus.restype = wintypes.BOOL
	status = SYSTEM_POWER_STATUS()
	if not GetSystemPowerStatus(pointer(status)):
		raise WinError()

	return status.ACLineStatus == 1


class SysTrayIcon(object):
	QUIT = 'QUIT'
	SPECIAL_ACTIONS = [QUIT]
	FIRST_ID = 1023
	
	def __init__(self, icon, hover_text, menu_options, on_quit=None, default_menu_index=None, window_class_name=None, ):
		self.menu = None
		self.icon = icon
		self.hover_text = hover_text
		self.on_quit = on_quit
		
		menu_options = menu_options + [['Quit', icon_power, self.QUIT], ]
		self._next_action_id = self.FIRST_ID
		self.menu_actions_by_id = set()
		self.menu_options = self._add_ids_to_menu_options(list(menu_options))
		self.menu_actions_by_id = dict(self.menu_actions_by_id)
		del self._next_action_id
		
		self.default_menu_index = (default_menu_index or 0)
		self.window_class_name = window_class_name or "SysTrayIconPy"
		
		message_map = {win32gui.RegisterWindowMessage("TaskbarCreated"): self.restart,
		               win32con.WM_DESTROY: self.destroy,
		               win32con.WM_COMMAND: self.command,
		               win32con.WM_USER + 20: self.notify,}
		# Register the Window class.
		window_class = win32gui.WNDCLASS()
		hinst = window_class.hInstance = win32gui.GetModuleHandle(None)
		window_class.lpszClassName = self.window_class_name
		window_class.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW
		window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
		window_class.hbrBackground = win32con.COLOR_WINDOW
		window_class.lpfnWndProc = message_map  # could also specify a wndproc.
		classAtom = win32gui.RegisterClass(window_class)
		# Create the Window.
		style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
		self.hwnd = win32gui.CreateWindow(classAtom, self.window_class_name, style, 0, 0,
		                                  win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, 0, 0, hinst, None)
		win32gui.UpdateWindow(self.hwnd)
		self.notify_id = None
		self.refresh_icon()
		win32gui.PumpMessages()

	def _add_ids_to_menu_options(self, menu_options):
		result = []
		for menu_option in menu_options:
			option_text, option_icon, option_action = menu_option
			if callable(option_action) or option_action in self.SPECIAL_ACTIONS:
				self.menu_actions_by_id.add((self._next_action_id, option_action))
				result.append(menu_option + [self._next_action_id, ])
			elif non_string_iterable(option_action):
				result.append([option_text,
				               option_icon,
				               self._add_ids_to_menu_options(option_action),
				               self._next_action_id])
			else:
				print('Unknown item', option_text, option_icon, option_action)
			self._next_action_id += 1
		return result
	
	def refresh_icon(self):
		# Try and find a custom icon
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
		                  win32con.WM_USER + 20, hicon, self.hover_text)
		win32gui.Shell_NotifyIcon(message, self.notify_id)

	def restart(self, hwnd, msg, wparam, lparam):
		self.refresh_icon()

	def destroy(self, hwnd, msg, wparam, lparam):
		print("destroy--->")
		if self.on_quit:
			self.on_quit(self)
		nid = (self.hwnd, 0)
		win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
		win32gui.PostQuitMessage(0)  # Terminate the app.
		print("<----destroy")

	def notify(self, hwnd, msg, wparam, lparam):
		if lparam == win32con.WM_LBUTTONDBLCLK:
			self.execute_menu_option(self.default_menu_index + self.FIRST_ID)
		elif lparam == win32con.WM_RBUTTONUP:
			self.show_menu()
		elif lparam == win32con.WM_LBUTTONUP:
			pass
		
		return True
	
	def show_menu(self):
		self.menu = win32gui.CreatePopupMenu()
		self.create_menu(self.menu, self.menu_options)
		pos = win32gui.GetCursorPos()
		win32gui.SetForegroundWindow(self.hwnd)
		win32gui.TrackPopupMenu(self.menu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, self.hwnd, None)
		win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)

	def create_menu(self, menu, menu_options):
		if len(menu_options) > 9:
			if recorder.mode != "Record":
				menu_options[0][0] = "Start recording\t\tCTRL+ALT+R"
				menu_options[0][1] = icon_record
			else:
				menu_options[0][0] = "Stop recording\t\tCTRL+ALT+R"
				menu_options[0][1] = icon_stop
			if recorder.mode == "Play":
				menu_options[1][0] = "Wait end of replay" # disabled
				menu_options[1][1] = icon_stop
			else:
				menu_options[1][0] = "Start replaying clipboard"
				menu_options[1][1] = icon_play
			if not recorder.mode == "Info":
				menu_options[2][0] = "Start displaying element information\tCTRL+SHIFT+D"
				menu_options[2][1] = icon_search
			else:
				menu_options[2][0] = "Stop displaying element information\tCTRL+SHIFT+D"
				menu_options[2][1] = icon_stop
			if not recorder.smart_mode:
				menu_options[3][0] = "Start Smart mode\t\tCTRL+ALT+S"
				menu_options[3][1] = icon_light_on
			else:
				menu_options[3][0] = "Stop Smart mode\t\tCTRL+ALT+S"
				menu_options[3][1] = icon_stop
			if not recorder.relative_coordinate_mode:
				menu_options[6][2][0][1] = icon_cross
			else:
				menu_options[6][2][0][1] = icon_check
			if not recorder.process_menu_click_mode:
				menu_options[6][2][1][1] = icon_cross
			else:
				menu_options[6][2][1][1] = icon_check
		for option_text, option_icon, option_action, option_id in menu_options[::-1]:
			if option_icon:
				option_icon = self.prep_menu_icon(option_icon)
			if option_id in self.menu_actions_by_id:
				item, _ = win32gui_struct.PackMENUITEMINFO(text=option_text, hbmpItem=option_icon, wID=option_id)
				if option_text == '- - - - - -':
					win32gui.InsertMenu(menu, 0, win32con.MF_SEPARATOR | win32con.MF_BYPOSITION, 0, None)
				else:
					win32gui.InsertMenuItem(menu, 0, 1, item)
			else:
				submenu = win32gui.CreatePopupMenu()
				self.create_menu(submenu, option_action)
				item, _ = win32gui_struct.PackMENUITEMINFO(text=option_text, hbmpItem=option_icon, hSubMenu=submenu)
				if option_text == '- - - - - -':
					win32gui.InsertMenu(menu, 0, win32con.MF_SEPARATOR | win32con.MF_BYPOSITION, 0, None)
				else:
					win32gui.InsertMenuItem(menu, 0, 1, item)

	def prep_menu_icon(self, icon):
		# First load the icon.
		ico_x = GetSystemMetrics(win32con.SM_CXSMICON)
		ico_y = GetSystemMetrics(win32con.SM_CYSMICON)
		hIcon = win32gui.LoadImage(0, icon, win32con.IMAGE_ICON, ico_x, ico_y, win32con.LR_LOADFROMFILE)
		hwndDC = win32gui.GetWindowDC(self.hwnd)
		dc = win32ui.CreateDCFromHandle(hwndDC)
		memDC = dc.CreateCompatibleDC()
		iconBitmap = win32ui.CreateBitmap()
		iconBitmap.CreateCompatibleBitmap(dc, ico_x, ico_y)
		oldBmp = memDC.SelectObject(iconBitmap)
		brush = win32gui.GetSysColorBrush(win32con.COLOR_MENU)
		win32gui.FillRect(memDC.GetSafeHdc(), (0, 0, ico_x, ico_y), brush)
		win32gui.DrawIconEx(memDC.GetSafeHdc(), 0, 0, hIcon, ico_x, ico_y, 0, 0, win32con.DI_NORMAL)
		memDC.SelectObject(oldBmp)
		memDC.DeleteDC()
		win32gui.ReleaseDC(self.hwnd, hwndDC)
		return iconBitmap.GetHandle()
	
	def command(self, hwnd, msg, wparam, lparam):
		print("command--->")
		id = win32gui.LOWORD(wparam)
		self.execute_menu_option(id)
		print("<---command")

	def execute_menu_option(self, id):
		menu_action = self.menu_actions_by_id[id]
		if menu_action == self.QUIT:
			win32gui.DestroyWindow(self.hwnd)
		else:
			menu_action(self)


def non_string_iterable(obj):
	try:
		iter(obj)
	except TypeError:
		return False
	else:
		return not isinstance(obj, str)


def overlay_add_pywinauto_recorder_icon(overlay, x, y):
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


def display_splash_screen():
	import urllib.request
	import urllib.error
	from pywinauto_recorder.__init__ import __version__
	url_version = "https://raw.githubusercontent.com/beuaaa/pywinauto_recorder/master/bin/VersionInfo.rc"
	latest_version = __version__
	try:
		with urllib.request.urlopen(url_version) as f:
			line = f.read().decode('utf-8')
		latest_version = line.split('ProductVersion')[1].split('"')[2]
	except urllib.error.URLError:
		pass
	
	if __version__ == latest_version:
		print("Your Pywinauto recorder is up to date.")
	else:
		print("Pywinauto recorder v" + latest_version + " is available!")

	splash_foreground = oaam.Overlay(transparency=0.0)
	time.sleep(0.2)
	splash_background = oaam.Overlay(transparency=0.1)
	screen_width = GetSystemMetrics(0)
	screen_height = GetSystemMetrics(1)
	nb_band = 24
	line_height = 22.4
	text_lines = [''] * nb_band
	if __version__ == latest_version:
		text_lines[6] = 'Pywinauto recorder ' + __version__
	else:
		text_lines[6] = 'Your version of Pywinauto recorder (' +  __version__ + ') is not up to date (' + latest_version + ' is available!)'
	text_lines[7] = 'by David Pratmarty'
	text_lines[9] = 'Record / Stop                            Smart mode On / Off'
	text_lines[10] = 'CTRL+ALT+R                                      CTRL+ALT+S'
	text_lines[12] = 'Search algorithm speed        Display / Hide information'
	text_lines[13] = '                                                       CTRL+SHIFT+D'
	text_lines[15] = 'Copy python code in clipboard          Click on "Quit" in'
	text_lines[16] = 'CTRL+SHIFT+F                                tray menu to quit'
	text_lines[18] = 'To replay your recorded file you can:                          '
	text_lines[19] = '- Drag and drop it on pywinauto_recorder.exe            '
	text_lines[20] = '- Click on "Start replaying clipboard" in tray menu       '
	text_lines[21] = '- Run "python.exe recorded_file.py"                           '
	
	splash_background.clear_all()
	w, h = 640, 540
	x, y = screen_width / 2 - w / 2, screen_height / 2 - h / 2
	splash_width = w
	splash_left = x
	splash_right = x + w
	splash_top = y
	splash_bottom = y + h
	splash_background.add(
		geometry=oaam.Shape.triangle, thickness=0,
		xyrgb_array=((x, y, 200, 0, 0), (x + w, y, 0, 128, 0), (x + w, y + h, 0, 0, 155)))
	splash_background.add(
		geometry=oaam.Shape.triangle, thickness=0,
		xyrgb_array=((x + w, y + h, 0, 0, 155), (x, y + h, 22, 128, 66), (x, y, 200, 0, 0)))
	
	splash_background.refresh()
	
	mouse_was_splash_screen = False
	x, y = GetCursorPos()
	if (splash_left < x < splash_right) and (splash_top < y < splash_bottom):
		mouse_was_splash_screen = True
		message_to_continue = 'To continue: move the mouse cursor out of this window'
	else:
		message_to_continue = 'To continue: move the mouse cursor over this window'

	py_rec_icon_rect = (splash_left + splash_width / 2 - 200 / 2, splash_top + 30, 200, 100)
	overlay_add_pywinauto_recorder_icon(splash_background, py_rec_icon_rect[0], py_rec_icon_rect[1])
	continue_after_splash_screen = True
	n = 0
	while continue_after_splash_screen:
		splash_foreground.clear_all()
		i = 0
		while i < nb_band:
			if i < 2:
				font_size = 44
			else:
				font_size = 20
			splash_foreground.add(
				x=splash_left, y=splash_top + i * line_height, width=splash_width, height=line_height,
				text=text_lines[i], font_size=font_size, text_color=(254, 255, 255), color=(255, 255, 255),
				geometry=oaam.Shape.rectangle, thickness=0
			)
			i = i + 1
		if n % 6 in [0, 1, 2]:
			_overlay_add_mode_icon(splash_foreground, IconSet.hicon_record, splash_left + 50, splash_top + line_height * 8.7)
		else:
			_overlay_add_mode_icon(splash_foreground, IconSet.hicon_stop, splash_left + 50, splash_top + line_height * 8.7)
		_overlay_add_mode_icon(splash_foreground, IconSet.hicon_light_on, int(splash_right - (52 + 50)), int(splash_top + line_height * 8.7))
		_overlay_add_progress_icon(splash_foreground, n % 6, splash_left + 50, splash_top + line_height * 11.7)
		_overlay_add_mode_icon(splash_foreground, IconSet.hicon_search, int(splash_right - (52 + 50)), splash_top + line_height * 11.7)
		_overlay_add_mode_icon(splash_foreground, IconSet.hicon_clipboard, splash_left + 50, splash_top + line_height * 14.7)
		_overlay_add_mode_icon(splash_foreground, IconSet.hicon_power, int(splash_right - (52 + 50)), splash_top + line_height * 14.7)
		_overlay_add_mode_icon(splash_foreground, IconSet.hicon_play, splash_left + 50, int(splash_top + line_height * 19.3))
		
		splash_foreground.refresh()
		time.sleep(0.4)
		if n % 2 == 0:
			text_lines[23] = message_to_continue
		else:
			text_lines[23] = ''
		continue_after_splash_screen = False
		x, y = GetCursorPos()
		if (splash_left < x < splash_right) and (splash_top < y < splash_bottom):
			if mouse_was_splash_screen:
				continue_after_splash_screen = True
		else:
			if not mouse_was_splash_screen:
				continue_after_splash_screen = True
		n = n + 1
	splash_foreground.clear_all()
	splash_foreground.refresh()
	splash_background.clear_all()
	splash_background.refresh()


if __name__ == '__main__':
	if not get_AC_line_satus():
		windll.user32.MessageBoxW(0, "Not plugging your laptop in the AC Power can dramatically slow down Pywinauto Recorder.",
		                                 "AC line status is off!", win32con.MB_OK | win32con.MB_ICONEXCLAMATION)
		
	windll.user32.ShowWindow(windll.kernel32.GetConsoleWindow(), 6)
	parser = argparse.ArgumentParser()
	parser.add_argument(
		"filename", metavar='path', help="replay a python script", type=str, action='store', nargs='?', default='')
	parser.add_argument(
		"--no_splash_screen", help="Does not display the splash screen", action='store_true')
	args = parser.parse_args()
	recorder = Recorder()
	if args.filename:
		from pywinauto_recorder.player import *
		if os_path_isfile(args.filename):
			print("Replaying: " + args.filename)
			recorder.playback(filename=args.filename)
		else:
			print("Error: file '" + args.filename + "' not found.")
			input("Press Enter to continue...")
		print("Exit")
	else:
		display_splash_screen()
		if "__compiled__" in globals():
			path_icons = Path(__file__).parent.absolute() / Path("Icons")
		else:
			path_icons = Path(__file__).parent.absolute() / Path("pywinauto_recorder") / Path("Icons")
		# print("ICONS PATH: " + str(path_icons))
		icon_pywinauto_recorder = str(path_icons / Path("IconPyRec.ico"))
		icon_record = str(path_icons / Path("record.ico"))
		icon_stop = str(path_icons / Path("stop.ico"))
		icon_play = str(path_icons / Path("play.ico"))
		icon_folder = str(path_icons / Path("folder.ico"))
		icon_search = str(path_icons / Path("search.ico"))
		icon_light_on = str(path_icons / Path("light-on.ico"))
		icon_settings = str(path_icons / Path("settings.ico"))
		icon_check = str(path_icons / Path("check.ico"))
		icon_cross = str(path_icons / Path("cross.ico"))
		icon_power = str(path_icons / Path("power.ico"))
		icon_help = str(path_icons / Path("help.ico"))
		icon_comments = str(path_icons / Path("comments.ico"))
		icon_favourite = str(path_icons / Path("favourite.ico"))
		hover_text = "Pywinauto recorder"
		
		def action_record(sys_tray_icon):
			if recorder.mode == 'Record':
				recorder.stop_recording()
			else:
				recorder.start_recording()

		def action_replay(sys_tray_icon):
			if recorder.mode != "Play":
				recorder.playback(pyperclip.paste())

		def action_display_element_info(sys_tray_icon):
			if recorder.mode == "Stop":
				recorder.mode = "Info"
			else:
				recorder.mode = "Stop"

		def action_smart_mode(sys_tray_icon):
			recorder.smart_mode = not recorder.smart_mode

		def action_relative_coordinates(sys_tray_icon):
			recorder.relative_coordinate_mode = not recorder.relative_coordinate_mode

		def action_process_menu_click(sys_tray_icon):
			recorder.process_menu_click_mode = not recorder.process_menu_click_mode

		def action_open_explorer(sys_tray_icon):
			pywinauto_recorder_path = Path.home() / Path("Pywinauto recorder")
			os.system('explorer "' + str(pywinauto_recorder_path) + '"')

		def action_display_help(sys_tray_icon):
			recorder.mode = "Stop"
			display_splash_screen()

		def action_display_web_site(sys_tray_icon):
			os.system('rundll32 url.dll,FileProtocolHandler "https://pywinauto-recorder.readthedocs.io"')
			
		def hello(sys_tray_icon):
			print("Menu 2")


		menu_options = [['Start recording\t\tCTRL+ALT+R', icon_stop, action_record],
		                ['Replay clipboard', icon_pywinauto_recorder, action_replay],
		                ['Start displaying element info\tCTRL+SHIFT+D', icon_play, action_display_element_info],
		                ['Start Smart mode\t\tCTRL+ALT+S', icon_stop, action_smart_mode],
		                ['- - - - - -', None, hello],
		                ['Open output folder', icon_folder, action_open_explorer],
		                ['Process events', icon_settings, [
			                ['% relative coordinates', icon_cross, action_relative_coordinates],
			                ['menu_click', icon_cross, action_process_menu_click],
			                ['set_text', icon_cross, hello],
			                ['set_combobox', icon_cross, hello], ]],
		                ['- - - - - -', None, hello],
		                ['Help', icon_help, [
			                ['About', icon_comments, action_display_help],
			                ['Web site', icon_favourite, action_display_web_site], ]]
		                ]
		
		def bye(sys_tray_icon):
			print("bye---->")
			recorder.quit()
			while recorder.is_alive():
				time.sleep(0.5)
			print("Exit")
			print("<----bye")

		SysTrayIcon(icon_pywinauto_recorder, hover_text, menu_options, on_quit=bye, default_menu_index=2)
		
		while recorder.is_alive():
			time.sleep(0.5)
		print("Exit2")

	sys_exit(0)

