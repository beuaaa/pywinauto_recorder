#!/usr/bin/env python3

# print('__file__={0:<35} | __name__={1:<20} | __package__={2:<20}'.format(__file__, __name__, str(__package__)))

# import comtypes.client
# comtypes.client.gen_dir = None
# import pywinauto

from ctypes import windll
import argparse
import overlay_arrows_and_more as oaam
from os.path import isfile as os_path_isfile
from sys import exit as sys_exit


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
	dir_path = os.path.dirname(os.path.realpath(__file__))
	hicon_light_on = oaam.load_png(dir_path + r'\pywinauto_recorder\light_on.png', 41, 41)
	hicon_light_off = oaam.load_png(dir_path + r'\pywinauto_recorder\light_off.png', 41, 41)
	
	splash_foreground = oaam.Overlay(transparency=0.0)
	time.sleep(0.2)
	splash_background = oaam.Overlay(transparency=0.1)
	screen_width = GetSystemMetrics(0)
	screen_height = GetSystemMetrics(1)
	nb_band = 24
	line_height = 22.5
	text_lines = [''] * nb_band
	text_lines[6] = 'Pywinauto recorder ' + __version__
	text_lines[7] = 'by David Pratmarty'
	text_lines[9] = 'CTRL+ALT+R : Pause / Record / Stop'
	text_lines[11] = 'Search algorithm speed'
	text_lines[13] = 'CTRL+SHIFT+S : Smart mode On / Off'
	text_lines[15] = 'CTRL+SHIFT+F : Copy element path in clipboard'
	text_lines[17] = 'CTRL+ALT+Q : Quit'
	text_lines[19] = 'Drag and drop a recorded file on '
	text_lines[20] = 'pywinauto_recorder.exe to replay it'
	
	splash_background.clear_all()
	w, h = 640, 540
	x, y = screen_width / 2 - w / 2, screen_height / 2 - h / 2
	splash_width = w
	splash_height = h
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
	x, y = win32api.GetCursorPos()
	if (splash_left < x < splash_right) and (splash_top < y < splash_bottom):
		mouse_was_splash_screen = True
		message_to_continue = 'To continue: move the mouse cursor out of this splash screen'
	else:
		message_to_continue = 'To continue: move the mouse cursor over this splash screen'
	
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
		if n % 6 in [0, 1]:
			overlay_add_record_icon(
				splash_foreground, splash_left + 99, splash_top + line_height * 8.6)
		if n % 6 in [2, 3]:
			overlay_add_pause_icon(
				splash_foreground, splash_left + 99, splash_top + line_height * 8.6)
		if n % 6 in [4, 5]:
			overlay_add_stop_icon(
				splash_foreground, splash_left + 99, splash_top + line_height * 8.6)
		overlay_add_progress_icon(
			splash_foreground, n % 5, splash_left + 99, splash_top + line_height * 10.6)
		overlay_add_play_icon(
			splash_foreground, splash_left + 99, int(splash_top + line_height * 19.1))
		if n % 4 == 0 or n % 4 == 1:
			overlay_add_search_mode_icon(splash_foreground, hicon_light_off, int(splash_left + 99),
			                             int(splash_top + line_height * 12.6))
		else:
			overlay_add_search_mode_icon(splash_foreground, hicon_light_on, int(splash_left + 99),
			                             int(splash_top + line_height * 12.6))
		splash_foreground.refresh()
		time.sleep(0.4)
		if n % 2 == 0:
			text_lines[22] = message_to_continue
		else:
			text_lines[22] = ''
		continue_after_splash_screen = False
		x, y = win32api.GetCursorPos()
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
	windll.user32.ShowWindow(windll.kernel32.GetConsoleWindow(), 6)
	parser = argparse.ArgumentParser()
	parser.add_argument(
		"filename", metavar='path', help="replay a python script", type=str, action='store', nargs='?', default='')
	parser.add_argument(
		"--no_splash_screen", help="Does not display the splash screen", action='store_true')
	args = parser.parse_args()
	if args.filename:
		main_overlay = oaam.Overlay(transparency=0.5)
		from pywinauto_recorder.recorder import overlay_add_play_icon
		import traceback
		import codecs
		
		overlay_add_play_icon(main_overlay, 10, 10)
		if os_path_isfile(args.filename):
			with codecs.open(args.filename, "r", encoding='utf-8') as python_file:
				data = python_file.read()
			print("Replaying: " + args.filename)
			try:
				compiled_code = compile(data, '<string>', 'exec')
				exit_code = eval(compiled_code)
			except Exception as e:
				windll.user32.ShowWindow(windll.kernel32.GetConsoleWindow(), 3)
				exc_type, exc_value, exc_traceback = sys.exc_info()
				output = traceback.format_exception(exc_type, exc_value, exc_traceback)
				for line in output:
					if 'SyntaxError' in output[-1]:
						if not '.py' in line:
							if 'File "<string>", line ' in line:
								print(line.split(',')[1])
							else:
								print(line)
					else:
						print(line)
				input("Press Enter to continue...")
		else:
			print("Error: file '" + args.filename + "' not found.")
		main_overlay.clear_all()
		main_overlay.refresh()
		print("Exit")
	else:
		from pywinauto_recorder.recorder import *
		from pywinauto_recorder import __version__
		from win32api import GetSystemMetrics
		
		recorder = Recorder()
		import pystray
		from pystray import Icon as icon, Menu as menu, MenuItem as item
		from PIL import Image, ImageDraw
		from pathlib import Path

		p = str(Path("../Images/logo.png").absolute())
		
		def recording_text(icon):
			if recorder.get_mode() == 'Record':
				return 'Stop recording'
			else:
				return 'Start recording'

		def on_clicked_recording(icon, item):
			if recorder.get_mode() == 'Record':
				recorder.stop_recording()
			else:
				recorder.start_recording()

		def mode_text(icon):
			if recorder.get_mode() == 'Stop':
				return 'Start colouring'
			else:
				return 'Stop colouring'

		def on_clicked_mode(icon, item):
			if recorder.get_mode() == 'Record':
				recorder.stop_recording()
			elif recorder.get_mode() == 'Stop':
				recorder.start_colouring()
			else:
				recorder.stop_colouring()

		state_display_element_info = False
		state_smart_mode = False
		state_menu_select = False
		
		def on_clicked_display_element_info(icon, item):
			global state_display_element_info
			if recorder.is_displaying_info_tip():
				state_display_element_info = False
			else:
				state_display_element_info = True
			recorder.set_display_info_tip(state_display_element_info)
			
		def on_clicked_smart_mode(icon, item):
			global state_smart_mode
			if recorder.is_smart_mode():
				state_smart_mode = False
			else:
				state_smart_mode = True
			recorder.set_smart_mode(state_smart_mode)

		def on_clicked(icon, item):
			global state_menu_select
			state_menu_select = item.checked

		state_set_text = False
		
		def on_clicked2(icon, item):
			global state_set_text
			state_set_text = item.checked

		def on_clicked_open_explorer(icon, item):
			pywinauto_recorder_path = Path.home() / Path("Pywinauto recorder")
			os.system('explorer "' + str(pywinauto_recorder_path) + '"')

		def on_clicked_display_help(icon, item):
			display_splash_screen()

		def on_clicked_display_web_site(icon, item):
			os.system('rundll32 url.dll,FileProtocolHandler "https://pywinauto-recorder.readthedocs.io"')

		def on_clicked_exit(icon, item):
			recorder.quit()
			icon.stop()
			while recorder.is_alive():
				time.sleep(1.0)
			print("Exit")
			sys_exit(0)

		SEPARATOR = item('- - - -', None)
		menu = menu(
			item(recording_text, on_clicked_recording),
			item(mode_text, on_clicked_mode),
			item('Open output folder', on_clicked_open_explorer),
			item(SEPARATOR, None),
			item('Display element info', on_clicked_display_element_info, checked=lambda item: state_display_element_info),
			item('Smart mode', on_clicked_smart_mode, checked=lambda item: state_smart_mode),
			item('Apply treatments', menu(
				item('menu_select', on_clicked, checked=lambda item: state_menu_select),
				item('set_text', on_clicked2, checked=lambda item: state_set_text)
			)),
			item(SEPARATOR, None),
			item('Help', menu(
				item('About', on_clicked_display_help),
				item('Web site', on_clicked_display_web_site)
			)),
			item('Exit', on_clicked_exit),
		)
		image = Image.open(p)
		icon = pystray.Icon("Pywinauto recorder", image, "Pywinauto recorder", menu=menu)
		icon.run()
		
		while recorder.is_alive():
			time.sleep(1.0)
		print("Exit")

	sys_exit(0)
	
	