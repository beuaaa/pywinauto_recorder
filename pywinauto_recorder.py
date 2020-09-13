#!/usr/bin/env python3

# print('__file__={0:<35} | __name__={1:<20} | __package__={2:<20}'.format(__file__, __name__, str(__package__)))

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
		geometry=oaam.Shape.rectangle, x=x+5, y=y+15, width=190, height=38, thickness=0,
		brush=oaam.Brush.solid, brush_color=(255, 254, 255), text_color=(66, 66, 66),
		text=u'PYWINAUTO', font_size=44, font_name='Impact')
	overlay.add(
		geometry=oaam.Shape.rectangle, x=x+20, y=y+50, width=160, height=38, thickness=0,
		brush=oaam.Brush.solid, brush_color=(255, 254, 255), text_color=(200, 40, 40),
		text=u'recorder', font_size=48, font_name='Arial Black')


def display_splash_screen():
	splash_foreground = oaam.Overlay(transparency=0.0)
	time.sleep(0.2)
	splash_background = oaam.Overlay(transparency=0.1)
	screen_width = GetSystemMetrics(0)
	screen_height = GetSystemMetrics(1)
	nb_band = 20
	line_height = (screen_height / 2) / nb_band
	text_lines = [''] * nb_band
	text_lines[6] = 'Pywinauto recorder ' + __version__
	text_lines[7] = 'by David Pratmarty'
	text_lines[9] = 'CTRL+ALT+R : Pause / Record / Stop'
	text_lines[11] = 'Search algorithm speed'
	text_lines[13] = 'CTRL+SHIFT+F : Copy element path in clipboard'
	text_lines[15] = 'CTRL+ALT+Q : Quit'
	text_lines[17] = 'Drag and drop a recorded file on '
	text_lines[18] = 'pywinauto_recorder.exe to replay it'

	splash_background.clear_all()
	x, y, w, h = screen_width / 3, screen_height / 4, screen_width / 3, line_height * nb_band
	splash_background.add(
		geometry=oaam.Shape.triangle, thickness=0,
		xyrgb_array=((x, y, 200, 0, 0), (x+w, y, 0, 128, 0), (x+w, y+h, 0, 0, 155)))
	splash_background.add(
		geometry=oaam.Shape.triangle, thickness=0,
		xyrgb_array=((x + w, y + h, 0, 0, 155), (x, y + h, 22, 128, 66), (x, y, 200, 0, 0)))

	splash_background.refresh()

	py_rec_icon_rect = (screen_width / 3 + screen_width / 6 - 100, screen_height /4 + 30, 200, 100)
	overlay_add_pywinauto_recorder_icon(splash_background, py_rec_icon_rect[0], py_rec_icon_rect[1])
	from pywinauto_recorder.player import move as mouse_move
	mouse_move((py_rec_icon_rect[0]+py_rec_icon_rect[2]/2, py_rec_icon_rect[1]+py_rec_icon_rect[3]/2))
	mouse_cursor_on_py_rec_icon = True
	n = 0
	while mouse_cursor_on_py_rec_icon:
	#for n in range(10):
		splash_foreground.clear_all()
		i = 0
		while i < nb_band:
			if i < 2:
				font_size = 44
			else:
				font_size = 20
			splash_foreground.add(
				x=screen_width / 3, y=screen_height / 4 + i * line_height, width=screen_width / 3, height=line_height,
				text=text_lines[i], font_size=font_size, text_color=(254, 255, 255), color=(255, 255, 255),
				geometry=oaam.Shape.rectangle, thickness=0
			)
			i = i + 1
		if n % 3 == 0:
			overlay_add_record_icon(
				splash_foreground, screen_width / 3 + screen_width / 18, screen_height / 4 + line_height * 8.6)
		if n % 3 == 1:
			overlay_add_pause_icon(
				splash_foreground, screen_width / 3 + screen_width / 18, screen_height / 4 + line_height * 8.6)
		if n % 3 == 2:
			overlay_add_stop_icon(
				splash_foreground, screen_width / 3 + screen_width / 18, screen_height / 4 + line_height * 8.6)
		overlay_add_progress_icon(
			splash_foreground, n % 5, screen_width / 3 + screen_width / 18, screen_height / 4 + line_height * 10.8)
		overlay_add_play_icon(
			splash_foreground, screen_width / 3 + screen_width / 18, int(screen_height / 4 + line_height * 17.3))
		splash_foreground.refresh()
		time.sleep(0.4)
		x, y = win32api.GetCursorPos()
		mouse_cursor_on_py_rec_icon = False
		if py_rec_icon_rect[0]< x < py_rec_icon_rect[0]+py_rec_icon_rect[2]:
			if py_rec_icon_rect[1] < y < py_rec_icon_rect[1] + py_rec_icon_rect[3]:
				mouse_cursor_on_py_rec_icon = True
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
		if not args.no_splash_screen:
			from pywinauto_recorder import __version__
			from win32api import GetSystemMetrics
			display_splash_screen()
		recorder = Recorder()
		while recorder.is_alive():
			time.sleep(1.0)
		print("Exit")

	sys_exit(0)
