# -*- coding: utf-8 -*-


from __init__ import __version__

from recorder import *
import overlay_arrows_and_more as oaam
import time
import argparse
import os
import sys
import codecs
from win32api import GetSystemMetrics
from player import *
import ctypes


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
	splash_screen = oaam.Overlay(transparency=0.0)
	time.sleep(0.2)
	splash_screen2 = oaam.Overlay(transparency=0.1)
	splash_screen3 = oaam.Overlay(transparency=0.1)
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

	splash_screen2.clear_all()
	splash_screen3.clear_all()
	x, y, w, h = screen_width / 3, screen_height / 4, screen_width / 3, line_height * nb_band
	splash_screen2.add(
		geometry=oaam.Shape.triangle, thickness=0,
		xyrgb_array=((x, y, 200, 0, 0), (x+w, y, 0, 128, 0), (x+w, y+h, 0, 0, 155)))
	splash_screen3.add(
		geometry=oaam.Shape.triangle, thickness=0,
		xyrgb_array=((x+w+1, y+h-1, 0, 0, 155), (x, y+h-1, 22, 128, 66), (x+1, y-1, 200, 0, 0)))
	splash_screen2.refresh()
	splash_screen3.refresh()

	overlay_add_pywinauto_recorder_icon(splash_screen2, screen_width / 3 + screen_width / 6 - 100, screen_height /4 + 30)

	for n in range(10):
		splash_screen.clear_all()
		i = 0
		while i < nb_band:
			if i < 2:
				font_size = 44
			else:
				font_size = 20
			splash_screen.add(
				x=screen_width / 3, y=screen_height / 4 + i * line_height, width=screen_width / 3, height=line_height,
				text=text_lines[i], font_size=font_size, text_color=(254, 255, 255), color=(255, 255, 255),
				geometry=oaam.Shape.rectangle, thickness=0
			)
			i = i + 1
		if n % 3 == 0:
			overlay_add_record_icon(
				splash_screen, screen_width / 3 + screen_width / 18, screen_height / 4 + line_height * 8.6)
		if n % 3 == 1:
			overlay_add_pause_icon(
				splash_screen, screen_width / 3 + screen_width / 18, screen_height / 4 + line_height * 8.6)
		if n % 3 == 2:
			overlay_add_stop_icon(
				splash_screen, screen_width / 3 + screen_width / 18, screen_height / 4 + line_height * 8.6)
		overlay_add_progress_icon(
			splash_screen, n % 5, screen_width / 3 + screen_width / 18, screen_height / 4 + line_height * 10.8)
		overlay_add_play_icon(
			splash_screen, screen_width / 3 + screen_width / 18, int(screen_height / 4 + line_height * 17.3))
		splash_screen.refresh()
		time.sleep(0.4)
	splash_screen.clear_all()
	splash_screen.refresh()
	splash_screen2.clear_all()
	splash_screen2.refresh()
	splash_screen3.clear_all()
	splash_screen3.refresh()


if __name__ == '__main__':
	ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 6)
	parser = argparse.ArgumentParser()
	parser.add_argument(
		"filename", metavar='path', help="replay a python script", type=str, action='store', nargs='?', default='')
	parser.add_argument(
		"--no_splash_screen", help="Does not display the splash screen", action='store_true')
	args = parser.parse_args()
	if args.filename:
		main_overlay = oaam.Overlay(transparency=0.5)
		overlay_add_play_icon(main_overlay, 10, 10)
		if os.path.isfile(args.filename):
			with codecs.open(args.filename, "r", encoding='utf-8') as python_file:
				data = python_file.read()
			strCode = data.encode('utf-8').replace("from pywinauto_recorder import *", "")
			print("Replaying: " + args.filename)
			code = compile(strCode, '<string>', 'exec')
			exec(code)
		else:
			print("Error: file '" + args.filename + "' not found.")
		main_overlay.clear_all()
		main_overlay.refresh()
		print("Exit")
	else:
		if not args.no_splash_screen:
			display_splash_screen()
		recorder = Recorder()
		while recorder.is_alive():
			time.sleep(1.0)
		print("Exit")

	sys.exit(0)
