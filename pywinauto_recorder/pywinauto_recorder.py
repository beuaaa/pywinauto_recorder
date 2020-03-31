# -*- coding: utf-8 -*-


from __init__ import __version__

from recorder import *
import overlay_arrows_and_more as oaam
import time
import argparse
import os
import codecs
from win32api import GetSystemMetrics
from player import *


def display_splash_screen():
	splash_screen = oaam.Overlay(transparency=0.0)
	screen_width = GetSystemMetrics(0)
	screen_height = GetSystemMetrics(1)
	nb_band = 15
	line_height = (screen_height / 2) / nb_band
	text_lines = [''] * nb_band
	text_lines[1] = 'Pywinauto recorder'
	text_lines[2] = 'version ' + __version__
	text_lines[4] = 'ALT+R : Recording/Pause'
	text_lines[6] = 'Search algorithm speed'
	text_lines[8] = 'ALT+Q : Quit'
	text_lines[12] = 'Drag & drop a recorded file on pywinauto_recorder.exe'
	text_lines[13] = ' to play it back'
	for n in xrange(5):
		splash_screen.clear_all()
		color = [60, 70, 90]
		i = 0
		while i < nb_band:
			splash_screen.add(
				x=screen_width / 3, y=screen_height / 4 + i * line_height, width=screen_width / 3,
				height=line_height, text=text_lines[i], text_color=(0, 0, 0), color=tuple(color),
				geometry=oaam.Shape.rectangle, thickness=1, brush=oaam.Brush.solid, brush_color=tuple(color))
			color[0], color[1], color[2] = color[0] + 2, color[1] + 1, color[2] + 3
			i = i + 1
		if n % 2 == 0:
			overlay_add_record_icon(
				splash_screen, screen_width / 3 + screen_width / 12, screen_height / 4 + line_height * 4)
		else:
			overlay_add_pause_icon(
				splash_screen, screen_width / 3 + screen_width / 12, screen_height / 4 + line_height * 4)
		overlay_add_progress_icon(
			splash_screen, n % 5, screen_width / 3 + screen_width / 12, screen_height / 4 + line_height * 6)
		overlay_add_play_icon(
			splash_screen, screen_width / 2 - 20, screen_height / 4 + line_height * 11 - 10)
		splash_screen.refresh()
		time.sleep(0.8)
	splash_screen.clear_all()
	splash_screen.refresh()


if __name__ == '__main__':
	parser = argparse.ArgumentParser()
	parser.add_argument(
		"filename", metavar='path', help="replay a python script", type=str, action='store', nargs='?', default='')
	args = parser.parse_args()
	if args.filename:
		main_overlay = oaam.Overlay(transparency=0.5)
		overlay_add_play_icon(main_overlay, 10, 10)
		if os.path.isfile(args.filename):
			with codecs.open(args.filename, "r", encoding='utf-8') as python_file:
				data = python_file.read()
			strCode = data.encode('utf-8')
			print("Replaying: " + args.filename)
			code = compile(strCode, '<string>', 'exec')
			exec code
		else:
			print("Error: file '" + args.filename + "' not found.")
		main_overlay.clear_all()
		main_overlay.refresh()
		print("Exit")
	else:
		display_splash_screen()
		recorder = Recorder()
		while recorder.is_alive():
			time.sleep(1.0)
		print("Exit")

	exit(0)
