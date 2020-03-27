# -*- coding: utf-8 -*-

from recorder import Recorder
import overlay_arrows_and_more as oaam
import time
import argparse
import os
import codecs
from player import *

if __name__ == '__main__':
	global recorder
	parser = argparse.ArgumentParser()
	parser.add_argument("filename", help="replay a python script", type=str)
	args = parser.parse_args()
	if args.filename:
		main_overlay = oaam.Overlay(transparency=0.5)
		main_overlay.add(
			geometry=oaam.Shape.rectangle, x=10, y=10, width=40, height=40,
			color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(255, 255, 254))
		main_overlay.add(
			geometry=oaam.Shape.polyline, xy_array=((15, 15), (15, 45), (45, 30)),
			color=(0, 0, 0), thickness=1, brush=oaam.Brush.solid, brush_color=(0, 255, 0))

		main_overlay.refresh()
		time.sleep(0.5)
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
		recorder = Recorder()
		while recorder.is_alive():
			time.sleep(1.0)
		print("Exit")

	exit(0)
