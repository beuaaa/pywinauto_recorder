# coding: utf-8
import time
import os
import shutil
from dp_toolbox import *
from pywinauto_recorder.player import *
from pathlib import Path
import pywinauto
import subprocess
import win32gui

def test():
	# Test performance of findwindows.find_elements vs descendants
	'''
	from pywinauto import findwindows
	from pywinauto.controls.uiawrapper import UIAWrapper
	import time
	
	t0 = time.time()
	for _ in range(9):
		eis = findwindows.find_elements(title_re="Zero", control_type="Button", backend="uia",
		                                        top_level_only=False)
		ei = eis[0]
		w = UIAWrapper(ei)
		#w.draw_outline()
	t1 = time.time()
	print("Duration: " + str(t1 - t0))
	return
	'''
	
	PlayerSettings.timeout = 2.5
	
	with UIPath(u"Calculator||Window"):
		wrapper = find(u"Calculator||Window->||Group->RegEx: Display is ?||Text")
		wrapper.draw_outline()
	
	print("\ntest(): ********************************************************")
	
	with UIPath(u"RegEx: .* - Google Chrome$||Pane", regex_title=True):
		wrapper = find(u"Google Chrome||Pane->*->||ToolBar->Address and search bar||Edit")
		wrapper.draw_outline(colour='blue', thickness=9)

if __name__ == '__main__':

	os.chdir(os.path.dirname(os.path.abspath(__file__)))

	brian = Voice(name='Brian')
	#obs = OBSStudio(r'D:\workbench-2020\Software\AutomatedValidation\SE-VALIDATION\OBSStudio\OBS-Studio-22.0.2-Full-x64\bin\64bit\obs64.exe')
	#obs.start_recording()
	

	pywinauto.Application().start(r'C:\Program Files\Notepad++\notepad++.exe')
	with UIPath(r".* - Notepad\+\+$||Window", regex_title=True):
		#find().set_focus()
		find().maximize()
		menu_click("", u"File->New\tCtrl+N", menu_type="NPP")

	
	
	exit(0)
	####################################
	# Open 'Pywinauto Recorder' folder #
	####################################
	brian.say("The generated file is in 'Pywinauto Recorder' folder under your home folder.")
	move((0, 0), duration=0)
	PlayerSettings.mouse_move_duration = 1
	pwr = subprocess.Popen(r'C:\Users\d_pra\AppData\Local\Programs\Pywinauto recorder\pywinauto_recorder.exe')
	time.sleep(1.5)
	
	
	
	calc = subprocess.Popen('calc.exe')
	time.sleep(1.5)
	hwnd = win32gui.FindWindow(None, 'Calculator')
	win32gui.MoveWindow(hwnd, 1226, 150, 700, 800, True)
	
	time.sleep(1)
	
	
	for i in range(2):
		with UIPath("Calculator||Window"):
			move("*->One||Button")
			time.sleep(1)
			move("*->Three||Button")
			time.sleep(1)
			move("*->Nine||Button")
			time.sleep(1)
			move("*->Seven||Button")
			time.sleep(1)
	
	with UIPath("Calculator||Window"):
		left_click("*->Close Calculator||Button")
	pwr.kill()
	
	brian.say("Thanks for watching! In the next tutorial, you will see how to make a robust recorded script. See you soon!", wait_until_the_end_of_the_sentence=True)

	#obs.stop_recordi0ng()
	#time.sleep(1)
	#obs.quit()
	


	exit(0)