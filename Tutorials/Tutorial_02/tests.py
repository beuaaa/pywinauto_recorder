# coding: utf-8
import time
import os
import shutil
from dp_toolbox import *
from pywinauto_recorder.player import *
from pathlib import Path


if __name__ == '__main__':
	os.chdir(os.path.dirname(os.path.abspath(__file__)))

	brian = Voice(name='Brian')
	obs = OBSStudio()
	
	#####################################################################
	# Copy script in Data to 'Pywinauto Recorder' folder in Home folder #
	#####################################################################
	shutil.copy(Path("Data")/Path("recorded Mon Apr  5 16_28_28 2021.py"), Path.home()/Path("Pywinauto recorder"))
	
	
	####################################
	# Open 'Pywinauto Recorder' folder #
	####################################
	brian.say("The generated file is in 'Pywinauto Recorder' folder under your home folder.")
	time.sleep(1.5)
	with Window(u"||List"):
		double_left_click("Home||ListItem")
		time.sleep(1.5)
		send_keys("{LWIN down}""{VK_RIGHT}""{LWIN up}")
	with Window("C:\\Users\\d_pra||Window"):
		double_left_click("*->Shell Folder View||Pane->Items View||List->Pywinauto recorder||ListItem->Name||Edit")
		
	#############################################################
	# Drag and drop the recorded file on pywinauto_recorder.exe #
	#############################################################
	with Window("C:\\\\Users\\\\.*\\\\Pywinauto recorder||Window", regex_title=True):
		drag_and_drop_start = find("*->Shell Folder View||Pane->Items View||List->||ListItem->Name||Edit#[0,0]")
		drag_and_drop_start.draw_outline(colour='blue')
	with Window("||List"):
		drag_and_drop_end = find(u"Pywinauto recorder||ListItem")
		drag_and_drop_end.draw_outline()
	#with Window(u"Program Manager||Pane"):
	#	drag_and_drop_end = find(u"Desktop||List->Pywinauto recorder||ListItem")
	#	drag_and_drop_end.draw_outline()
	
	time.sleep(1)
	drag_and_drop(drag_and_drop_start, drag_and_drop_end)
	brian.say("Thanks for watching! In the next tutorial, you will see how to make a robust recorded script. See you soon!", wait_until_the_end_of_the_sentence=True)

	obs.stop_recording()
	time.sleep(1)
	obs.quit()


	exit(0)

