# coding: utf-8
import time
import os
import shutil
from dp_toolbox import Voice, OBSStudio, highlight, clear_overlay
import pywinauto
from pywinauto_recorder.player import UIPath, click, right_click, move, find, send_keys, exists, drag_and_drop
from pathlib import Path

# WARNING: This script must be launched with admin rights because of the installer


if __name__ == '__main__':
	os.chdir(os.path.dirname(os.path.abspath(__file__)))

	brian = Voice(name='Brian')
	obs = OBSStudio()

	###############################################
	# Remove Pywinauto Recorder icon from desktop #
	###############################################
	pr_ico_path = Path.home() / Path("Desktop") / Path("Pywinauto recorder")
	if pr_ico_path.is_file():
		pr_ico_path.unlink()
	
	###############################################################################
	# Remove Downloads\pywinauto_recorder.dist.zip and R:\pywinauto_recorder.dist #
	###############################################################################
	pywinauto_recorder_path = Path.home() / Path("Pywinauto recorder")
	downloads_path = Path.home() / Path("Downloads")
	pywinauto_recorder_installer = downloads_path / Path("Pywinauto_recorder_installer.exe")
	if pywinauto_recorder_installer.is_file():
		pywinauto_recorder_installer.unlink()
	if pywinauto_recorder_path.is_dir():
		shutil.rmtree(pywinauto_recorder_path, ignore_errors=True)

	# time.sleep(2.5)

	############################################################
	# Start google chrome 'GitHub - beuaaa/pywinauto_recorder' #
	############################################################
	chrome_dir = r'"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"'
	chrome = pywinauto.Application(backend='uia')
	chrome.start(chrome_dir + ' --force-renderer-accessibility --start-maximized --guest ')
	move((66, 66))
	time.sleep(0.5)
	obs.start_recording()
	#screen_zoom_in()
	time.sleep(0.5)
	#screen_zoom_in()
	time.sleep(1.5)
	#screen_zoom_in()
	brian.say("Hi, let me introduce you to 'Pywinauto recorder'.")
	click("New Tab - Google Chrome (Guest)||Pane->*->Address and search bar||Edit%(-99,-23.33)")
	
	brian.say("Go to pywinauto-recorder dot read the docs dot io")
	send_keys("pywinauto-recorder.readthedocs.io{ENTER}", pause=0.15, vk_packet=True)
	#screen_zoom_out()


	brian.say("""
	"Pywinauto recorder" is a great tool! It can record user actions and saves them in a Python script. 
	Then the saved Python script can be run to replay the user actions previously recorded.
	""")

	########################################################
	# Wait page to be ready and click on the download link #
	########################################################
	pywinauto_recorder_documentation_page = "Description — pywinauto_recorder documentation - Google Chrome (Guest)||Pane"
	with UIPath(pywinauto_recorder_documentation_page):
		button_wrapper = find("*->||Group->Download installer||Hyperlink", timeout=120)
		click(
			"*->” is a standalone application, it’s the compiled version of “pywinauto_recorder.py” for 64-bit Windows.||Text%(-70,-80)",
			timeout=120)
	brian.say("""
	"Pywinauto recorder" is a unique "record-replay" tool in the open source field because it generates reliable scripts without hard-coded coordinates thanks to Pywinauto.
	Pywinauto is a library that uses accessibility technologies allowing you to automate almost any type of GUI.
	The functions of the generated Python script return Pywinauto wrappers so it can be enhanced with Pywinauto methods.
	""")
	brian.say("Let's get started! Download Pywinauto recorder installer for 64-bit Windows on this web page")
	send_keys("{VK_DOWN 3}")
	time.sleep(1)
	highlight(button_wrapper, color=(0, 0, 255))
	move(button_wrapper)
	brian.say("Click on this download button.", wait_until_the_end_of_the_sentence=True)
	clear_overlay()

	##################################################
	# Wait until the exe is downloaded and launch it #
	##################################################
	brian.say("Run Pywinauto recorder installer.")
	with UIPath(pywinauto_recorder_documentation_page):
		move("*->Options menu||Button")
		click("*->Options menu||Button")
		click("*->||Menu->Open||MenuItem")
	time.sleep(1.0)
	with UIPath("*->Windows Defender SmartScreen||Window"):
		click("*->More info||Hyperlink")
		click("*->Run anyway||Button")
	
	########################################################################
	#                 Do Pywinauto recorder installer steps                #
	########################################################################
	time.sleep(1.5)
	with UIPath("Pywinauto recorder Setup ||Window"):
		click("Install||Button")
		click("Show details||Button")
		wrapper_progress_bar = find("||ProgressBar")
		while wrapper_progress_bar.legacy_properties()['Value'] != '90%':
			time.sleep(1)
		click("*->Yes||Button")
		time.sleep(1)
	
	########################################################################
	#          Run pywinauto_recorder.exe from the installer               #
	########################################################################
	config_ini = Path.home() / r"Pywinauto recorder/config.ini"
	config_ini.parent.mkdir(exist_ok=True, parents=True)
	config_ini.write_text("""[window_filtering]
		mode = ignore_windows
		admit_windows = []
		ignore_windows = ['Google Chrome', 'Program Manager']""")
	send_keys("{ENTER}")  # click("*->Yes||Button")
	speech = """
		When you run Pywinauto Recorder it starts with this window.
		It contains useful information.
		All actions can be triggered by a keyboard shortcut or by clicking in the tray menu.
		This window is closed my moving the mouse outside.
	"""
	brian.say(speech, wait_until_the_end_of_the_sentence=True)
	move((66, 66))
	speech = """
	Now It is running in inspection mode.
	In this mode represented by the magnifying glass icon displayed in the top left corner of the screen,
	Pywinauto Recorder colors green the UI element it can identify under the mouse cursor.
		"""
	brian.say(speech)
	# admin
	with UIPath(u"Pywinauto recorder Setup ||Window"):
		move("||TitleBar")
		time.sleep(1)
		brian.say("The mouse cursor can be moved around the installer and Pywinauto recorder shows UI element information in real time based on what UI element the cursor is hovering over.")
		move("||TitleBar->RegEx: Minimi(s|z)?e||Button")
		time.sleep(1)
		move("||TitleBar->RegEx: Maximi(s|z)?e||Button")
		time.sleep(1)
		move("||TitleBar->Close||Button")
		time.sleep(1)
		wrapper = find("*->Completed||List").get_item(-2)
		move(wrapper)
		time.sleep(1)
		move("||ProgressBar") # get value 100%
		time.sleep(1)
		brian.say("For instance this progress bar is interesting because Pywinauto recorder informs you that you can use a Pywinauto method to retrieve the progress percentage.", wait_until_the_end_of_the_sentence=True)
		move("< Back||Button")
		time.sleep(1)
		move("Close||Button")
		time.sleep(1)
		move("Cancel||Button")
		time.sleep(1)
		click("Close||Button")
	move((66, 66))
	brian.say("To start recording my next actions I press 'alt' 'control' 'r'.", wait_until_the_end_of_the_sentence=True)
	time.sleep(1.0)

	######################################
	# Start Recording and run Calculator #
	######################################
	send_keys("{VK_CONTROL down}{VK_MENU down}r{VK_MENU up}{VK_CONTROL up}")
	brian.say("Now it's recording!")
	time.sleep(2.0)
	send_keys("{VK_LWIN}")  # Task bar must be configured to hide and show when mouse hover it.
	time.sleep(1.0)
	brian.say("Then I'm going to start calculator to do a little addition.")
	move("Start||Window->||Window->Type here to search||Button")
	send_keys("calculator{ENTER}", pause=0.2, vk_packet=False)
	time.sleep(2.0)
	
	################################
	# Move Calculator to the right #
	################################
	import win32gui
	hwnd = win32gui.FindWindow(None, 'Calculator')
	win32gui.MoveWindow(hwnd, 1100, 150, 800, 800, True)

	###################################
	# Press 1+2= and close Calculator #
	###################################
	with UIPath("Calculator||Window->*"):
		move("One||Button")
		time.sleep(2.0)
		brian.say("One")
		click("One||Button")

		move("Plus||Button")
		time.sleep(1.0)
		brian.say("Plus")
		click("Plus||Button")

		move("Two||Button")
		time.sleep(1.0)
		brian.say("Two")
		click("Two||Button")

		move("Equals||Button")
		time.sleep(1.0)
		brian.say("Equals Three")
		click("Equals||Button")

		move("Close Calculator||Button")
		brian.say("Finally, I close the calculator")
		time.sleep(1.5)
		click("Close Calculator||Button")

	#######################################
	# Stop Recording and replay clipboard #
	#######################################
	brian.say("I press 'alt' 'control' 'r' to stop recording.")
	send_keys("{VK_CONTROL down}{VK_MENU down}r{VK_MENU up}{VK_CONTROL up}")
	brian.say("And click on Start playing in the tray menu.")
	time.sleep(1.5)
	send_keys("{VK_LWIN}")
	time.sleep(1.5)
	click("Taskbar||Pane->*->Notification Chevron||Button")
	time.sleep(0.5)
	right_click("Notification Overflow||Pane->*->Pywinauto recorder||Button")
	click("Context||Menu->Start replaying clipboard||MenuItem")
	time.sleep(1)
	brian.say("Now it's replaying.")
	while exists("Calculator||Window"):
		time.sleep(1)

	####################################
	# Open 'Pywinauto Recorder' folder #
	####################################
	brian.say("The generated file is in 'Pywinauto Recorder' folder under your home folder.")
	time.sleep(1.5)
	send_keys("{VK_LWIN}")
	time.sleep(1.5)
	click("Taskbar||Pane->*->Notification Chevron||Button")
	time.sleep(0.5)
	right_click("Notification Overflow||Pane->*->Pywinauto recorder||Button")
	click("Context||Menu->Open output folder||MenuItem")
	time.sleep(0.5)
	with UIPath("RegEx: ^C:\\\\Users\\\\.*\\\\Pywinauto recorder$||Window"):
		click("||TitleBar")
		send_keys("{LWIN down}{VK_RIGHT}{LWIN up}")
		time.sleep(0.5)
	with UIPath(u"Snap Assist||Window"):
		click("*->Description — pywinauto_recorder documentation - Google Chrome||ListItem->Close||Button")
		click((10, 10))
	
	############################
	# Locate the recorded file #
	############################
	with UIPath("RegEx: ^C:\\\\Users\\\\.*\\\\Pywinauto recorder$||Window"):
		drag_and_drop_start = click("*->Shell Folder View||Pane->*->RegEx: recorded.*py||ListItem->Name||Edit")
		drag_and_drop_start.draw_outline(colour='red')

	#########################################
	# Open the recorded file with Notepad++ #
	#########################################
	brian.say("Let's open the file that was just created to see what it looks like.")
	right_click(drag_and_drop_start)
	click("Context||Menu->*->Edit with Notepad++||MenuItem%(-64,28)")
	find("RegEx: .* - Notepad\+\+||Window").maximize()
	brian.say("You can see that it is an easy to read Python script.")

	send_keys("{PGUP}{VK_HOME}")

	brian.say("The functions return Pywinauto wrappers.")

	send_keys("{VK_DOWN 9}")
	send_keys("{VK_RIGHT 2}")
	send_keys("pywinauto_wrapper1 = ")
	brian.say("I'm going to call draw outline on some of these wrappers")
	send_keys("{VK_END}{ENTER}")
	brian.say("The first draw outline will display a red rectangle around the One button.")
	send_keys("pywinauto_wrapper1.draw_outline(colour='red', thickness=22)")
	send_keys("{VK_HOME}{VK_DOWN}")
	send_keys("pywinauto_wrapper2 = ")
	send_keys("{VK_END}{ENTER}")
	brian.say("The second one will display a green rectangle around the Plus button.")
	send_keys("pywinauto_wrapper2.draw_outline(colour='green', thickness=22)")
	send_keys("{VK_HOME}{VK_DOWN}")
	send_keys("pywinauto_wrapper3 = ")
	send_keys("{VK_END}{ENTER}")
	brian.say("And the third one will display a blue rectangle around the Two button.")
	send_keys("pywinauto_wrapper3.draw_outline(colour='blue', thickness=22)")
	send_keys("{VK_CONTROL down}s{VK_CONTROL up}")
	brian.say("Let's replay it!", wait_until_the_end_of_the_sentence=True)
	send_keys("{VK_MENU down}{VK_F4}{VK_MENU up}")
	
	################################
	# Close pywinauto_recorder.exe #
	################################
	time.sleep(1.5)
	send_keys("{VK_LWIN}")
	time.sleep(1.5)
	click("Taskbar||Pane->*->Notification Chevron||Button")
	right_click("Notification Overflow||Pane->*->Pywinauto recorder||Button")
	time.sleep(0.5)
	right_click()
	click("Context||Menu->Quit||MenuItem")

	#############################################################
	# Drag and drop the recorded file on pywinauto_recorder.exe #
	#############################################################
	with UIPath("RegEx: ^C:\\\\Users\\\\.*\\\\Pywinauto recorder$||Window"):
		drag_and_drop_start = find("*->Shell Folder View||Pane->*->RegEx: recorded.*py||ListItem->Name||Edit")
		drag_and_drop_start.draw_outline(colour='blue')
	with UIPath("||List"):
		drag_and_drop_end = find("Pywinauto recorder||ListItem")
		drag_and_drop_end.draw_outline()
	time.sleep(1)
	drag_and_drop(drag_and_drop_start, drag_and_drop_end)
	brian.say("""
		Now Pywinauto recorder is replaying. You can see the Replay icon in the top left corner of the screen.
		The actions we recorded in the Python script have been replayed and the three draw outline functions we've just added have been called.
		""")
	time.sleep(15)
	brian.say("Thanks for watching! In the next tutorial, you will see how to make a robust recorded script. See you soon!", wait_until_the_end_of_the_sentence=True)

	obs.stop_recording()
	obs.quit()
	exit(0)

