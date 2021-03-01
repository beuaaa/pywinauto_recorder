# coding: utf-8
import time
import os
import shutil
from dp_toolbox import *
from pywinauto_recorder.player import *
from pathlib import Path


# send_keys("{VK_CONTROL down}""{VK_MENU down}""q""{VK_MENU up}""{VK_CONTROL up}")  # works
# send_keys("{VK_MENU down}""{VK_CONTROL down}""q""{VK_CONTROL up}""{VK_MENU up}") # does not work

if __name__ == '__main__':
	brian = Voice(name='Brian')

	obs = OBSStudio(
		r'D:\workbench-2020\Software\AutomatedValidation\SE-VALIDATION\OBSStudio\OBS-Studio-22.0.2-Full-x64\bin\64bit\obs64.exe')

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

	time.sleep(2.5)

	############################################################
	# Start google chrome 'GitHub - beuaaa/pywinauto_recorder' #
	############################################################
	chrome_dir = r'"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"'
	chrome = pywinauto.Application(backend='uia')
	chrome.start(chrome_dir + ' --force-renderer-accessibility --start-maximized --guest ')
	move((0, 0))
	time.sleep(1.5)
	send_keys("{VK_LWIN down}""{VK_UP}""{VK_LWIN up}")
	time.sleep(0.5)
	obs.start_recording()
	screen_zoom_in()
	time.sleep(0.5)
	screen_zoom_in()
	time.sleep(0.5)
	screen_zoom_in()
	brian.say("Hi, let me introduce you to 'Pywinauto recorder'.")
	with Window(u"New Tab - Google Chrome (Guest)||Pane"):
		left_click(u"Google Chrome||Pane->||Pane->||Pane->||Pane->Address and search bar||Edit%(-38.88,-23.33)")

	brian.say("Go to pywinauto-recorder dot read the docs dot io")
	send_keys(u"pywinauto-recorder.readthedocs.io""{ENTER}", pause=0.15, vk_packet=True)
	screen_zoom_out()


	brian.say("""
	"Pywinauto recorder" is a great tool! It can record user actions and saves them in a Python script. 
	Then the saved Python script can be run to replay the user actions previously recorded.
	""")

	########################################################
	# Wait page to be ready and click on the download link #
	########################################################
	pywinauto_recorder_documentation_page = u"Description — Python documentation - Google Chrome (Guest)||Pane"
	with Window(pywinauto_recorder_documentation_page):
		button_wrapper = None
		while not button_wrapper:
			try:
				button_wrapper = find(
					"*->https://raw.githubusercontent.com/beuaaa/pywinauto_recorder/master/Images/Download.png?sanitize=true||Hyperlink#[0,0]")
			except IndexError:
				pass
		left_click(
			"*->” is a standalone application, it’s the compiled version of “pywinauto_recorder.py” for 64-bit Windows.||Text%(-70,-80)",
			timeout=120)

	brian.say("""
	"Pywinauto recorder" is a unique "record-replay" tool in the open source field because it generates reliable scripts without hard-coded coordinates thanks to Pywinauto.
	Pywinauto is a library that uses accessibility technologies allowing you to automate almost any type of GUI.
	The functions of the generated Python script return Pywinauto wrappers so it can be enhanced with Pywinauto functions.
	""")
	brian.say("Let's get started! Download Pywinauto recorder for Windows on this web page")
	#send_keys("{PGDN}")
	time.sleep(1)
	highlight(button_wrapper, color=(0, 0, 255))
	move(button_wrapper)
	brian.say("Click on this download button.", wait_until_the_end_of_the_sentence=True)
	clear_overlay()

	##################################################
	# Wait until the exe is downloaded and launch it #
	##################################################
	brian.say("Run Pywinauto recorder installer.")
	with Window(pywinauto_recorder_documentation_page):
		move("Google Chrome||Pane->||Pane->Options menu||Button")
		left_click("Google Chrome||Pane->||Pane->Options menu||Button")
		left_click("||Pane->||MenuBar->||Pane->||Menu->Open||MenuItem")
	time.sleep(1.0)
	with Window("||Window->Windows Defender SmartScreen||Window"):
		left_click("about:blank||Pane->||Pane->Windows protected your PC||Group->More info||Hyperlink")
		left_click("about:blank#||Pane->||Pane->Run anyway||Button")

	########################################################################
	# Do Pywinauto recorder installer steps and Run pywinauto_recorder.exe #
	########################################################################
	time.sleep(1.5)
	with Window(u"Pywinauto recorder Setup ||Window"):
		left_click("Install||Button")
		left_click("Pywinauto recorder Setup||Window->OK||Button")
		time.sleep(1)
		send_keys("{ENTER}")
	speech = """
    When you run Pywinauto Recorder it starts with the help dialog.
    All actions can be triggered by a keyboard shortcut or by clicking in the tray menu.
    The help dialog is closed my moving the mouse outside.
	"""
	brian.say(speech, wait_until_the_end_of_the_sentence=True)
	move((0,0))
	speech = """
	It is runnning displaying element information.
	In this mode represented by the magnifying glass icon displayed in the top left corner of the screen,
	Pywinauto Recorder colors green the interface elements it can identify under the mouse cursor.
		"""
	brian.say(speech)
	# admin
	with Window(u"Pywinauto recorder Setup ||Window"):
		move("||TitleBar")
		time.sleep(1)
		move("||TitleBar->Minimize||Button")
		time.sleep(1)
		move("||TitleBar->Maximize||Button")
		time.sleep(1)
		move("||TitleBar->Close||Button")
		time.sleep(1)
		move("Installation Complete||Text")
		time.sleep(1)
		move("Setup was completed successfully.||Text")
		time.sleep(1)
		move("Setup was completed successfully.||Image")
		time.sleep(1)
		move("||ProgressBar") # get value 100%
		time.sleep(1)
		left_click("Show details||Button") #click
		move("Completed||List->Completed||ListItem->Completed||Text", duration=9)
		move("< Back||Button")
		time.sleep(1)
		move("Close||Button")
		time.sleep(1)
		move("Cancel||Button")
		time.sleep(1)
		left_click("Close||Button")
		
	brian.say("To start recording my next actions I press 'alt' 'control' 'r'.", wait_until_the_end_of_the_sentence=True)
	time.sleep(1.0)

	######################################
	# Start Recording and run Calculator #
	######################################
	send_keys("{VK_CONTROL down}""{VK_MENU down}""r""{VK_MENU up}""{VK_CONTROL up}")
	brian.say("Now it's recording!")
	time.sleep(2.0)
	send_keys("{VK_LWIN}")
	time.sleep(1.0)
	brian.say("Then I'm going to start calculator to do a little addition.")
	with Window(u"Start||Window"):
		move(u"Search||Window->||Edit%(-14.65,-24.00)")
		time.sleep(1.0)
		left_click(u"Search||Window->||Edit")
		send_keys("calculator""{ENTER}", pause=0.2, vk_packet=False)
	time.sleep(2.0)

	###################################
	# Press 1+2= and close Calculator #
	###################################
	with Window(u"Calculator||Window"):
		with Region(u"Calculator||Window->||Group"):
			move(u"Number pad||Group->One||Button%(13.45,21.92)")
			time.sleep(2.0)
			brian.say("One")
			left_click(u"Number pad||Group->One||Button")

			move(u"Standard operators||Group->Plus||Button%(-15.79,-17.81)")
			time.sleep(1.0)
			brian.say("Plus")
			left_click(u"Standard operators||Group->Plus||Button")

			move(u"Number pad||Group->Two||Button%(16.37,0.00)")
			time.sleep(1.0)
			brian.say("Two")
			left_click(u"Number pad||Group->Two||Button")

			move(u"Standard operators||Group->Equals||Button%(-20.47,-8.22)")
			time.sleep(1.0)
			brian.say("Equals Three")
			left_click(u"Standard operators||Group->Equals||Button")

		with Region(u"Calculator||Window"):
			move(u"Close Calculator||Button%(0.00,17.50)")
			brian.say("Finally, I close the calculator")
			time.sleep(1.5)
			left_click(u"Close Calculator||Button")

	#######################################
	# Stop Recording and replay clipboard #
	#######################################
	brian.say("I press 'alt' 'control' 'r' to stop recording.")
	send_keys("{VK_CONTROL down}""{VK_MENU down}""r""{VK_MENU up}""{VK_CONTROL up}")
	brian.say("And click on Start playing in the tray menu.")
	time.sleep(1.5)
	send_keys("{VK_LWIN}")
	with Window(u"Taskbar||Pane"):
		with Region(u"||Pane"):
			left_click(u"Notification Chevron||Button")
	with Window(u"Notification Overflow||Pane"):
		with Region(u"Overflow Notification Area||ToolBar"):
			right_click(u"Pywinauto recorder||Button")
	with Window(u"Context||Menu"):
		left_click(u"Start replaying clipboard||MenuItem")

	####################################
	# Open 'Pywinauto Recorder' folder #
	####################################
	brian.say("The generated file is in 'Pywinauto Recorder' folder under your home folder.")
	send_keys("{VK_LWIN}")
	with Window(u"Taskbar||Pane"):
		with Region(u"||Pane"):
			left_click(u"Notification Chevron||Button")
	with Window(u"Notification Overflow||Pane"):
		with Region(u"Overflow Notification Area||ToolBar"):
			right_click(u"Pywinauto recorder||Button")
	with Window(u"Context||Menu"):
		left_click(u"Open output folder||MenuItem")
	send_keys("{LWIN down}""{VK_RIGHT}""{LWIN up}")
	with Window(u"Snap Assist||Window"):
		left_click("Dismiss Task Switching Window||Button%(95,-98)")
	
	############################
	# Locate the recorded file #
	############################
	with Window(u"C:\\\\Users\\\\.*\\\\Pywinauto recorder||Window", regex_title=True):
		wrapper = find(
			u"*->||Pane->Shell Folder View||Pane->Items View||List->||ListItem->Name||Edit#[0,0]")
		wrapper.draw_outline(colour='red')
		drag_and_drop_start = get_wrapper_path(wrapper)

	#########################################
	# Open the recorded file with Notepad++ #
	#########################################
	brian.say("Let's open the file that was just created to see what it looks like.")
	right_click(drag_and_drop_start)
	with Window(u"Context||Menu"):
		with Region():
			left_click(u"Edit with Notepad++||MenuItem%(-32.11,14.29)")
	time.sleep(0.5)
	send_keys("{VK_LWIN down}""{VK_UP}""{VK_LWIN up}")
	brian.say("You can see that it is an easy to read Python script.")

	send_keys("{PGUP}""{VK_HOME}")

	brian.say("The functions return Pywinauto wrappers.")

	send_keys("{VK_DOWN 16}")
	send_keys("{VK_RIGHT 2}")
	send_keys("pywinauto_wrapper1 = ")
	brian.say("I'm going to call draw outline on some of these wrappers")
	send_keys("{VK_END}""{ENTER}")
	brian.say("The first draw outline will display a red rectangle around the One button.")
	send_keys("pywinauto_wrapper1.draw_outline{(}colour='red', thickness=22{)}")
	send_keys("{VK_HOME}""{VK_DOWN}")
	send_keys("pywinauto_wrapper2 = ")
	send_keys("{VK_END}""{ENTER}")
	brian.say("The second one will display a green rectangle around the Plus button.")
	send_keys("pywinauto_wrapper2.draw_outline{(}colour='green', thickness=22{)}")
	send_keys("{VK_HOME}""{VK_DOWN}")
	send_keys("pywinauto_wrapper3 = ")
	send_keys("{VK_END}""{ENTER}")
	brian.say("And the third one will display a blue rectangle around the Two button.")
	send_keys("pywinauto_wrapper3.draw_outline{(}colour='blue', thickness=22{)}")
	send_keys("{VK_CONTROL down}""s""{VK_CONTROL up}")
	brian.say("Let's replay it!", wait_until_the_end_of_the_sentence=True)
	send_keys("{VK_MENU down}""{VK_F4}""{VK_MENU up}")

	#############################################################
	# Drag and drop the recorded file on pywinauto_recorder.exe #
	#############################################################
	with Window(u"||List"):
		wrapper = find(u"Pywinauto recorder||ListItem")
	wrapper.draw_outline()
	drag_and_drop_end = get_wrapper_path(wrapper)
	drag_and_drop(drag_and_drop_start, drag_and_drop_end)
	brian.say("""
		Now Pywinauto recorder is replaying. You can see the Replay icon in the top left corner of the screen.
		The actions we recorded in the Python script are going to be replayed and the three draw outline functions we've just added are going to be called.
		""")
	time.sleep(25)
	brian.say("Thank you for watching!", wait_until_the_end_of_the_sentence=True)

	obs.stop_recording()
	obs.quit()

	exit(0)

