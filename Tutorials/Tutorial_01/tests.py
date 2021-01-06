# coding: utf-8
import time
import os
import shutil
from dp_toolbox import *
from pywinauto_recorder.player import *

def test_chrome():
	chrome_dir = r'"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"'
	chrome = pywinauto.Application(backend='uia')
	chrome.start(chrome_dir + ' --force-renderer-accessibility --start-maximized --guest ')
	
# send_keys("{VK_CONTROL down}""{VK_MENU down}""q""{VK_MENU up}""{VK_CONTROL up}")  # works
# send_keys("{VK_MENU down}""{VK_CONTROL down}""q""{VK_CONTROL up}""{VK_MENU up}") # does not work

if __name__ == '__main__':
	brian = Voice(name='Brian')

	obs = OBSStudio(
		r'D:\workbench-2020\Software\AutomatedValidation\SE-VALIDATION\OBSStudio\OBS-Studio-22.0.2-Full-x64\bin\64bit\obs64.exe')

	###############################################################################
	# Remove Downloads\pywinauto_recorder.dist.zip and R:\pywinauto_recorder.dist #
	###############################################################################
	if os.path.isfile(r'D:\Users\d_pra\Downloads\pywinauto_recorder.exe'):
		os.remove(r'D:\Users\d_pra\Downloads\pywinauto_recorder.exe')
	shutil.rmtree(r'R:\Record_files', ignore_errors=True)

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
		text_wrapper = None
		while not text_wrapper:
			try:
				wrapper = find()
				text_wrapper = wrapper.descendants(
					title=u"” is a standalone application, it’s the compiled version for Windows.",
					control_type="Text")
			except:
				time.sleep(1)
				continue
		wrapper_path = get_wrapper_path(text_wrapper[0])
	left_click(wrapper_path+"%(-60,0)")

	brian.say("""
	"Pywinauto recorder" is a unique "record-replay" tool in the open source field because it generates reliable scripts without hard-coded coordinates thanks to Pywinauto.
	Pywinauto is a library that uses accessibility technologies allowing you to automate almost any type of GUI.
	The functions of the generated Python script return Pywinauto wrappers so it can be enhanced with Pywinauto functions.
	""")
	brian.say("Let's get started! Download Pywinauto recorder for Windows on this web page")
	send_keys("{PGDN}")
	time.sleep(1)
	with Window(pywinauto_recorder_documentation_page):
		wrapper = find("")
		button_wrapper = wrapper.descendants(
			title=u"https://raw.githubusercontent.com/beuaaa/pywinauto_recorder/master/Images/Download.png?sanitize=true",
			control_type="Hyperlink")
		highlight(button_wrapper[0], color=(0, 0, 255))
	move(get_wrapper_path(button_wrapper[0]))
	brian.say("Click on this download button.", wait_until_the_end_of_the_sentence=True)
	clear_overlay()

	#########################################################################
	# Wait until the exe is downloaded and on show it in an Explorer window #
	#########################################################################
	brian.say("Locate the downloaded file and move it in a convenient folder.")
	with Window(pywinauto_recorder_documentation_page):
		wrapper = move(u"Google Chrome||Pane->||Pane-> pywinauto_recorder.exe||Button%(-40.85,1.64)")
		left_click(u"Google Chrome||Pane->||Pane->Options menu||Button%(-8.70,0.00)")
		left_click(u"||Pane->||MenuBar->||Pane->||Menu->Show in folder||MenuItem%(-23.47,6.67)")
	time.sleep(1.0)
	send_keys("{VK_LWIN down}""{VK_LEFT}""{VK_LWIN up}")
	with Window(u"Snap Assist||Window"):
		with Region(u"||Group->||Group->Running Applications||List"):
			left_click(pywinauto_recorder_documentation_page.replace(" (Guest)||Pane", "||ListItem->Close||Button"))

	#####################################
	# Move pywinauto_recorder.exe in R: #
	#####################################
	time.sleep(1.5)
	send_keys("{VK_CONTROL down}""x""{VK_CONTROL up}")
	with Window(u"D:\\Users\\d_pra\\Downloads||Window"):
		left_click(u"||Pane->||Pane->||ProgressBar->||Pane->Address: D:\\Users\\d_pra\\Downloads||ToolBar%(35.26,-5.13)")
		send_keys("r:""{ENTER}")
	send_keys("{VK_CONTROL down}""v""{VK_CONTROL up}")

	##############################
	# Run pywinauto_recorder.exe #
	##############################
	brian.say("Run 'pywinauto recorder'")
	with Window(u"R:\\||Window"):
		with Region(u"R:\\||Pane->||Pane->Shell Folder View||Pane"):
			p_r_exe = u"Items View||List->pywinauto_recorder.exe||ListItem"
			wrapper = find(p_r_exe)
			wrapper.draw_outline(colour="blue")
			wrapper = double_left_click(p_r_exe)
			time.sleep(6)
	drag_and_drop_end = get_wrapper_path(wrapper)
	right_click(drag_and_drop_end)
	brian.say(
		"""
		Now Pywinauto Recorder colors green the interface elements it can identify under the mouse cursor.
		It started in Pause mode which is represented by the first icon in the top left corner of the screen.
		""")
	move(drag_and_drop_end+"%(50,220)", duration=5)
	brian.say("To record my next actions I press 'alt' 'control' 'r'.", wait_until_the_end_of_the_sentence=True)
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

	##################################################
	# Stop Recording and quit pywinauto_recorder.exe #
	##################################################
	brian.say("I press 'alt' 'control' 'r' to stop recording.")
	send_keys("{VK_CONTROL down}""{VK_MENU down}""r""{VK_MENU up}""{VK_CONTROL up}")
	brian.say("And 'alt' 'control' 'q' to quit Pywinauto recorder.")
	time.sleep(2)
	send_keys("{VK_CONTROL down}""{VK_MENU down}""q""{VK_MENU up}""{VK_CONTROL up}")

	#######################################################
	# Open in a new Explorer window 'Record files' folder #
	#######################################################
	brian.say("The generayed file is in 'Record files' folder.")
	with Window(u"R:\\||Window"):
		with Region(u"R:\\||Pane->||Pane->Shell Folder View||Pane"):
			r_f = u"Items View||List->Record_files||ListItem"
			wrapper = find(r_f)
			wrapper.draw_outline(colour="blue")
			right_click(r_f)
	with Window(u"Context||Menu"):
		with Region():
			left_click(u"Open in new window||MenuItem%(-17.98,0.00)")
			time.sleep(1.0)
			send_keys("{LWIN down}""{VK_RIGHT}""{LWIN up}")

	############################
	# Locate the recorded file #
	############################
	with Window(u"R:\\Record_files||Window"):
		with Region(u"R:\\Record_files||Pane"):
			wrapper = find(u"||Pane->Shell Folder View||Pane->Items View||List")
			wrapper = wrapper.descendants(title='Name', control_type='Edit')[1]
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

