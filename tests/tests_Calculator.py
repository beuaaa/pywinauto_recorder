import pytest
import os
from pywinauto_recorder.player import PlayerSettings, UIPath, load_dictionary, shortcut, find_main_windows, \
	find, move, click, double_click, triple_click, move_window, start_application, focus_on_application, kill_application, connect_application
from pywinauto_recorder.recorder import Recorder
import time

################################################################################################
#                               Tests using Windows 11 Calculator                              #
################################################################################################

@pytest.mark.parametrize('start_kill_app', ["calc"], indirect=True)
def test_dictionary(start_kill_app):
	"""
	It opens the calculator, clicks 1, +, 2, and =, and then checks that the result is 3.
	"""
	load_dictionary("Calculator.key", "Calculator.def")
	with UIPath(shortcut("Calculator")):
		with UIPath(shortcut("Number pad")):
			click(shortcut("1"))
		with UIPath(shortcut("Standard operators")):
			click(shortcut("+"))
		with UIPath(shortcut("Number pad")):
			click(shortcut("2"))
		with UIPath(shortcut("Standard operators")):
			click(shortcut("="))
	with UIPath(shortcut("Calculator")):
		wrapper = find("RegEx: Display is ?||Text")
	assert wrapper.window_text() == 'Display is 3', wrapper.window_text() + '. The expected result is 3.'


@pytest.mark.parametrize('start_kill_app', ["calc"], indirect=True)
def test_asterisk(start_kill_app):
	"""
	It clicks the "1" button, then the "+" button, then the "2" button, then the "=" button.
	Then is does it again, but this time it searches within all the windows with UIPath("*").
	"""
	for root_path in ["Calculator||Window->*", "*->*"]:
		with UIPath(root_path):
			click("One||Button")
			click("Plus||Button")
			click("Two||Button")
			click("Equals||Button")


@pytest.mark.parametrize('start_kill_app', ["calc"], indirect=True)
def test_recorder_click(start_kill_app, wait_recorder_ready):
	""" Tests the ability to record all clicks. """
	focus_on_application(None)
	move_window("Calculator||Window", x=0, y=100, width=400, height=400)
	move("Calculator||Window->*->Zero||Button", duration=0)
	recorder = Recorder()
	time.sleep(4)
	recorder.start_recording()
	start_time = time.time()
	with UIPath("Calculator||Window->*"):
		move("Zero||Button", duration=0)
		move("One||Button", duration=0)
		wait_recorder_ready(recorder, "One||Button")
		click("One||Button")
		move("Two||Button", duration=0)
		wait_recorder_ready(recorder, "Two||Button")
		double_click("Two||Button")
		move("Three||Button", duration=0)
		wait_recorder_ready(recorder, "Three||Button")
		triple_click("Three||Button")
		move("Four||Button", duration=0)
		wait_recorder_ready(recorder, "Four||Button")
		triple_click("Four||Button")
		click("Four||Button")
		move("Five||Button", duration=0)
		wait_recorder_ready(recorder, "Five||Button")
		triple_click("Five||Button")
		double_click("Five||Button")
		move("Six||Button", duration=0)
		wait_recorder_ready(recorder, "Six||Button")
		triple_click("Six||Button")
		triple_click("Six||Button")
	duration = time.time() - start_time
	record_file_name = recorder.stop_recording()
	recorder.quit()

	assert duration < 8, "The duration of this test is " + str(duration) + " s. It must be lower than 8 s"
	
	str2num = {"Zero": 0, "One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5, "Six": 6}
	click_count = [0, 0, 0, 0, 0, 0, 0]
	with open(record_file_name, 'r') as f:
		line = f.readline()
		while line:
			line = f.readline()
			for str_number, i_number in str2num.items():
				if str_number in line:
					if "triple_click" in line:
						click_count[i_number] += 3
					elif "double_click" in line:
						click_count[i_number] += 2
					elif "click" in line:
						click_count[i_number] += 1
	for i_number in str2num.values():
		assert click_count[i_number] == i_number
	os.remove(record_file_name)


@pytest.mark.parametrize('start_kill_app', ["calc"], indirect=True)
def test_recorder_performance(start_kill_app, wait_recorder_ready):
	""" Tests the performance to find a unique path. """
	focus_on_application(None)
	move("Calculator||Window->*->Zero||Button", duration=0)
	recorder = Recorder()
	time.sleep(4)
	str_num = ["Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"]
	start_time = time.time()
	with UIPath("Calculator||Window->*"):
		for _ in range(9):
			for num in str_num:
				move(num+"||Button", duration=0)
				wait_recorder_ready(recorder, num+"||Button", sleep=0)
	duration = time.time() - start_time
	recorder.quit()
	print('duration:', duration)
	assert duration < 30, "The duration of the loop is " + str(duration) + " s. It must be lower than 30 s"


@pytest.mark.parametrize('start_kill_app', ["calc"], indirect=True)
def test_player_performance(start_kill_app):
	""" Tests the performance to find a unique path. """
	old_mouse_move_duration = PlayerSettings.mouse_move_duration
	iterations = 22
	result_wrapper = find("Calculator||Window->*->RegEx: Display is||Text")
	calculator_1 = connect_application(title="Calculator")
	focus_on_application(calculator_1)
	for mouse_move_duration, expected_duration in [(0, 11), (-1, 1)]:
		PlayerSettings.mouse_move_duration = mouse_move_duration
		start_time = time.time()
		with UIPath("Calculator||Window->*"):
			if result_wrapper.window_text() != "Display is 0":
				click("Clear entry||Button")
			for _ in range(iterations):
				click("One||Button")
				click("Zero||Button")
				click("Plus||Button")
			click("Equals||Button")
		assert result_wrapper.window_text() == "Display is " + str(10*iterations*2)
		duration = time.time() - start_time
		message = "Duration of the loop using 'duration=" + str(mouse_move_duration) + "': " + str(duration) + " s."
		print(message)
		assert duration < expected_duration, message + " It must be lower than " + str(expected_duration) + " s."
	PlayerSettings.mouse_move_duration = old_mouse_move_duration


@pytest.mark.parametrize('start_kill_app', ["calc"], indirect=True)
def test_ocr(start_kill_app):
	"""
	Searches the text of some elements of type "Text" and compares it with the text found by the OCR method.
	"""
	test_data = ("Calculator||Window->||Custom->Scientific Calculator mode||Text",
	             "Calculator||Window->||Custom->||Group->*->Trigonometry||Text",
	             "Calculator||Window->||Custom->||Group->*->Function||Text")
	
	with UIPath("Calculator||Window"):
		for ui_path in test_data:
			ref_text = find(ui_path).texts()[0]
			ocr_text = find("*->" + ref_text + "||OCR_Text").result[1]
			assert ocr_text == ref_text


#TODO: Insure that the Calculators started are killed even if the test fails.
def test_multi_instances():
	"""
	Demonstrates automation of two instances of the Calculator application.

	1. Starts two instances of Calculator.
	2. Performs actions on the first instance:
	    - Focuses on the application.
	    - Moves its window to a specified position and size.
	    - Clicks the 'One' button.
	3. Performs actions on the second instance using a UIPath context:
	    - Focuses on the application using an asterisk in UIPath.
	    - Moves the window to a different position and size.
	    - Performs a series of button clicks: 'Three', 'Minus', 'One', 'Equals'.
	4. Waits for 2 seconds to observe the changes.
	5. Closes both instances of the Calculator application.
	"""
	calculator_1 = start_application("calc")
	calculator_2 = start_application("calc")
	
	focus_on_application(calculator_1)
	move_window("Calculator||Window", x=1500, y=0, width=400, height=500)
	click("Calculator||Window->*->One||Button")
	
	focus_on_application(calculator_2)
	# This time, we'll use a UIPath context to simplify the paths to the user interface for multiple actions
	# As focus_on_application focus on the application calculator_2, we can use a UIPath with an asterisk
	# because there's no ambiguity when it comes to finding the application's main window.
	with UIPath("*"):
		move_window(x=1500, y=500, width=400, height=500)
		click("Three||Button")
		click("Minus||Button")
		click("One||Button")
		click("Equals||Button")
	
	time.sleep(2)
	kill_application(calculator_1)
	kill_application(calculator_2)


def test_new_window_connection():
	"""
	Opens a new application within an application, navigating through UI elements,
	and then connecting to the new window that opens as a result of a user action.

	Steps:
	1. Start the application "calc" (calculator).
	2. Focus on the calculator application.
	3. Click on the "Open Navigation" button.
	4. Click on the "Settings" list item.
	5. Record all main windows before clicking on a UI element.
	6. Click on the "Send feedback" hyperlink.
	7. Connect to the Feedback Hub window that opens after excluding previously recorded main windows.
	8. Focus on the Feedback Hub application.
	9. Resize the Feedback Hub window to 800x800 pixels.
	10. Click on the "Close Feedback Hub" button.
	"""
	calculator = start_application("calc")
	focus_on_application(calculator)
	with UIPath("*"):
		click("Open Navigation||Button")
		click("Settings||ListItem")
		all_main_windows_before_click = find_main_windows('*')
		time.sleep(1)  # wait for the animation completion
		click("Send feedback||Hyperlink")
	feedback_hub = connect_application(exclude_main_windows=all_main_windows_before_click, main_window_uipath="Feedback Hub||Window", timout=10)
	focus_on_application(feedback_hub)
	with UIPath("*"):
		move_window(x=0, y=0)
		click("Close Feedback Hub||Button")
	kill_application(calculator)



