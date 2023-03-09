import pytest
import os
from pywinauto_recorder.player import PlayerSettings, UIPath, load_dictionary, shortcut, \
	find, find_all, move, click, double_click, triple_click
from pywinauto_recorder.recorder import Recorder
from pywinauto_recorder.core import get_wrapper_path, get_entry, get_entry_list
import time
import win32gui


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
	hwnd = None
	while not hwnd:
		time.sleep(0.1)
		hwnd = win32gui.FindWindow(None, 'Calculator')
	print(hwnd)
	win32gui.MoveWindow(hwnd, 0, 100, 400, 400, True)
	# with UIPath("Calculator||Window"):  # 'UIAWrapper' object had no 'move_window' until Pywinauto 0.7.0
	# 	find().move_window(0, 100, 400, 400, True)
	recorder = Recorder()
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
	recorder = Recorder()
	time.sleep(4)
	str_num = ["Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"]
	move("Calculator||Window->*->Zero||Button", duration=0)
	start_time = time.time()
	with UIPath("Calculator||Window->*"):
		for _ in range(9):
			for num in str_num:
				move(num+"||Button", duration=0)
				wait_recorder_ready(recorder, num+"||Button", sleep=0)
	duration = time.time() - start_time
	recorder.quit()
	assert duration < 30, "The duration of the loop is " + str(duration) + " s. It must be lower than 30 s"


@pytest.mark.parametrize('start_kill_app', ["calc"], indirect=True)
def test_player_performance(start_kill_app):
	""" Tests the performance to find a unique path. """
	old_mouse_move_duration = PlayerSettings.mouse_move_duration
	iterations = 22
	result_wrapper = find("Calculator||Window->*->RegEx: Display is||Text")
	for mouse_move_duration, expected_duration in [(0, 10), (-1, 1)]:
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
	             "Calculator||Window->||Custom->*->Trigonometry||Text",
	             "Calculator||Window->||Custom->*->Function||Text")
	
	with UIPath("Calculator||Window"):
		for ui_path in test_data:
			ref_text = find(ui_path).texts()[0]
			ocr_text = find("*->" + ref_text + "||OCR_Text").result[1]
			assert ocr_text == ref_text
