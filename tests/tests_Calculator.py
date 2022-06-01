# -*- coding: utf-8 -*-

import pytest
import os
from pywinauto_recorder.player import UIPath, load_dictionary, shortcut, \
	find, move, left_click, double_left_click, triple_left_click
from pywinauto_recorder.recorder import Recorder
from pywinauto_recorder.core import Strategy
import pywinauto
import time
import win32api


@pytest.fixture
def run_app(request):
	app_name, window_name = request.param
	app = pywinauto.Application(backend="win32")
	app.start(app_name)
	time.sleep(1)
	app.connect(title=window_name, timeout=10)
	yield app
	app.kill(soft=True)


def wait_recorder_ready(recorder, path_end, sleep=0.5):
	time.sleep(sleep)
	while l_e_e := recorder.get_last_element_event():
		x, y = win32api.GetCursorPos()
		if l_e_e.rectangle.top < y < l_e_e.rectangle.bottom:
			if l_e_e.rectangle.left < x < l_e_e.rectangle.right:
				if l_e_e.strategy == Strategy.unique_path:
					if path_end in l_e_e.path:
						break
		time.sleep(0.1)


################################################################################################
#                               Tests using Windows 10 Calculator                              #
################################################################################################

@pytest.mark.parametrize('run_app', [("calc.exe", "Calculator")], indirect=True)
def test_dictionary(run_app):
	"""
	It opens the calculator, clicks 1, +, 2, and =, and then checks that the result is 3.
	"""
	load_dictionary("Calculator.key", "Calculator.def")
	with UIPath(shortcut("Calculator")):
		with UIPath(shortcut("Number pad")):
			left_click(shortcut("1"))
		with UIPath(shortcut("Standard operators")):
			left_click(shortcut("+"))
		with UIPath(shortcut("Number pad")):
			left_click(shortcut("2"))
		with UIPath(shortcut("Standard operators")):
			left_click(shortcut("="))
	with UIPath(shortcut("Calculator")):
		wrapper = find("RegEx: Display is ?||Text")
	assert wrapper.window_text() == 'Display is 3', wrapper.window_text() + '. The expected result is 3.'


@pytest.mark.parametrize('run_app', [("calc.exe", "Calculator")], indirect=True)
def test_asterisk(run_app):
	"""
	It clicks the "1" button, then the "+" button, then the "2" button, then the "=" button.
	Then is does it again, but this time it searches within all the windows with UIPath("*").
	"""
	with UIPath("Calculator||Window"):
		left_click("*->One||Button")
		left_click("*->Plus||Button")
		left_click("*->Two||Button")
		left_click("*->Equals||Button")
	with UIPath("*"):
		left_click("*->One||Button")
		left_click("*->Plus||Button")
		left_click("*->Two||Button")
		left_click("*->Equals||Button")


@pytest.mark.parametrize('run_app', [("calc.exe", "Calculator")], indirect=True)
def test_clicks(run_app):
	""" Tests the ability to record all clicks. """
	import win32gui
	hwnd = win32gui.FindWindow(None, 'Calculator')
	win32gui.MoveWindow(hwnd, 0, 100, 400, 400, True)
	
	start_time = time.time()
	
	recorder = Recorder()
	recorder.start_recording()
	
	with UIPath("Calculator||Window->Calculator||Window->||Group->Number pad||Group"):
		move("Zero||Button", duration=0)
		move("One||Button", duration=0)
		wait_recorder_ready(recorder, "One||Button")
		left_click("One||Button")
		move("Two||Button", duration=0)
		wait_recorder_ready(recorder, "Two||Button")
		double_left_click("Two||Button")
		move("Three||Button", duration=0)
		wait_recorder_ready(recorder, "Three||Button")
		triple_left_click("Three||Button")
		move("Four||Button", duration=0)
		wait_recorder_ready(recorder, "Four||Button")
		triple_left_click("Four||Button")
		left_click("Four||Button")
		move("Five||Button", duration=0)
		wait_recorder_ready(recorder, "Five||Button")
		triple_left_click("Five||Button")
		double_left_click("Five||Button")
		move("Six||Button", duration=0)
		wait_recorder_ready(recorder, "Six||Button")
		triple_left_click("Six||Button")
		triple_left_click("Six||Button")
	with UIPath("Calculator||Window->Calculator||Window"):
		move("Close Calculator||Button", duration=0)
		wait_recorder_ready(recorder, "Close Calculator||Button")
		left_click("Close Calculator||Button")

	record_file_name = recorder.stop_recording()
	recorder.quit()
	
	duration = time.time() - start_time
	assert duration < 8.1, "The duration of this test is " + str(duration) + " s. It must be lower than 7 s"
	
	str2num = {"Zero": 0, "One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5, "Six": 6}
	click_count = [0, 0, 0, 0, 0, 0, 0]
	with open(record_file_name, 'r') as f:
		line = f.readline()
		while line:
			line = f.readline()
			for str_number, i_number in str2num.items():
				if str_number in line:
					if "triple_left_click" in line:
						click_count[i_number] += 3
					elif "double_left_click" in line:
						click_count[i_number] += 2
					elif "left_click" in line:
						click_count[i_number] += 1
	for i_number in str2num.values():
		assert click_count[i_number] == i_number
	os.remove(record_file_name)


@pytest.mark.parametrize('run_app', [("calc.exe", "Calculator")], indirect=True)
def test_recorder_performance(run_app):
	""" Tests the performance to find a unique path. """
	recorder = Recorder()
	time.sleep(1)  # wait for the recorder to start
	str_num = ["Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"]
	with UIPath("Calculator||Window->Calculator||Window->||Group->Number pad||Group"):
		move("Zero||Button", duration=0)
	start_time = time.time()
	with UIPath("Calculator||Window->Calculator||Window->||Group->Number pad||Group"):
		for _ in range(9):
			for num in str_num:
				move(num+"||Button", duration=0)
				# wait_recorder_ready seems not working like expected => TODO: recode the test
				wait_recorder_ready(recorder, num+"||Button", sleep=0)
	duration = time.time() - start_time
	recorder.quit()
	assert duration < 35, "The duration of the loop is " + str(duration) + " s. It must be lower than 35 s"


