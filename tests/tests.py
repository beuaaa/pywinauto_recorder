# -*- coding: utf-8 -*-

import unittest
import os
import platform
from pywinauto_recorder.player import *
from pywinauto_recorder.recorder import Recorder
from pywinauto_recorder.core import *
import pyperclip
import random
import win32api
import win32con
from PIL import Image
import pywinauto
import time

	
def percentage_white_pixels(recorded_image):
	img = Image.open(recorded_image)
	count = 0
	for y in range(img.height):
		for x in range(img.width):
			pixel = img.getpixel((x, y))
			if pixel == (255, 255, 255):
				count += 1
	p_white_pixels = 100 * count / (img.height * img.width)
	print(count, "/", img.height * img.width, " pixels are white: ", p_white_pixels, "%")
	return p_white_pixels


class TestMouseMethods(unittest.TestCase):

	def setUp(self):
		"""Set some data and ensure the application is in the state we want"""
		self.app = pywinauto.Application()
		self.app.start("mspaint.exe")

	def tearDown(self):
		time.sleep(0.5)
		self.app.kill()

	def test_mouse_move(self):
		""" Tests the precision of the relative coordinates in an element"""
		with Window("Untitled - Paint||Window"):
			find().set_focus()
			
			left_click("*->||Custom->Home||Custom->Image||ToolBar->Resize||Button")
			with Region("Resize and Skew||Window"):
				left_click("Pixels||RadioButton")
				left_click("Resize Horizontal||Edit")
				set_text("Resize Horizontal||Edit", "1900")

			left_click("*->||Custom->Home||Custom->Tools||ToolBar->Pencil||Button")
			send_keys("{VK_LWIN down}""{VK_UP}""{VK_LWIN up}")
			wrapper = move("||Pane->Using Pencil tool on Canvas||Pane")
		
		time.sleep(1)
		recorder = Recorder()
		recorder.relative_coordinate_mode = True
		recorder.start_recording()
		time.sleep(2.0)
		for i in range(9):
			win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
			x0 = random.randint(wrapper.rectangle().left+22, wrapper.rectangle().right-22)
			y0 = random.randint(wrapper.rectangle().top+22, wrapper.rectangle().bottom-22)
			move((x0, y0), duration=0.05)
			x, y = win32api.GetCursorPos()
			assert (x0, y0) == (x, y)
			
			send_keys("{VK_CONTROL down}""{VK_SHIFT down}""f""{VK_SHIFT up}""{VK_CONTROL up}", vk_packet=False)
			#time.sleep(0.5)
			code = pyperclip.paste()
			words = code.split("%(")
			words = words[1].split(')"')
			with Window(u"Untitled - Paint||Window"):
				move(u"||Pane->Using Pencil tool on Canvas||Pane%(" + words[0] + ")", duration=0)
			x, y = win32api.GetCursorPos()
			assert (x0, y0) == (x, y)
			
			win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
			time.sleep(0.5)		# This pause is mandatory for the recorder
			#wait_recorder_ready(recorder, "||Pane->Using Pencil tool on Canvas||Pane")
		recorded_file = recorder.stop_recording()
		recorder.quit()
		# time.sleep(5.5)
		
		# The thickness is set to 3px
		with Window(u"Untitled - Paint||Window"):
			left_click("UIRibbonDockTop||Pane->Ribbon||Pane->Ribbon||Pane->||Pane->Ribbon||Pane->Lower Ribbon||Pane->||Custom->Home||Custom->||ToolBar->Size||SplitButton")
			with Region(u"||Pane->Size||Window->Size||List->||Custom->||Group"):
				left_click(u"3px||ListItem")

		# Now the lines are overed in white using the previously recorded drag and drops
		with Window("Untitled - Paint||Window"):
			with Region("UIRibbonDockTop||Pane->Ribbon||Pane->Ribbon||Pane->||Pane->Ribbon||Pane"):
				left_click("Lower Ribbon||Pane->||Custom->Home||Custom->Colors||ToolBar->||Group%(-95,0)")
		data = ""
		with open(recorded_file) as fp:
			for line in fp:
				if not (("send_keys" in line) or ("wrapper" in line)):
					data = data + line
			
		print(data)
		with open(recorded_file, "w") as text_file:
			print(data, file=text_file)
		compiled_code = compile(data, '<string>', 'exec')
		eval(compiled_code)
		#os.remove(recorded_file)
		send_keys("{VK_CONTROL down}""s""{VK_CONTROL up}")
		recorded_image = str(recorded_file).replace('.py', '.png')
		send_keys(recorded_image, pause=0)
		send_keys("{ENTER}")
		time.sleep(2)
		percentage = percentage_white_pixels(recorded_image)
		#os.remove(recorded_image)
		assert percentage == 100, "All the pixels should be white"


class TestEntryMethods(unittest.TestCase):

	def test_get_entry_list(self):
		""" Tests get_entry_list(element_path) """
		element_path = "D:\\Name||Window->||Pane->Property||Group->"
		element_path = element_path + "Label:||Text#[Name||Window->||Pane->Property||Group->Label:||Text,0]%(0,0)"
		entry_list = get_entry_list(element_path)
		assert entry_list[0] == 'D:\\Name||Window'
		assert entry_list[1] == '||Pane'
		assert entry_list[2] == 'Property||Group'
		assert entry_list[3] == 'Label:||Text#[Name||Window->||Pane->Property||Group->Label:||Text,0]%(0,0)'
		
	def test_get_entry_list_with_asterisk(self):
		""" Tests get_entry_list(element_path) element_path has some asterisks"""
		element_path = "Name||Window->*->||Group->Property||%(0,0)"
		entry_list2 = get_entry_list(element_path)
		assert entry_list2[1] == '*'
		assert entry_list2[2] == '||Group'
		assert entry_list2[3] == 'Property||%(0,0)'
		
	def test_get_entry_elements(self):
		""" Tests get_entry(entry) """
		entry_list = [
			'Name:||Type#[0,0]%(2,-24)', '||Type#[0,0]%(2,-24)', 'Name:||#[0,0]%(2,-24)',
			'Name:||Type#[0,0]', '||Type#[0,0]', 'Name:||#[0,0]',
			'Name:||Type', '||Type', 'Name:||',
			'Name||Type', '||Type', 'Name||',
			'Name||Type#[y_name:||y_type,0]', '||Type#[y_name:||y_type,0]', 'Name||#[y_name:||y_type,0]',
			'Name||Type#[y_name||y_type,0]', '||Type#[y_name||y_type,0]', 'Name||#[y_name||y_type,0]',
		]

		for i, entry in enumerate(entry_list):
			str_name, str_type, y_x, dx_dy = get_entry(entry)
			if i % 3 == 0:
				if i < 9:
					self.assertEqual(str_name, 'Name:')
				else:
					self.assertEqual(str_name, 'Name')
				self.assertEqual(str_type, 'Type')
			if i % 3 == 1:
				self.assertEqual(str_name, '')
				self.assertEqual(str_type, 'Type')
			if i % 3 == 2:
				if i < 9:
					self.assertEqual(str_name, 'Name:')
				else:
					self.assertEqual(str_name, 'Name')
				self.assertEqual(None, str_type)
			if i < 3:
				self.assertEqual(dx_dy, (2, -24))
			else:
				self.assertEqual(dx_dy, None)
			if i < 6:
				self.assertEqual(y_x, [0, 0])
			elif i >= 15:
				self.assertEqual(y_x, ['y_name||y_type', 0])
			elif i >= 12:
				self.assertEqual(y_x, ['y_name:||y_type', 0])
			else:
				self.assertEqual(y_x, None)

	def test_same_entry_list(self):
		""" Tests same_entry_list(wrapper, entry_list) """
		time.sleep(0.5)
		send_keys("{LWIN}")
		element_path = 'Taskbar||Pane->Start||Button%(0,0)'
		entry_list = get_entry_list(element_path)
		with Region(''):
			element = find(element_path)
		result = same_entry_list(element, entry_list)
		self.assertTrue(result)

		element_path2 = 'Taskbar||Pane->Start2||Button%(0,0)'
		entry_list2 = get_entry_list(element_path2)
		result = same_entry_list(element, entry_list2)
		self.assertFalse(result)

		element_path3 = 'Taskbar||Pane->Start||Button2%(0,0)'
		entry_list3 = get_entry_list(element_path3)
		result = same_entry_list(element, entry_list3)
		self.assertFalse(result)

		send_keys("{LWIN}")


class TestNotepad(unittest.TestCase):

	def test_send_keys(self):
		""" Tests send keys """
		pywinauto.Application().start('Notepad')

		with Region("Untitled - Notepad||Window"):
			edit = left_click("Text Editor||Edit%(0,0)")
			send_keys("This is a error{BACKSPACE}{BACKSPACE}{BACKSPACE}{BACKSPACE}{BACKSPACE}test.{ENTER}")
		result = edit.get_value()

		with Region("*Untitled - Notepad||Window"):
			left_click("||TitleBar->Close||Button%(0,0)")
			left_click("Notepad||Window->Don't Save||Button%(0,0)")

		self.assertEqual(result, 'This is a test.\r\n')

	def test_drag_and_drop(self):
		""" Tests drag and drop """
		pywinauto.Application().start('Notepad')

		with Window("Untitled - Notepad||Window"):
			menu_click("", "Format->Font...", menu_type='NPP')

		with Window("Untitled - Notepad||Window->Font||Window"):
			size_list_box = find("Size:||ComboBox->Size:||List")
			with Region("Size:||ComboBox->Size:||List"):
				drag_and_drop("Vertical||ScrollBar->Position||Thumb", "Vertical||ScrollBar->Line down||Button")
				left_click("Vertical||ScrollBar->Line down||Button%(-200,0)")
			self.assertEqual(size_list_box.get_selection()[0].name, size_list_box.children_texts()[-1])
			left_click("Cancel||Button")

		with Window("Untitled - Notepad||Window"):
			left_click("||TitleBar->Close||Button")

	def test_wheel(self):
		""" Tests mouse wheel """
		pywinauto.Application().start('Notepad')

		with Window("Untitled - Notepad||Window"):
			menu_click("", "Format->Font...", menu_type='NPP')

		with Window("Untitled - Notepad||Window->Font||Window"):
			size_list_box = find("Size:||ComboBox->Size:||List")
			with Region("Size:||ComboBox->Size:||List"):
				left_click("Vertical||ScrollBar->Position||Thumb")
				mouse_wheel(-10)
				left_click("Vertical||ScrollBar->Line down||Button%(-200,0)")
			self.assertEqual(size_list_box.get_selection()[0].name, size_list_box.children_texts()[-1])
			left_click("Cancel||Button")

		with Window("Untitled - Notepad||Window"):
			left_click("||TitleBar->Close||Button")


def wait_recorder_ready(recorder, path_end, sleep=0.4):
	time.sleep(sleep)
	while l_e_e := recorder.get_last_element_event():
		x, y = win32api.GetCursorPos()
		if l_e_e.rectangle.top < y < l_e_e.rectangle.bottom:
			if l_e_e.rectangle.left < x < l_e_e.rectangle.right:
				if l_e_e.strategy == Strategy.unique_path:
					if path_end in l_e_e.path:
						break
		time.sleep(0.1)


def test_dictionary():
	load_dictionary("Calculator.key", "Calculator.def")
	os.system('calc.exe')
	time.sleep(1)
	with Window(shortcut("Calculator")):
		with Region(shortcut("Number pad")):
			left_click(shortcut("1"))
		with Region(shortcut("Standard operators")):
			left_click(shortcut("+"))
		with Region(shortcut("Number pad")):
			left_click(shortcut("2"))
		with Region(shortcut("Standard operators")):
			left_click(shortcut("="))
	display = None
	with Window(shortcut("Calculator")):
		results = find().children(control_type='Text')
		for result in results:
			if 'Display is' in result.window_text():
				display = result.window_text()
	assert display == 'Display is 3', display + '. The expected result is 3.'
	
	with Region("Calculator||Window->Calculator||Window"):
		left_click("Close Calculator||Button")


def test_asterisk():
	os.system('calc.exe')
	time.sleep(1)
	with Window("Calculator||Window"):
		left_click("*->One||Button")
		left_click("*->Plus||Button")
		left_click("*->Two||Button")
		left_click("*->Equals||Button")
		left_click("*->Close Calculator||Button")


@unittest.skipUnless(platform.system() == 'Windows' and platform.release() == '10', "requires Windows 10")
class TestCalculator(unittest.TestCase):

	def test_clicks(self):
		""" Tests the ability to record all clicks """
		os.system('calc.exe')
		import win32gui
		hwnd = win32gui.FindWindow(None, 'Calculator')
		win32gui.MoveWindow(hwnd, 0, 100, 400, 400, True)
		
		start_time = time.time()
		
		recorder = Recorder()
		recorder.start_recording()
		
		with Region("Calculator||Window->Calculator||Window->||Group->Number pad||Group"):
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
		with Region("Calculator||Window->Calculator||Window"):
			move("Close Calculator||Button", duration=0)
			wait_recorder_ready(recorder, "Close Calculator||Button")
			left_click("Close Calculator||Button")

		record_file_name = recorder.stop_recording()
		recorder.quit()
		
		duration = time.time() - start_time
		assert duration < 7, "The duration of this test is " + str(duration) + " s. It must be lower than 7 s"
		
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

	def test_recorder_performance(self):
		""" Tests the performance to find a unique path """
		recorder = Recorder()
		os.system('calc.exe')
		str_num = ["Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"]
		with Region("Calculator||Window->Calculator||Window->||Group->Number pad||Group"):
			move("Zero||Button", duration=0)
		startTime = time.time()
		with Region("Calculator||Window->Calculator||Window->||Group->Number pad||Group"):
			for i in range(9):
				for num in str_num:
					move(num+"||Button", duration=0)
					wait_recorder_ready(recorder, num+"||Button", sleep=0)
		duration = time.time() - startTime
		with Region("Calculator||Window->Calculator||Window"):
			move("Close Calculator||Button", duration=0)
			wait_recorder_ready(recorder, "Close Calculator||Button")
			left_click("Close Calculator||Button")
		recorder.quit()
		assert duration < 37, "The duration of the loop is " + str(duration) + " s. It must be lower than 37 s"


if __name__ == '__main__':
	unittest.main(verbosity=2)
	exit(0)
