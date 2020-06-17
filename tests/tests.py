# -*- coding: utf-8 -*-

import unittest
import os
import platform
from pywinauto_recorder.player import *
from pywinauto_recorder.recorder import Recorder
import pyperclip
import random


class TestMouseMethods(unittest.TestCase):

	def setUp(self):
		"""Set some data and ensure the application is in the state we want"""
		self.app = pywinauto.Application()
		self.app.start("mspaint.exe")
		self.recorder = Recorder()
		self.recorder.start_recording()

	def tearDown(self):
		self.app.kill()
		self.recorder.stop_recording()

	def test_mouse_move(self):
		with Window(u"Sans titre - Paint||Window"):
			wrapper = find(u"||Pane->||Pane")
		wrapper.draw_outline()

		for i in range(10):
			x0 = random.randint(wrapper.rectangle().left, wrapper.rectangle().right)
			y0 = random.randint(wrapper.rectangle().top, wrapper.rectangle().bottom)
			move((x0, y0))
			x, y = win32api.GetCursorPos()
			print("0 " + str(x) + " " + str(y))

			self.assertEqual(x0, x)
			self.assertEqual(y0, y)


			time.sleep(0.5)
			send_keys("{VK_CONTROL down}""{VK_SHIFT down}""f""{VK_SHIFT up}""{VK_CONTROL up}", vk_packet=False)
			time.sleep(0.5)
			code = pyperclip.paste()
			words = code.split("%(")
			words = words[1].split(')"')
			with Window(u"Sans titre - Paint||Window"):
				move(u"||Pane->||Pane%(" + words[0] + ")")
			x, y = win32api.GetCursorPos()
			print("1 " + str(x) + " " + str(y))

			self.assertEqual(x0, x)
			self.assertEqual(y0, y)


class TestEntryMethods(unittest.TestCase):

	def test_get_entry_list(self):
		element_path = "D:\\Name||Window->||Pane->Property||Group->"
		element_path = element_path + "Label:||Text#[Name||Window->||Pane->Property||Group->Label:||Text,0]%(0,0)"
		entry_list = get_entry_list(element_path)
		self.assertEqual(entry_list[0], 'D:\\Name||Window')
		self.assertEqual(entry_list[1], '||Pane')
		self.assertEqual(entry_list[2], 'Property||Group')
		self.assertEqual(entry_list[3], 'Label:||Text#[Name||Window->||Pane->Property||Group->Label:||Text,0]%(0,0)')

	def test_get_entry_elements(self):
		entry_list = [
			'Name:||Type#[0,0]%(2,-24)', '||Type#[0,0]%(2,-24)', 'Name:||#[0,0]%(2,-24)',
			'Name:||Type#[0,0]', '||Type#[0,0]', 'Name:||#[0,0]',
			'Name:||Type', '||Type', 'Name:||',
			'Name||Type', '||Type', 'Name||',
			'Name||Type#[y_name:||y_type,0]', '||Type#[y_name:||y_type,0]', 'Name||#[y_name:||y_type,0]',
			'Name||Type#[y_name||y_type,0]', '||Type#[y_name||y_type,0]', 'Name||#[y_name||y_type,0]'
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
				self.assertEqual(str_type, None)
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
		time.sleep(0.5)
		send_keys("{LWIN}Notepad{ENTER}")

		with Region("Untitled - Notepad||Window"):
			edit = left_click("Text Editor||Edit%(0,0)")
			send_keys("This is a error{BACKSPACE}{BACKSPACE}{BACKSPACE}{BACKSPACE}{BACKSPACE}test.{ENTER}")
		result = edit.get_value()

		with Region("*Untitled - Notepad||Window"):
			left_click("||TitleBar->Close||Button%(0,0)")
			left_click("Notepad||Window->Don't Save||Button%(0,0)")

		self.assertEqual(result, 'This is a test.\r\n')

	def test_drag_and_drop(self):
		time.sleep(0.5)
		send_keys("{LWIN}Notepad{ENTER}")

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
		time.sleep(0.5)
		send_keys("{LWIN}Notepad{ENTER}")

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


@unittest.skipUnless(platform.system() == 'Windows' and platform.release() == '10', "requires Windows 10")
class TestCalculator(unittest.TestCase):

	def test_clicks(self):
		recorder = Recorder()
		recorder.start_recording()

		time.sleep(0.5)
		send_keys("{LWIN}Calculator{ENTER}")

		with Region("Calculator||Window->Calculator||Window->||Group->Number pad||Group"):
			move("One||Button")
			time.sleep(0.5)
			left_click("One||Button")
			move("Two||Button")
			time.sleep(0.5)
			double_left_click("Two||Button")
			move("Three||Button")
			time.sleep(0.5)
			triple_left_click("Three||Button")
			move("Four||Button")
			time.sleep(0.5)
			triple_left_click("Four||Button")
			left_click("Four||Button")
			move("Five||Button")
			time.sleep(0.5)
			triple_left_click("Five||Button")
			double_left_click("Five||Button")
			move("Six||Button")
			time.sleep(0.5)
			triple_left_click("Six||Button")
			triple_left_click("Six||Button")
			move("Three||Button")
			time.sleep(0.5)
			triple_left_click("Three||Button")
			move("Two||Button")
			time.sleep(0.5)
			double_left_click("Two||Button")
			move("One||Button")
			time.sleep(0.5)
			left_click("One||Button")

		with Region("Calculator||Window->Calculator||Window"):
			move("Close Calculator||Button")
			time.sleep(0.5)
			left_click("Close Calculator||Button")

		record_file_name = recorder.stop_recording()
		recorder.quit()

		with open(record_file_name, 'r') as f:
			line = f.readline()
			while line:
				line = f.readline()
				if "One" in line:
					self.assertTrue(line.find("left_click") != -1)
				elif "Two" in line:
					self.assertTrue(line.find("double_left_click") != -1)
				elif "Three" in line:
					self.assertTrue(line.find("triple_left_click") != -1)
				elif "Four" in line:
					self.assertTrue(line.find("triple_left_click") != -1)
					line = f.readline()
					self.assertTrue(line.find("left_click") != -1)
				elif "Five" in line:
					self.assertTrue(line.find("triple_left_click") != -1)
					line = f.readline()
					self.assertTrue(line.find("double_left_click") != -1)
				elif "Six" in line:
					self.assertTrue(line.find("triple_left_click") != -1)
					line = f.readline()
					self.assertTrue(line.find("triple_left_click") != -1)
		os.remove(record_file_name)


if __name__ == '__main__':
	unittest.main(verbosity=2)
	exit(0)
