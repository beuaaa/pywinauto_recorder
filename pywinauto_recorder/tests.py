# -*- coding: utf-8 -*-

import unittest
import os
from core import *
from player import *
from recorder import Recorder


class TestEntryMethods(unittest.TestCase):

	def test_get_entry_list(self):
		element_path = "Name::Window->::Pane->Property::Group->"
		element_path = element_path + "Label:::Text#[Name::Window->::Pane->Property::Group->Label:::Text,0]%(0,0)"
		entry_list = get_entry_list(element_path)
		self.assertEqual(entry_list[0], 'Name::Window')
		self.assertEqual(entry_list[1], '::Pane')
		self.assertEqual(entry_list[2], 'Property::Group')
		self.assertEqual(entry_list[3], 'Label:::Text#[Name::Window->::Pane->Property::Group->Label:::Text,0]%(0,0)')

	def test_get_entry_elements(self):
		entry_list = [
			'Name:::Type#[0,0]%(2,-24)', '::Type#[0,0]%(2,-24)', 'Name:::#[0,0]%(2,-24)',
			'Name:::Type#[0,0]', '::Type#[0,0]', 'Name:::#[0,0]',
			'Name:::Type', '::Type', 'Name:::',
			'Name::Type', '::Type', 'Name::',
			'Name::Type#[y_name::y_type,0]', '::Type#[y_name::y_type,0]', 'Name::#[y_name::y_type,0]',
			'Name::Type#[y_name:::y_type,0]', '::Type#[y_name:::y_type,0]', 'Name::#[y_name:::y_type,0]'
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
				self.assertEqual(y_x, ['y_name:::y_type', 0])
			elif i >= 12:
				self.assertEqual(y_x, ['y_name::y_type', 0])
			else:
				self.assertEqual(y_x, None)

	def test_same_entry_list(self):
		time.sleep(0.5)
		send_keys("{LWIN}")
		element_path = 'Taskbar::Pane->Start::Button%(0,0)'
		entry_list = (element_path.decode('utf-8')).split("->")
		element = find(element_path)
		result = same_entry_list(element, entry_list)
		self.assertTrue(result)

		element_path2 = 'Taskbar::Pane->Start2::Button%(0,0)'
		entry_list2 = (element_path2.decode('utf-8')).split("->")
		result = same_entry_list(element, entry_list2)
		self.assertFalse(result)

		element_path3 = 'Taskbar::Pane->Start::Button2%(0,0)'
		entry_list3 = (element_path3.decode('utf-8')).split("->")
		result = same_entry_list(element, entry_list3)
		self.assertFalse(result)

		send_keys("{LWIN}")


class TestNotepad(unittest.TestCase):

	def test_send_keys(self):
		time.sleep(0.5)
		send_keys("{LWIN}Notepad{ENTER}")
		edit = left_click("Untitled - Notepad::Window->Text Editor::Edit%(0,0)")
		send_keys("This is a error{BACKSPACE}{BACKSPACE}{BACKSPACE}{BACKSPACE}{BACKSPACE}test.{ENTER}")
		result = edit.get_value()
		left_click("*Untitled - Notepad::Window->::TitleBar->Close::Button%(0,0)")
		left_click("*Untitled - Notepad::Window->Notepad::Window->Don't Save::Button%(0,0)")
		self.assertEqual(result, 'This is a test.\r\n')

	def test_drag_and_drop(self):
		time.sleep(0.5)
		send_keys("{LWIN}Notepad{ENTER}")
		in_region("Untitled - Notepad::Window")
		left_click("Application::MenuBar->Format::MenuItem")
		left_click("Format::Menu->Font...::MenuItem")
		in_region("Untitled - Notepad::Window->Font::Window")
		line_down = find("Size:::ComboBox->Size:::List->Vertical::ScrollBar->Line down::Button")
		position = find("Size:::ComboBox->Size:::List->Vertical::ScrollBar->Position::Thumb")
		dy = line_down.rectangle().top - (position.rectangle().top + position.rectangle().bottom)/2
		drag_and_drop("Size:::ComboBox->Size:::List->Vertical::ScrollBar->Position::Thumb%(0,0)%(0," + str(dy) + ")")
		size_list_box = find("Size:::ComboBox->Size:::List%(0,0)")
		left_click("Size:::ComboBox->Size:::List->Vertical::ScrollBar->Line down::Button%(-20,0)")
		self.assertEqual(size_list_box.get_selection()[0].name, size_list_box.children_texts()[-1])
		left_click("Cancel::Button")
		in_region("Untitled - Notepad::Window")
		left_click("::TitleBar->Close::Button")
		in_region("")

	def test_wheel(self):
		time.sleep(0.5)
		send_keys("{LWIN}Notepad{ENTER}")
		in_region("Untitled - Notepad::Window")
		left_click("Application::MenuBar->Format::MenuItem")
		left_click("Format::Menu->Font...::MenuItem")
		in_region("Untitled - Notepad::Window->Font::Window")
		left_click("Size:::ComboBox->Size:::List->Vertical::ScrollBar->Position::Thumb")
		mouse_wheel(-10)
		size_list_box = find("Size:::ComboBox->Size:::List%(0,0)")
		left_click("Size:::ComboBox->Size:::List->Vertical::ScrollBar->Line down::Button%(-20,0)")
		self.assertEqual(size_list_box.get_selection()[0].name, size_list_box.children_texts()[-1])
		left_click("Cancel::Button")
		in_region("Untitled - Notepad::Window")
		left_click("::TitleBar->Close::Button")
		in_region("")


class TestCalculator(unittest.TestCase):

	def test_clicks(self):
		recorder = Recorder()
		record_file_name = recorder.start_recording()

		time.sleep(0.5)
		send_keys("{LWIN}Calculator{ENTER}")
		in_region("Calculator::Window->Calculator::Window->::Group->Number pad::Group")
		move("One::Button")
		time.sleep(0.5)
		left_click("One::Button")
		move("Two::Button")
		time.sleep(0.5)
		double_left_click("Two::Button")
		move("Three::Button")
		time.sleep(0.5)
		triple_left_click("Three::Button")
		move("Four::Button")
		time.sleep(0.5)
		triple_left_click("Four::Button")
		left_click("Four::Button")
		move("Five::Button")
		time.sleep(0.5)
		triple_left_click("Five::Button")
		double_left_click("Five::Button")
		move("Six::Button")
		time.sleep(0.5)
		triple_left_click("Six::Button")
		triple_left_click("Six::Button")
		move("Three::Button")
		time.sleep(0.5)
		triple_left_click("Three::Button")
		move("Two::Button")
		time.sleep(0.5)
		double_left_click("Two::Button")
		move("One::Button")
		time.sleep(0.5)
		left_click("One::Button")
		move("Calculator::Window->Calculator::Window->Close Calculator::Button")
		time.sleep(0.5)
		left_click("Calculator::Window->Calculator::Window->Close Calculator::Button")
		in_region("")

		recorder.stop_recording()
		recorder.quit()
		del recorder

		with open(record_file_name, 'r') as f:
			line = f.readline()
			while line:
				line = f.readline()
				if "One" in line:
					self.assertTrue(line[0:len("left_click")] == "left_click")
				elif "Two" in line:
					self.assertTrue(line[0:len("double_left_click")] == "double_left_click")
				elif "Three" in line:
					self.assertTrue(line[0:len("triple_left_click")] == "triple_left_click")
				elif "Four" in line:
					self.assertTrue(line[0:len("triple_left_click")] == "triple_left_click")
					line = f.readline()
					self.assertTrue(line[0:len("left_click")] == "left_click")
				elif "Five" in line:
					self.assertTrue(line[0:len("triple_left_click")] == "triple_left_click")
					line = f.readline()
					self.assertTrue(line[0:len("double_left_click")] == "double_left_click")
				elif "Six" in line:
					self.assertTrue(line[0:len("triple_left_click")] == "triple_left_click")
					line = f.readline()
					self.assertTrue(line[0:len("triple_left_click")] == "triple_left_click")
		os.remove(record_file_name)


if __name__ == '__main__':
	unittest.main(verbosity=2)
	exit(0)
