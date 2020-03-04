# -*- coding: utf-8 -*-

import unittest
from recorder_fn import *


class TestEntryMethods(unittest.TestCase):

	def test_get_entry_elements(self):
		entry_list = [
			'Name:::Type#[0,0]%(2,-24)', '::Type#[0,0]%(2,-24)', 'Name:::#[0,0]%(2,-24)',
			'Name:::Type#[0,0]', '::Type#[0,0]', 'Name:::#[0,0]',
			'Name:::Type', '::Type', 'Name:::',
			'Name::Type', '::Type', 'Name::'
		]
		for i, entry in enumerate(entry_list):
			str_name, str_type, y_x, dx_dy = get_entry(entry)
			#print get_entry(entry)
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
			else:
				self.assertEqual(y_x, None)

	def test_same_entry_list(self):
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

	def test_drag_and_drop(self): #With paint
		time.sleep(0.5)
		send_keys("{LWIN}Notepad{ENTER}")
		left_click("Untitled - Notepad::Window->Application::MenuBar->Format::MenuItem")
		left_click("Untitled - Notepad::Window->Format::Menu->Font...::MenuItem")
		line_down = find("Untitled - Notepad::Window->Font::Window->Size:::ComboBox->Size:::List->Vertical::ScrollBar->Line down::Button")
		position = find("Untitled - Notepad::Window->Font::Window->Size:::ComboBox->Size:::List->Vertical::ScrollBar->Position::Thumb")
		dy = line_down.rectangle().top - (position.rectangle().top + position.rectangle().bottom)/2
		drag_and_drop(
			"Untitled - Notepad::Window->Font::Window->Size:::ComboBox->Size:::List->Vertical::ScrollBar->Position::Thumb%(0,0)%(0,"+ str(dy) +")")
		size_list_box = find("Untitled - Notepad::Window->Font::Window->Size:::ComboBox->Size:::List%(0,0)")
		left_click("Untitled - Notepad::Window->Font::Window->Size:::ComboBox->Size:::List->Vertical::ScrollBar->Line down::Button%(-20,0)")
		self.assertEqual(size_list_box.get_selection()[0].name, size_list_box.children_texts()[-1])
		left_click("Untitled - Notepad::Window->Font::Window->Cancel::Button")
		left_click("Untitled - Notepad::Window->::TitleBar->Close::Button")


if __name__ == '__main__':
	unittest.main(verbosity=2)
