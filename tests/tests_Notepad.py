import pytest
from pywinauto_recorder.player import UIPath, find, send_keys, left_click, menu_click, drag_and_drop, mouse_wheel


################################################################################################
#                                    Tests using Notepad                                       #
################################################################################################

@pytest.mark.parametrize('run_app', [("Notepad.exe", "Untitled - Notepad")], indirect=True)
def test_send_keys(run_app):
	"""
	Opens Notepad, clicks on the text editor, and sends the keys "This is a error" followed
	by 5 backspaces, "test.", and an enter.
	"""
	with UIPath("Untitled - Notepad||Window"):
		edit = left_click("Text Editor||Edit%(0,0)")
		send_keys("This is a error{BACKSPACE}{BACKSPACE}{BACKSPACE}{BACKSPACE}{BACKSPACE}test.{ENTER}")
	result = edit.get_value()
	assert result == 'This is a test.\r\n'


@pytest.mark.parametrize('run_app', [("Notepad.exe", "Untitled - Notepad")], indirect=True)
def test_drag_and_drop(run_app):
	"""
	It opens the font dialog of Notepad, scrolls to the bottom of the font size list using
	a drag and drop, clicks on the last item in the list box,
	and then asserts that the selected item is the last one.
	"""
	with UIPath("Untitled - Notepad||Window"):
		menu_click("", "Format->Font...", menu_type='NPP')
	with UIPath("Untitled - Notepad||Window->Font||Window"):
		size_list_box = find("Size:||ComboBox->Size:||List")
		with UIPath("Size:||ComboBox->Size:||List"):
			drag_and_drop("Vertical||ScrollBar->Position||Thumb", "Vertical||ScrollBar->Line down||Button")
			left_click("Vertical||ScrollBar->Line down||Button%(-200,0)")  # select the last item in the list
		assert size_list_box.get_selection()[0].name == size_list_box.children_texts()[-1]


@pytest.mark.parametrize('run_app', [("Notepad.exe", "Untitled - Notepad")], indirect=True)
def test_wheel(run_app):
	"""
	It opens the font dialog of Notepad, scrolls to the bottom of the font size list using
	the mouse wheel, clicks on the last item in the list box,
	and then asserts that the selected item is the last one.
	"""
	with UIPath("Untitled - Notepad||Window"):
		menu_click("", "Format->Font...", menu_type='NPP')
	with UIPath("Untitled - Notepad||Window->Font||Window"):
		size_list_box = find("Size:||ComboBox->Size:||List")
		with UIPath("Size:||ComboBox->Size:||List"):
			left_click("Vertical||ScrollBar->Position||Thumb")
			mouse_wheel(-10)
			left_click("Vertical||ScrollBar->Line down||Button%(-200,0)")  # select the last item in the list
		assert size_list_box.get_selection()[0].name == size_list_box.children_texts()[-1]
