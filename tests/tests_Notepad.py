import pytest
import time
from pywinauto_recorder.recorder import Recorder
from pywinauto_recorder.player import UIPath, send_keys, find, click, menu_click, move, mouse_wheel, playback
import re
import os


################################################################################################
#                              Tests using Windows 11 Notepad                                  #
################################################################################################

@pytest.mark.parametrize('start_kill_app', ["Notepad"], indirect=True)
def test_send_keys(start_kill_app):
	"""
	Opens Notepad, clicks on the text editor, and sends the keys "This is a error" followed
	by 5 backspaces, "test.", and an enter.
	"""
	with UIPath("Untitled - Notepad||Window"):
		# edit = click("Text Editor||Edit")  # Windows 10
		edit = click(u"Text editor||Document")  # Windows 11
		send_keys("This is a error{BACKSPACE}{BACKSPACE}{BACKSPACE}{BACKSPACE}{BACKSPACE}test.{ENTER}")
	result = edit.legacy_properties()['Value']
	# assert result == 'This is a test.\r\n'  # Windows 10
	assert result == 'This is a test.\r'  # Windows 11


@pytest.mark.parametrize('start_kill_app', ["Notepad"], indirect=True)
def test_wheel(start_kill_app):
	"""
	It opens the font dialog of Notepad,
	scrolls to the top of the font size list using the mouse wheel,  clicks on the first item in the list box,
	and asserts that the selected item is the first one.
	Then it scrolls to the bottom of the font size list using the mouse wheel, clicks on the last item in the list box,
	and then asserts that the selected item is the last one.
	"""
		
	with UIPath("Untitled - Notepad||Window"):
		menu_click("Edit->Font")
		
		font_list_box = click("*->Font||Group->Family||ComboBox%(76.86,-12.50)")
		move("*->Font||Group->Family||ComboBox%(0,400)")
		mouse_wheel(80, pause=0.05)
		click("*->Font||Group->Family||ComboBox->RegEx: .*||ListItem#[0, 0]")
		assert font_list_box.get_selection()[0].name == font_list_box.children_texts()[0]
		
		font_list_box = click("*->Font||Group->Family||ComboBox%(76.86,-12.50)")
		move("*->Font||Group->Family||ComboBox%(0,400)")
		mouse_wheel(-80, pause=0.05)
		click("*->Font||Group->Family||ComboBox->RegEx: .*||ListItem#[-1, 0]")
		assert font_list_box.get_selection()[0].name == font_list_box.children_texts()[-1]


@pytest.mark.parametrize('start_kill_app', ["Notepad"], indirect=True)
def test_recorder_menu_click(start_kill_app, wait_recorder_ready):
	def get_zoom_value():
		with UIPath("Untitled - Notepad||Window"):
			zoom_text = find("*->Zoom||Text")
			regex = r"\' \d+%\'"
			re_percentage_zoom = re.findall(regex, str(zoom_text.window_text))
			assert len(re_percentage_zoom) == 1
			percentage_zoom = int(re_percentage_zoom[0][2:-2])
			return percentage_zoom
	
	recorder = Recorder()
	recorder.process_menu_click_mode = True
	recorder.start_recording()
	with UIPath("Untitled - Notepad||Window"):
		for i in range(3):
			w = move("*->View||MenuItem", duration=0)
			wait_recorder_ready(recorder, "View||MenuItem", sleep=0)
			click(w, duration=0)
			time.sleep(0.5)  # wait for the menu to open (it is not always instantaneous depending on the animation settings)
			w = move("*->Zoom||MenuItem", duration=0)
			wait_recorder_ready(recorder, "Zoom||MenuItem", sleep=0)
			click(w, duration=0)
			time.sleep(0.5)  # wait for the menu to open (it is not always instantaneous depending on the animation settings)
			w = move("*->Zoom in||MenuItem", duration=0)
			wait_recorder_ready(recorder, "Zoom in||MenuItem", sleep=0)
			click(w, duration=0)
			time.sleep(0.5)  # wait for the menu to open (it is not always instantaneous depending on the animation settings)
	recorded_python_script = recorder.stop_recording()
	assert get_zoom_value() == 100 + (i + 1) * 10
	recorder.quit()
	
	playback(filename=recorded_python_script)
	assert get_zoom_value() == 100 + (i+1)*2*10
	os.remove(recorded_python_script)

