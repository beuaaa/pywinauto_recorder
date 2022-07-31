import pytest
from pywinauto_recorder.player import UIPath, find, send_keys, click, menu_click, set_text, move, select_file
from pywinauto_recorder.recorder import Recorder
import pyperclip
import random
import win32api
import win32con
from PIL import Image
import time
import os


################################################################################################
#                                   Tests using Windows 11 Paint                               #
################################################################################################

def percentage_white_pixels(recorded_image):
	"""
	It opens the image, counts the number of white pixels, and returns the percentage of white pixels
	:param recorded_image: the path to the image you want to analyze
	:return: The percentage of white pixels in the image.
	"""
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


@pytest.mark.parametrize('start_kill_app', ["mspaint"], indirect=True)
def test_recorder_mouse_move(start_kill_app):
	"""
	Tests the accuracy of the relative coordinates of an element.
	It records a series of drag-and-drops to draw lines and then plays them back to overlay the lines in white.
	"""
	with UIPath("Untitled - Paint||Window"):
		find().set_focus()
		find().maximize()
		click("*->Resize||Button")
		with UIPath("Resize and Skew||Window->*"):
			click("Pixels||RadioButton")
			set_text("Horizontal||Edit#[0,0]", "1900")  # Resize the image to 1900 pixels wide
			click("OK||Button")
		click("*->Line||Button")
		click("*->Size||Group->Size||Button")
		menu_click("3px")  # The thickness is set to 3px
		wrapper = move("*->Using Line tool on Canvas||Pane")
	
	time.sleep(1)
	recorder = Recorder()
	recorder.relative_coordinate_mode = True
	recorder.start_recording()
	time.sleep(2.0)
	
	for _ in range(9):
		x0 = random.randint(wrapper.rectangle().left+22, wrapper.rectangle().right-22)
		y0 = random.randint(wrapper.rectangle().top+22, wrapper.rectangle().bottom-22)
		move((x0, y0), duration=0)
		x, y = win32api.GetCursorPos()
		assert (x, y) == (x0, y0)
		
		send_keys("{VK_CONTROL down}{VK_SHIFT down}f{VK_SHIFT up}{VK_CONTROL up}", vk_packet=False)
		time.sleep(0.5)
		code = pyperclip.paste()
		words = code.split("%(")
		words = words[1].split(')"')
		move("Untitled - Paint||Window->*->Using Line tool on Canvas||Pane%(" + words[0] + ")", duration=0)
		x, y = win32api.GetCursorPos()
		assert (x, y) == (x0, y0)
		
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
		x1 = random.randint(wrapper.rectangle().left + 22, wrapper.rectangle().right - 22)
		y1 = random.randint(wrapper.rectangle().top + 22, wrapper.rectangle().bottom - 22)
		move((x1, y1), duration=0.4)  # This duration is mandatory for the recorder
		time.sleep(0.6)   # This pause is mandatory for the recorder
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
		time.sleep(0.5)		# This pause is mandatory for the recorder

	recorded_file = recorder.stop_recording()
	recorder.quit()

	# The recorded file is patched to remove all that is not needed
	data = ""
	with open(recorded_file) as text_file:
		for line in text_file:
			if not (("send_keys" in line) or ("wrapper" in line)):
				if "drag_and_drop" in line:
					data += line.replace("drag_and_drop", "right_drag_and_drop")
				else:
					data += line
	print(data)
	
	# Now the lines are overed in white using the previously recorded drag and drops
	compiled_code = compile(data, '<string>', 'exec')
	eval(compiled_code)
	os.remove(recorded_file)
	send_keys("{VK_CONTROL down}s{VK_CONTROL up}")
	recorded_image = str(recorded_file).replace('.py', '.png')
	select_file("Untitled - Paint||Window->Save As||Window", recorded_image)
	time.sleep(2)
	percentage = percentage_white_pixels(recorded_image)
	os.remove(recorded_image)
	assert percentage == 100, "All the pixels should be white"
