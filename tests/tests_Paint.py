import pytest
from pywinauto_recorder.player import UIPath, find, send_keys, left_click, set_text, move, select_file
from pywinauto_recorder.recorder import Recorder
import pyperclip
import random
import win32api
import win32con
from PIL import Image
import time
import os


################################################################################################
#                                   Tests using Windows 10 Paint                               #
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


@pytest.mark.parametrize('run_app', [("mspaint.exe", "Untitled - Paint")], indirect=True)
def test_mouse_move(run_app):
	"""
	Tests the precision of the relative coordinates in an element.
	It records a series of drag and drops to draw lines
	and then replays them to cover the lines in white.
	"""
	with UIPath("Untitled - Paint||Window"):
		find().set_focus()
		
		left_click("*->||Custom->Home||Custom->Image||ToolBar->Resize||Button")
		with UIPath("Resize and Skew||Window"):
			left_click("Pixels||RadioButton")
			left_click("Resize Horizontal||Edit")
			set_text("Resize Horizontal||Edit", "1900")
		
		left_click("*->||Custom->Home||Custom->Shapes||ToolBar->Shapes||Group->Shapes||List->||Custom->Line||ListItem")
		left_click("*->||ToolBar->Size||SplitButton")
		with UIPath("||Pane->Size||Window->Size||List->||Custom->||Group"):
			left_click("3px||ListItem")
		send_keys("{VK_LWIN down}{VK_UP}{VK_LWIN up}")
		wrapper = move("||Pane->Using Line tool on Canvas||Pane")
	
	time.sleep(1)
	recorder = Recorder()
	recorder.relative_coordinate_mode = True
	recorder.start_recording()
	time.sleep(2.0)
	for _ in range(9):
		
		x0 = random.randint(wrapper.rectangle().left+22, wrapper.rectangle().right-22)
		y0 = random.randint(wrapper.rectangle().top+22, wrapper.rectangle().bottom-22)
		move((x0, y0), duration=0.1)
		x, y = win32api.GetCursorPos()
		assert (x0, y0) == (x, y)
		
		send_keys("{VK_CONTROL down}{VK_SHIFT down}f{VK_SHIFT up}{VK_CONTROL up}", vk_packet=False)
		time.sleep(0.5)
		code = pyperclip.paste()
		words = code.split("%(")
		words = words[1].split(')"')
		with UIPath("Untitled - Paint||Window"):
			move("||Pane->Using Line tool on Canvas||Pane%(" + words[0] + ")", duration=0)
		x, y = win32api.GetCursorPos()
		assert (x0, y0) == (x, y)
		
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
		x1 = random.randint(wrapper.rectangle().left + 22, wrapper.rectangle().right - 22)
		y1 = random.randint(wrapper.rectangle().top + 22, wrapper.rectangle().bottom - 22)
		move((x1, y1), duration=0.1)
		time.sleep(0.6)   # This pause is mandatory for the recorder
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
		time.sleep(0.5)		# This pause is mandatory for the recorder
	
	recorded_file = recorder.stop_recording()
	recorder.quit()
	# time.sleep(5.5)
	
	# The thickness is set to 3px
	with UIPath("Untitled - Paint||Window"):
		left_click("*->||Custom->Home||Custom->Tools||ToolBar->Pencil||Button")
		left_click("*->||Custom->Home||Custom->Shapes||ToolBar->Shapes||Group->Shapes||List->||Custom->Line||ListItem")
		left_click("*->||ToolBar->Size||SplitButton")
		with UIPath("||Pane->Size||Window->Size||List->||Custom->||Group"):
			left_click("3px||ListItem")

	# The white color is selected
	with UIPath("Untitled - Paint||Window"):
		left_click("*->RegEx: Colo(u)?rs||ToolBar->||Group%(-95,0)")
		
	# The recorded file is patched to remove all that is not needed
	data = ""
	with open(recorded_file) as fp:
		for line in fp:
			if not (("send_keys" in line) or ("wrapper" in line)):
				data = data + line
	print(data)
	with open(recorded_file, "w") as text_file:
		print(data, file=text_file)
	
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
