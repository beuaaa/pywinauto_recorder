class OCRWrapper:
	...


def find_all_ocr(wrapper, allowlist=None, mag_ratio=2, width_ths=0.5, contrast_ths=0.1, adjust_contrast=0.5,
                 rotation_info=None):
	return []

r'''
from pywinauto import mouse
from pywinauto.win32structures import RECT
from pywinauto import win32defines, win32structures, win32functions
import ctypes
import time
import traceback
from numpy import asarray as numpy_asarray


class OCRWrapper(object):
	reader = None
	
	def __init__(self, result):
		if OCRWrapper.reader is None:
			import easyocr
			try:
				OCRWrapper.reader = easyocr.Reader(['fr', 'en'])  # this needs to run only once to load the model into memory
			except ImportError:
				print("Easyocr not found ! Install it if you want to use the OCR_Text type.")
				raise
		self.result = result
	
	def click(self):
		x = (self.result[0][0][0] + self.result[0][2][0]) / 2
		y = (self.result[0][0][1] + self.result[0][2][1]) / 2
		mouse.click(coords=(x, y))
	
	def click_input(self):
		self.click()
	
	def rectangle(self):
		rect = RECT()
		rect.left = self.result[0][0][0]
		rect.top = self.result[0][0][1]
		rect.right = self.result[0][2][0]
		rect.bottom = self.result[0][2][1]
		return rect
	
	def is_enabled(self):
		return True
	
	def is_visible(self):
		return True
	
	def draw_outline(self,
	                 colour='green',
	                 thickness=2,
	                 fill=win32defines.BS_NULL,
	                 rect=None):
		"""
		Draw an outline around the window.
		* **colour** can be either an integer or one of 'red', 'green', 'blue'
		  (default 'green')
		* **thickness** thickness of rectangle (default 2)
		* **fill** how to fill in the rectangle (default BS_NULL)
		* **rect** the coordinates of the rectangle to draw (defaults to
		  the rectangle of the control)
		"""
		# don't draw if dialog is not visible
		if not self.is_visible():
			return
		
		colours = {
			"green": 0x00ff00,
			"blue": 0xff0000,
			"red": 0x0000ff,
		}
		
		# if it's a known colour
		if colour in colours:
			colour = colours[colour]
		
		if rect is None:
			rect = self.rectangle()
		
		# create the pen(outline)
		pen_handle = win32functions.CreatePen(
			win32defines.PS_SOLID, thickness, colour)
		
		# create the brush (inside)
		brush = win32structures.LOGBRUSH()
		brush.lbStyle = fill
		brush.lbHatch = win32defines.HS_DIAGCROSS
		brush_handle = win32functions.CreateBrushIndirect(ctypes.byref(brush))
		
		# get the Device Context
		dc = win32functions.CreateDC("DISPLAY", None, None, None)
		
		# push our objects into it
		win32functions.SelectObject(dc, brush_handle)
		win32functions.SelectObject(dc, pen_handle)
		
		# draw the rectangle to the DC
		win32functions.Rectangle(
			dc, rect.left, rect.top, rect.right, rect.bottom)
		
		# Delete the brush and pen we created
		win32functions.DeleteObject(brush_handle)
		win32functions.DeleteObject(pen_handle)
		
		# delete the Display context that we created
		win32functions.DeleteDC(dc)


def find_all_ocr(wrapper, allowlist=None, mag_ratio=2, width_ths=0.5, contrast_ths=0.1, adjust_contrast=0.5,
                 rotation_info=None):
	"""
	It takes an image, a searching area, and a list of allowed characters, and returns a list of all the text it finds in
	the image

	:param wrapper: the wrapper and its region to be searched
	:param searching_area: the area to search for text
	:param allowlist: a list of characters that you want to search for. If you don't want to search for any specific characters, just leave it as None
	:param width_ths: the minimum width of a character to be recognized
	:param contrast_ths: text box with contrast lower than this value will be passed into model 2 times. First is with original image and second with contrast adjusted to 'adjust_contrast' value. The one with more confident level will be returned as a result.
	:param adjust_contrast: target contrast level for low contrast text box
	:return: A list of tuples. Each tuple contains the text, the bounding box, and the confidence.
	"""
	if OCRWrapper.reader is None:
		try:
			import easyocr
			OCRWrapper.reader = easyocr.Reader(['fr', 'en'])
		except ImportError as ie:
			print(
				"EasyOCR not found ! Install it if you want to use the OCR_Text type. See installation instructions at https://github.com/jaidedai/easyocr")
			raise
	cropped_img = wrapper.capture_as_image()
	results = OCRWrapper.reader.readtext(numpy_asarray(cropped_img), width_ths=width_ths, allowlist=allowlist,
	                                     batch_size=2,
	                                     mag_ratio=mag_ratio, contrast_ths=contrast_ths, adjust_contrast=adjust_contrast,
	                                     rotation_info=rotation_info)
	for r in results:
		for i in range(4):
			r[0][i][0] += wrapper.rectangle().left
			r[0][i][1] += wrapper.rectangle().top
	results.sort(key=lambda r: (r[0][0][1], r[0][1][0]))
	return results
'''