from pywinauto import mouse
from pywinauto.win32structures import RECT
from pywinauto import win32defines, win32structures, win32functions
import ctypes
import time
import easyocr
from numpy import asarray as numpy_asarray


class OCRWrapper(object):
	reader = None
	
	def __init__(self, result):
		if OCRWrapper.reader is None:
			OCRWrapper.reader = easyocr.Reader(['fr'])  # this needs to run only once to load the model into memory
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


def find_ocr_str(str_to_find, wrapper, allowlist=None, realiability_min=0.98, time_out=9.0):
	if OCRWrapper.reader is None:
		OCRWrapper.reader = easyocr.Reader(['fr'])  # this needs to run only once to load the model into memory
	
	mag_ratio = 0
	reliability = -1
	t0 = time.time()
	while reliability <= realiability_min:
		mag_ratio += 1
		cropped_img = wrapper.capture_as_image()
		results = OCRWrapper.reader.readtext(numpy_asarray(cropped_img), allowlist=allowlist, batch_size=2, mag_ratio=mag_ratio)
		for ir, r in enumerate(results):
			print(r[1])
			if str_to_find in r[1]:
				for i in range(4):
					r[0][i][0] += wrapper.rectangle().left
					r[0][i][1] += wrapper.rectangle().top
				reliability = r[2]
				if reliability > realiability_min:
					return r
		t1 = time.time()
		if (t1 - t0) > time_out:
			raise TimeoutError

