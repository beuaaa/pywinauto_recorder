"""
The tooltip that displays information about the element under the mouse cursor.

The tooltip runs in its own process, so that the slow requests of UI Automation do not slow down the recorder.
"""

import overlay_arrows_and_more as oaam
from pywinauto_recorder.core import path_separator, get_wrapper_path, get_entry
from pywinauto import Desktop
from win32api import GetSystemMetrics, GetCursorPos
from time import sleep
from multiprocessing import Process, Event
from pywinauto_recorder.recorder import IconSet, _find_common_path, _overlay_add_mode_icon

wrapper_old_info_tip = None
common_path_info_tip = ''

_TOOLTIP_WIDTH = 500
"""Width of the tooltip, in pixels."""

_TOOLTIP_LINE_HEIGHT = 16
"""Height of a line of text of the tooltip, in pixels."""

_TEXT_FORMAT = "win32con.DT_LEFT|win32con.DT_WORDBREAK|win32con.DT_EXPANDTABS|win32con.DT_EDITCONTROL|win32con.DT_NOCLIP"
"""Format of the texts of the tooltip."""


def _text_height(text, char_width, height):
	"""
	Returns the height of a box of the tooltip that displays a text.

	:param text: The text
	:param char_width: The estimated width of a character of the text, in pixels
	:param height: The height of the box without the text
	:return: height, plus the height of a line for each line break of the text, and for each line that is wrapped
		because it is wider than the tooltip
	"""
	text_width = 0
	for c in text:
		if c == '\n':
			text_width = 0
			height = height + _TOOLTIP_LINE_HEIGHT
		else:
			text_width = text_width + char_width
		if text_width > _TOOLTIP_WIDTH:
			text_width = 0
			height = height + _TOOLTIP_LINE_HEIGHT
	return height


def _ancestor_with_path(wrapper, path):
	"""Returns the ancestor of an element whose path is 'path', or None if there is none."""
	parent_wrapper = wrapper.parent()
	while parent_wrapper and get_wrapper_path(parent_wrapper) != path:
		parent_wrapper = parent_wrapper.parent()
	return parent_wrapper


def _element_info_text(wrapper, end_path):
	"""
	Returns the information about an element displayed by the tooltip.

	It gives the process and the native window handle of the element, its value, and information that depends on its
	type: the state of a button, the selected text of a combo box, or the text of an edit box.

	:param wrapper: The element
	:param end_path: The end of the path of the element
	:return: The text displayed by the tooltip
	"""
	text = ""
	text += "process id: " + str(wrapper.element_info.process_id) + "\n"
	text += "native window handle: " + str(wrapper.element_info.handle) + "\n"
	try:
		has_get_value = getattr(wrapper, "get_value", None)
		if callable(has_get_value):
			text += "wrapper.get_value(): " + wrapper.get_value()
	except Exception:
		has_get_value = False
	has_legacy_properties = getattr(wrapper, "legacy_properties", None)
	if callable(has_legacy_properties):
		if wrapper.legacy_properties()['Value']:
			if not has_get_value or (has_get_value and wrapper.get_value() != wrapper.legacy_properties()['Value']):
				if text:
					text = text + "\n"
				text = text + " wrapper.legacy_properties()['Value']: " + wrapper.legacy_properties()['Value']
	str_name, str_type, _, _ = get_entry(end_path.split(path_separator)[-1])
	return _add_type_info(text, wrapper, str_name, str_type)


def _add_type_info(text, wrapper, str_name, str_type):
	"""
	Returns the text followed by the information that depends on the type of the element.

	The information that could be read before an error is kept: for example the state of a button whose
	ButtonWrapper cannot be created.

	:param text: The text displayed by the tooltip
	:param wrapper: The element
	:param str_name: The name of the element
	:param str_type: The control type of the element
	:return: The text, followed by the information
	"""
	try:
		if str_type in ("Button", "CheckBox", "RadioButton", "GroupBox"):
			if text:
				text = text + "\n"
			text = text + "wrapper.legacy_properties()['State']: " + str(wrapper.legacy_properties()['State']) + "\n"
			# l'un ou l'autre fonctionne mais pas les 2 en même temps! Pourquoi? Une exception??
			# text = text + "wrapper.get_toggle_state(): " + str(wrapper.get_toggle_state()) + "\n"

			from pywinauto.controls.win32_controls import ButtonWrapper
			text = text + "ButtonWrapper(wrapper).is_checked(): " + str(ButtonWrapper(wrapper).is_checked()) + "\n"
		elif str_type == "ComboBox":
			from pywinauto.controls.win32_controls import ComboBoxWrapper
			if text:
				text = text + "\n"
				text = text + " ComboBoxWrapper(wrapper).selected_text(): " + str(
					ComboBoxWrapper(wrapper).selected_text()) + "\n"
		elif str_type == "Edit":
			if text:
				str_wrapper_text_block = str(wrapper.text_block())
				if str_wrapper_text_block != str_name:
					text = text + "\n"
					text = text + "wrapper.text_block(): " + str_wrapper_text_block + "\n"
	except Exception:
		return text  # the information that could not be read is not displayed
	return text


def _add_rectangle(info_overlay, rectangle, thickness, color):
	"""Draws the outline of the rectangle of an element."""
	info_overlay.add(
		geometry=oaam.Shape.rectangle, x=rectangle.left, y=rectangle.top, width=rectangle.width(),
		height=rectangle.height(), thickness=thickness, color=color)


def _add_box(info_overlay, left, top, height, brush_color):
	"""Draws a box of the tooltip: a rectangle with a black border."""
	info_overlay.add(
		geometry=oaam.Shape.rectangle, x=left, y=top, width=_TOOLTIP_WIDTH, height=height,
		thickness=1, color=(0, 0, 0), brush=oaam.Brush.solid, brush_color=brush_color)


def _add_text(info_overlay, left, top, height, text, text_color, brush_color):
	"""Draws a text of the tooltip, in a box drawn with _add_box."""
	info_overlay.add(
		x=left + 5, y=top, width=_TOOLTIP_WIDTH - 7, height=height, text=text, text_format=_TEXT_FORMAT,
		font_size=16, text_color=text_color, brush=oaam.Brush.solid, brush_color=brush_color,
		geometry=oaam.Shape.rectangle, thickness=0)


def _display_info_tooltip(desktop, info_overlay, screen_width):
	"""
	Displays the tooltip of the element under the mouse cursor, unless it is the element already displayed.

	The tooltip shows, in 3 boxes, the beginning of the path of the element that is common with the previous element,
	the end of its path, and information about the element. The outline of the element is drawn in blue, and the outline
	of the ancestor whose path is the common path is drawn in magenta.

	:param desktop: The pywinauto Desktop
	:param info_overlay: The overlay where the tooltip is drawn
	:param screen_width: The width of the screen: the tooltip is on the side of the screen that is opposite to the mouse
	"""
	global wrapper_old_info_tip
	global common_path_info_tip
	x, y = GetCursorPos()
	wrapper = desktop.from_point(x, y)
	if wrapper is None:
		return
	common_path_info_tip = ''
	if wrapper == wrapper_old_info_tip:
		return
	parent_wrapper = None
	if wrapper_old_info_tip:
		common_path_info_tip = _find_common_path(get_wrapper_path(wrapper), get_wrapper_path(wrapper_old_info_tip))
		if common_path_info_tip:
			parent_wrapper = _ancestor_with_path(wrapper, common_path_info_tip)
	wrapper_old_info_tip = wrapper

	info_overlay.clear_all()
	if parent_wrapper:
		_add_rectangle(info_overlay, parent_wrapper.rectangle(), thickness=1, color=(255, 0, 255))
	_add_rectangle(info_overlay, wrapper.rectangle(), thickness=4, color=(0, 0, 255))

	end_path = get_wrapper_path(wrapper)[len(common_path_info_tip)::]
	tooltip_height = _text_height(common_path_info_tip, 7.8, 25)
	tooltip2_height = _text_height(end_path, 6.8, tooltip_height)
	if x > screen_width / 2:
		info_left = 9
	else:
		info_left = screen_width - _TOOLTIP_WIDTH - 10
	info_top = 150

	_overlay_add_mode_icon(info_overlay, IconSet.hicon_search, info_left, info_top - 54)
	_add_box(info_overlay, info_left, info_top - 2, tooltip_height, (222, 254, 255))
	_add_text(info_overlay, info_left, info_top + 1, tooltip_height - 5, common_path_info_tip, (255, 0, 255), (222, 254, 255))
	_add_box(info_overlay, info_left, info_top - 2 + tooltip_height, tooltip2_height, (180, 254, 255))
	_add_text(info_overlay, info_left, info_top + tooltip_height, tooltip_height - 5, end_path, (0, 0, 255), (180, 254, 255))

	text = _element_info_text(wrapper, end_path)
	tooltip3_height = _text_height(text, 6.8, tooltip_height)
	if text:
		top = info_top + tooltip_height + tooltip2_height
		_add_box(info_overlay, info_left, top - 4, tooltip3_height, (2, 254, 255))
		_add_text(info_overlay, info_left, top, tooltip3_height - 5, text, (0, 0, 0), (2, 254, 255))
	info_overlay.refresh()


TOOLTIP_STOP_TIMEOUT = 5
"""Maximum duration, in seconds, of the wait for the end of the process of the tooltip when it is deleted."""


def task(info_mode_event, info_mode_quit_event):
	"""
	The loop of the process of the tooltip: it displays the tooltip while info_mode_event is set.

	:param info_mode_event: The event set to display the tooltip
	:param info_mode_quit_event: The event set to end the loop
	"""
	global wrapper_old_info_tip
	global common_path_info_tip
	screen_width = GetSystemMetrics(0)
	info_overlay = oaam.Overlay(transparency=0.0)
	desktop = Desktop(backend='uia', allow_magic_lookup=False)
	while not info_mode_quit_event.is_set():
		try:
			if info_mode_event.is_set():
				_display_info_tooltip(desktop, info_overlay, screen_width)
			else:
				info_overlay.clear_all()
				info_overlay.refresh()
			sleep(0.1)
		except BaseException:
			# Any error, even an interruption: the tooltip must keep working until the application stops it
			wrapper_old_info_tip = None
			common_path_info_tip = ''
			info_overlay.clear_all()
			info_overlay.refresh()
			sleep(2.0)
	info_overlay.clear_all()
	info_overlay.refresh()
	# No info_overlay.quit(): the thread of the overlay is a daemon thread, so it ends with this process. Before
	# overlay_arrows_and_more 0.6.0, quit() never returned if the window of the overlay was not created yet, which
	# happens when the tooltip is stopped just after it started (replay of a short file).


class ElementInfoTooltip:
	"""The tooltip that displays information about the element under the mouse cursor, in its own process."""

	def __init__(self):
		"""Starts the process of the tooltip. The tooltip is hidden until show() is called."""
		self.info_mode_quit = Event()
		self.info_mode_quit.clear()
		self.info_mode = Event()
		self.info_mode.clear()
		# daemon process: it is terminated when the application exits, even if the tooltip was not deleted
		self.process = Process(target=task, args=(self.info_mode, self.info_mode_quit), daemon=True)
		self.process.start()

	def show(self):
		"""It displays the tooltip, with the information about the element under the mouse cursor."""
		if not self.info_mode.is_set():
			self.info_mode.set()

	def hide(self):
		"""It hides the tooltip."""
		if self.info_mode.is_set():
			self.info_mode.clear()

	def __del__(self):
		"""Stops the process of the tooltip, or terminates it if it does not stop within TOOLTIP_STOP_TIMEOUT seconds."""
		self.info_mode_quit.set()
		self.process.join(timeout=TOOLTIP_STOP_TIMEOUT)
		if self.process.is_alive():
			self.process.terminate()  # a tooltip that does not stop must not block the application


if __name__ == '__main__':
	element_info_tooltip = ElementInfoTooltip()

	element_info_tooltip.show()
	print('\nMain process blocking', flush=True)
	sleep(88)

	element_info_tooltip.hide()
	print("\n2 seconds without displaying info tip")
	sleep(2)

	del element_info_tooltip
