import overlay_arrows_and_more as oaam
from pywinauto_recorder.core import path_separator, get_wrapper_path, get_entry
from pywinauto import Desktop
from win32api import GetSystemMetrics, GetCursorPos
from time import sleep
from multiprocessing import Process, Event
from pywinauto_recorder.recorder import IconSet, _find_common_path, _overlay_add_mode_icon
import traceback

wrapper_old_info_tip = None
common_path_info_tip = ''


def _display_info_tiptool(desktop, info_overlay, screen_width):
	global wrapper_old_info_tip
	global common_path_info_tip
	tooltip_width = 500
	tooltip_height = 25
	
	x, y = GetCursorPos()
	wrapper = desktop.from_point(x, y)
	
	if wrapper is None:
		return
	
	common_path_info_tip = ''
	parent_wrapper = None
	if wrapper != wrapper_old_info_tip:
		if wrapper_old_info_tip:
			common_path_info_tip = _find_common_path(get_wrapper_path(wrapper), get_wrapper_path(wrapper_old_info_tip))
			if common_path_info_tip:
				parent_wrapper = wrapper.parent()
				while parent_wrapper and get_wrapper_path(parent_wrapper) != common_path_info_tip:
					parent_wrapper = parent_wrapper.parent()
		wrapper_old_info_tip = wrapper
	else:
		return
		
	info_overlay.clear_all()
	
	if parent_wrapper:
		r = parent_wrapper.rectangle()
		info_overlay.add(
			geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(), height=r.height(),
			thickness=1, color=(255, 0, 255))
		
	r = wrapper.rectangle()
	info_overlay.add(
		geometry=oaam.Shape.rectangle, x=r.left, y=r.top, width=r.width(), height=r.height(),
		thickness=4, color=(0, 0, 255))
	
	end_path = get_wrapper_path(wrapper)[len(common_path_info_tip)::]

	text = ''
	text_width = 0
	for c in common_path_info_tip:
		if c == '\n':
			text_width = 0
			tooltip_height = tooltip_height + 16
		else:
			text_width = text_width + 7.8
		text = text + c
		if text_width > tooltip_width:
			text = text + '\n'
			text_width = 0
			tooltip_height = tooltip_height + 16
	text = common_path_info_tip # TEST
	
	tooltip2_height = tooltip_height
	text2 = ''
	text_width = 0
	for c in end_path:
		if c == '\n':
			text_width = 0
			tooltip2_height = tooltip2_height + 16
		else:
			text_width = text_width + 6.8
		text2 = text2 + c
		if text_width > tooltip_width:
			text2 = text2 + '\n'
			text_width = 0
			tooltip2_height = tooltip2_height + 16
	text2 = end_path # TEST

	if x > screen_width / 2:
		info_left = 9
	else:
		info_left = screen_width - tooltip_width - 10
	info_top = 150

	_overlay_add_mode_icon(info_overlay, IconSet.hicon_search, info_left, info_top - 54)
	
	info_overlay.add(
		geometry=oaam.Shape.rectangle, x=info_left, y=info_top - 2, width=tooltip_width, height=tooltip_height,
		thickness=1, color=(0, 0, 0), brush=oaam.Brush.solid, brush_color=(222, 254, 255))
	info_overlay.add(
		x=info_left + 5, y=info_top + 1, width=tooltip_width - 7,
		height=tooltip_height - 5,
		text=text,
		text_format="win32con.DT_LEFT|win32con.DT_WORDBREAK|win32con.DT_EXPANDTABS|win32con.DT_EDITCONTROL|win32con.DT_NOCLIP",
		font_size=16, text_color=(255, 0, 255), brush=oaam.Brush.solid, brush_color=(222, 254, 255),
		geometry=oaam.Shape.rectangle, thickness=0
	)
	info_overlay.add(
		geometry=oaam.Shape.rectangle, x=info_left, y=info_top - 2 + tooltip_height,
		width=tooltip_width, height=tooltip2_height,
		thickness=1, color=(0, 0, 0), brush=oaam.Brush.solid, brush_color=(180, 254, 255))
	info_overlay.add(
		x=info_left + 5, y=info_top + tooltip_height, width=tooltip_width - 7,
		height=tooltip_height - 5,
		text=text2,
		text_format="win32con.DT_LEFT | win32con.DT_WORDBREAK | win32con.DT_EXPANDTABS | win32con.DT_EDITCONTROL|win32con.DT_NOCLIP",
		font_size=16, text_color=(0, 0, 255), brush=oaam.Brush.solid, brush_color=(180, 254, 255),
		geometry=oaam.Shape.rectangle, thickness=0
	)
	
	text = ""
	try:
		has_get_value = getattr(wrapper, "get_value", None)
		if callable(has_get_value):
			text = "wrapper.get_value(): " + wrapper.get_value()
	except:
		has_get_value = False
	has_legacy_properties = getattr(wrapper, "legacy_properties", None)
	if callable(has_legacy_properties):
		if wrapper.legacy_properties()['Value']:
			if not has_get_value or (has_get_value and wrapper.get_value() != wrapper.legacy_properties()['Value']):
				if text:
					text = text + "\n"
				text = text + " wrapper.legacy_properties()['Value']: " + wrapper.legacy_properties()['Value']
	str_name, str_type, _, _ = get_entry(end_path.split(path_separator)[-1])
	try:
		if str_type in ("Button", "CheckBox", "RadioButton", "GroupBox"):
			if text:
				text = text + "\n"
			text = text + "wrapper.legacy_properties()['State']: " + str(wrapper.legacy_properties()['State']) + "\n"
			# l'un ou l'autre fonctionne mais pas les 2 en même temps! Pourquoi? Une exception??
			# text = text + "wrapper.get_toggle_state(): " + str(wrapper.get_toggle_state()) + "\n"
			
			from pywinauto.controls.win32_controls import ButtonWrapper
			text = text + "ButtonWrapper(wrapper).is_checked(): " + str(ButtonWrapper(wrapper).is_checked()) + "\n"
		elif str_type in ("ComboBox"):
			from pywinauto.controls.win32_controls import ComboBoxWrapper
			if text:
				text = text + "\n"
				text = text + " ComboBoxWrapper(wrapper).selected_text(): " + str(
					ComboBoxWrapper(wrapper).selected_text()) + "\n"
		elif str_type in ("Edit"):
			if text:
				str_wrapper_text_block = str(wrapper.text_block())
				if str_wrapper_text_block != str_name:
					text = text + "\n"
					text = text + "wrapper.text_block(): " + str_wrapper_text_block + "\n"
	except:
		pass
	tooltip3_height = tooltip_height
	text3 = ''
	text_width = 0
	for c in text:
		if c == '\n':
			text_width = 0
			tooltip3_height = tooltip3_height + 16
		else:
			text_width = text_width + 6.8
		text3 = text3 + c
		if text_width > tooltip_width:
			text3 = text3 + '\n'
			text_width = 0
			tooltip3_height = tooltip3_height + 16
	text3 = text
	if text3:
		info_overlay.add(
			geometry=oaam.Shape.rectangle, x=info_left, y=info_top - 4 + tooltip_height + tooltip2_height,
			width=tooltip_width,
			height=tooltip3_height, thickness=1, color=(0, 0, 0), brush=oaam.Brush.solid, brush_color=(2, 254, 255))
		info_overlay.add(
			x=info_left + 5, y=info_top + tooltip_height + tooltip2_height, width=tooltip_width - 7,
			height=tooltip3_height - 5,
			text=text3,
			text_format="win32con.DT_LEFT|win32con.DT_WORDBREAK|win32con.DT_EXPANDTABS|win32con.DT_EDITCONTROL|win32con.DT_NOCLIP",
			font_size=16, text_color=(0, 0, 0), brush=oaam.Brush.solid, brush_color=(2, 254, 255),
			geometry=oaam.Shape.rectangle, thickness=0
		)
	info_overlay.refresh()


def task(info_mode_event, info_mode_quit_event):
	global wrapper_old_info_tip
	global common_path_info_tip
	screen_width = GetSystemMetrics(0)
	info_overlay = oaam.Overlay(transparency=0.0)
	desktop = Desktop(backend='uia', allow_magic_lookup=False)
	while not info_mode_quit_event.is_set():
		try:
			if info_mode_event.is_set():
					_display_info_tiptool(desktop, info_overlay, screen_width)
			else:
				info_overlay.clear_all()
				info_overlay.refresh()
			sleep(0.1)
		except:
			wrapper_old_info_tip = None
			common_path_info_tip = ''
			info_overlay.clear_all()
			info_overlay.refresh()
			sleep(2.0)
	info_overlay.clear_all()
	info_overlay.refresh()
	info_overlay.quit()


class ElementInfoTooltip:
	def __init__(self):
		self.info_mode_quit = Event()
		self.info_mode_quit.clear()
		self.info_mode = Event()
		self.info_mode.clear()
		self.process = Process(target=task, args=(self.info_mode, self.info_mode_quit))
		self.process.start()
	
	def show(self):
		if not self.info_mode.is_set():
			self.info_mode.set()
			print("def show(self):")
	
	def hide(self):
		if self.info_mode.is_set():
			self.info_mode.clear()
			print("def hide(self):")
	
	def __del__(self):
		self.info_mode_quit.set()
		self.process.join()


if __name__ == '__main__':
	element_info_tooltip = ElementInfoTooltip()
	
	element_info_tooltip.show()
	print('\nMain process blocking', flush=True)
	sleep(88)
	
	element_info_tooltip.hide()
	print("\n2 seconds without displaying info tip")
	sleep(2)
	
	del element_info_tooltip
