import builtins
import io
import pathlib
import sys
import time
import warnings
import pytest
from pywinauto_recorder.core import get_entry_list, get_entry, match_entry_list, get_sorted_region
from pywinauto_recorder import core, player
from pywinauto_recorder.player import UIPath, PlayerSettings, FailedSearch, click


################################################################################################
#                                   Tests of low level functions                               #
################################################################################################

def test_get_entry_list():
	"""Tests get_entry_list(element_path)."""
	element_path = "D:\\Name||Window->||Pane->Property||Group->"
	element_path = element_path + "Label:||Text#[Name||Window->||Pane->Property||Group->Label:||Text,0]%(0,0)"
	entry_list = get_entry_list(element_path)
	assert entry_list[0] == 'D:\\Name||Window'
	assert entry_list[1] == '||Pane'
	assert entry_list[2] == 'Property||Group'
	assert entry_list[3] == 'Label:||Text#[Name||Window->||Pane->Property||Group->Label:||Text,0]%(0,0)'


def test_get_entry_list_with_asterisk():
	"""Tests get_entry_list(element_path) having element_path with some asterisks."""
	element_path = "Name||Window->*->||Group->Property||%(0,0)"
	entry_list2 = get_entry_list(element_path)
	assert entry_list2[1] == '*'
	assert entry_list2[2] == '||Group'
	assert entry_list2[3] == 'Property||%(0,0)'


def test_get_entry_elements():
	"""Tests get_entry(entry)."""
	str_name_with_colon = 'Name:'
	str_name_without_colon = 'Name'
	str_y_x_with_colon = 'y_name:||y_type'
	str_y_x_without_colon = 'y_name||y_type'
	test_data = [
		('Name:||Type#[0,0]%(2,-24)',      (str_name_with_colon, 'Type', [0,0], (2,-24))),
		('||Type#[0,0]%(2,-24)',           ('', 'Type', [0,0], (2,-24))),
		('Name:||#[0,0]%(2,-24)',          (str_name_with_colon, None, [0,0], (2,-24))),
		('Name:||Type#[0,0]',              (str_name_with_colon, 'Type', [0,0], None)),
		('||Type#[0,0]',                   ('', 'Type', [0,0], None)),
		('Name:||#[0,0]',                  (str_name_with_colon, None, [0,0], None)),
		('Name:||Type',                    (str_name_with_colon, 'Type', None, None)),
		('||Type',                         ('', 'Type', None, None)),
		('Name:||',                        (str_name_with_colon, None, None, None)),
		('Name||Type',                     (str_name_without_colon, 'Type', None, None)),
		('||Type',                         ('', 'Type', None, None)),
		('Name||',                         (str_name_without_colon, None, None, None)),
		('Name||Type#[y_name:||y_type,0]', (str_name_without_colon, 'Type', [str_y_x_with_colon, 0], None)),
		('||Type#[y_name:||y_type,0]',     ('', 'Type', [str_y_x_with_colon, 0], None)),
		('Name||#[y_name:||y_type,0]',     (str_name_without_colon, None, [str_y_x_with_colon, 0], None)),
		('Name||Type#[y_name||y_type,0]',  (str_name_without_colon, 'Type', [str_y_x_without_colon, 0], None)),
		('||Type#[y_name||y_type,0]',      ('', 'Type', [str_y_x_without_colon, 0], None)),
		('Name||#[y_name||y_type,0]',      (str_name_without_colon, None, [str_y_x_without_colon, 0], None))
	]
	for data in test_data:
		str_name, str_type, y_x, dx_dy = get_entry(data[0])
		assert str_name == data[1][0]
		assert str_type == data[1][1]
		assert y_x == data[1][2]
		assert dx_dy == data[1][3]


def test_match_entry_list():
	"""Tests if the entry_list matches the template_list."""
	entry = "wName||Window->pName||Pane->||Pane->materialFileLineEdit||Pane"
	template = "wName||Window->pName||Pane->*->RegEx: .*||Pane"
	assert match_entry_list(get_entry_list(entry), get_entry_list(template))

	template_list = get_entry_list("wName||Window->*->Tab1||Tab->A||Group->B||Group->*->Login||Button")
	entry_path_set = [
		("wName||Window->Tab1||Tab->A||Group->B||Group->Login||Button", True),
		("wName||Window->Tab1||Tab->A||Group->Login||Button", False),
		("wName||Window->bla||bla->Tab1||Tab->A||Group->B||Group->bla||bla->Login||Button", True),
		("wName||Window->bla||bla->Tab1||Tab->bla||bla->A||Group->B||Group->bla||bla->Login||Button", False),
		("wName||Window->bla||bla->bla||bla->Tab1||Tab->A||Group->B||Group->bla||bla->bla||bla->Login||Button", True),
	]
	for entry_path in entry_path_set:
		assert match_entry_list(get_entry_list(entry_path[0]), template_list) == entry_path[1]



################################################################################################
#                                       Tests of UIPath                                        #
################################################################################################

def test_uipath_nested():
	"""Tests that nested UIPath objects build the full path and restore it when they exit."""
	assert UIPath.get_full_path() == ""
	with UIPath("W1||Window"):
		assert UIPath.get_full_path() == "W1||Window"
		with UIPath("A||Pane"):
			assert UIPath.get_full_path("OK||Button") == "W1||Window->A||Pane->OK||Button"
			with UIPath("B||Group"):
				assert UIPath.get_full_path("OK||Button") == "W1||Window->A||Pane->B||Group->OK||Button"
			assert UIPath.get_full_path("OK||Button") == "W1||Window->A||Pane->OK||Button"
		assert UIPath.get_full_path("OK||Button") == "W1||Window->OK||Button"
	assert UIPath.get_full_path() == ""


def test_uipath_absolute_path_with_nested_uipath():
	"""
	Tests that a UIPath nested in a UIPath having absolute_path=True does not restore
	the path saved by the absolute UIPath when it exits.
	"""
	with UIPath("W1||Window"):
		with UIPath("W2||Window", absolute_path=True):
			assert UIPath.get_full_path("OK||Button") == "W2||Window->OK||Button"
			with UIPath("child||Pane"):
				assert UIPath.get_full_path("OK||Button") == "W2||Window->child||Pane->OK||Button"
			# The path of the absolute UIPath must still be the current one
			assert UIPath.get_full_path("Cancel||Button") == "W2||Window->Cancel||Button"
			with UIPath("other||Pane"):
				assert UIPath.get_full_path("OK||Button") == "W2||Window->other||Pane->OK||Button"
			assert UIPath.get_full_path("Cancel||Button") == "W2||Window->Cancel||Button"
		# The path of the outer UIPath must be restored only when the absolute UIPath exits
		assert UIPath.get_full_path("X||Button") == "W1||Window->X||Button"
	assert UIPath.get_full_path() == ""


def test_uipath_absolute_path_at_top_level():
	"""Tests a UIPath having absolute_path=True that is not nested in another UIPath."""
	with UIPath("W2||Window", absolute_path=True):
		with UIPath("child||Pane"):
			assert UIPath.get_full_path("OK||Button") == "W2||Window->child||Pane->OK||Button"
		assert UIPath.get_full_path("Cancel||Button") == "W2||Window->Cancel||Button"
	assert UIPath.get_full_path() == ""


def test_uipath_nested_absolute_paths():
	"""Tests two nested UIPath objects having absolute_path=True."""
	with UIPath("W1||Window"):
		with UIPath("W2||Window", absolute_path=True):
			with UIPath("W3||Window", absolute_path=True):
				with UIPath("child||Pane"):
					assert UIPath.get_full_path("OK||Button") == "W3||Window->child||Pane->OK||Button"
				assert UIPath.get_full_path("OK||Button") == "W3||Window->OK||Button"
			assert UIPath.get_full_path("OK||Button") == "W2||Window->OK||Button"
		assert UIPath.get_full_path("OK||Button") == "W1||Window->OK||Button"
	assert UIPath.get_full_path() == ""


def test_uipath_restored_after_exception():
	"""Tests that the path is restored when an exception is raised in nested UIPath blocks."""
	try:
		with UIPath("W1||Window"):
			with UIPath("W2||Window", absolute_path=True):
				with UIPath("child||Pane"):
					raise RuntimeError("error raised in the nested UIPath")
	except RuntimeError:
		pass
	assert UIPath.get_full_path() == ""
	assert UIPath._stack == []


def test_uipath_regex_title_is_deprecated():
	"""
	Tests that the parameter 'regex_title' of UIPath, ignored since Pywinauto Recorder 0.6.0,
	is still accepted but emits a warning, and that it does not change the path.
	"""
	with pytest.warns(FutureWarning, match="regex_title"):
		ui_path = UIPath(".* - Notepad||Window", regex_title=True)
	with ui_path:
		assert UIPath.get_full_path("Text Editor||Edit") == ".* - Notepad||Window->Text Editor||Edit"
	assert UIPath.get_full_path() == ""


def test_uipath_without_regex_title_does_not_warn():
	"""Tests that no warning is emitted when the parameter 'regex_title' is not used."""
	with warnings.catch_warnings():
		warnings.simplefilter("error")
		with UIPath("RegEx: .* - Notepad||Window"):
			with UIPath("Text Editor||Edit", regex_title=False):
				assert UIPath.get_full_path() == "RegEx: .* - Notepad||Window->Text Editor||Edit"
	assert UIPath.get_full_path() == ""


################################################################################################
#                                        Tests of click                                        #
################################################################################################

class _FakeWrapper:
	"""Stands for the Pywinauto wrapper returned by 'move' in the tests of 'click'."""


@pytest.mark.parametrize("use_cache", [True, False])
def test_click_restores_use_cache_when_the_search_fails(monkeypatch, use_cache):
	"""
	Tests that 'click' restores PlayerSettings.use_cache when the element is not found.
	'click' disables the cache while it searches the element and waits until it is ready.
	"""
	def move_raising_failed_search(*args, **kwargs):
		assert PlayerSettings.use_cache is False  # the cache is disabled during the search
		raise FailedSearch("No element found")

	monkeypatch.setattr(PlayerSettings, "use_cache", use_cache)
	monkeypatch.setattr(player, "move", move_raising_failed_search)
	with pytest.raises(FailedSearch):
		click("Not found||Window->OK||Button")
	assert PlayerSettings.use_cache is use_cache


@pytest.mark.parametrize("use_cache", [True, False])
def test_click_restores_use_cache_when_the_element_is_never_ready(monkeypatch, use_cache):
	"""Tests that 'click' restores PlayerSettings.use_cache when waiting for the element to be ready times out."""
	def wait_raising_timeout_error(*args, **kwargs):
		assert PlayerSettings.use_cache is False
		raise TimeoutError("Time out!")

	monkeypatch.setattr(PlayerSettings, "use_cache", use_cache)
	monkeypatch.setattr(player, "move", lambda *args, **kwargs: _FakeWrapper())
	monkeypatch.setattr(player, "wait_is_ready_try1", wait_raising_timeout_error)
	with pytest.raises(TimeoutError):
		click("Window||Window->Disabled||Button")
	assert PlayerSettings.use_cache is use_cache


@pytest.mark.parametrize("use_cache", [True, False])
def test_click_restores_use_cache_when_it_succeeds(monkeypatch, use_cache):
	"""Tests that 'click' restores PlayerSettings.use_cache and returns the clicked element when it succeeds."""
	wrapper = _FakeWrapper()
	clicks = []
	monkeypatch.setattr(PlayerSettings, "use_cache", use_cache)
	monkeypatch.setattr(player, "move", lambda *args, **kwargs: wrapper)
	monkeypatch.setattr(player, "wait_is_ready_try1", lambda *args, **kwargs: None)
	monkeypatch.setattr(player, "_win32api_mouse_click", lambda button, click_count: clicks.append((button, click_count)))
	assert click("Window||Window->OK||Button") is wrapper
	assert clicks == [(player.ButtonLocation.left, 1)]
	assert PlayerSettings.use_cache is use_cache


################################################################################################
#                                        Tests of _find                                        #
################################################################################################

class _FakeRectangle:
	"""Stands for the rectangle of a Pywinauto wrapper."""
	def __init__(self, left, top, width, height):
		self.left, self.top, self.right, self.bottom = left, top, left + width, top + height

	def width(self):
		return self.right - self.left

	def height(self):
		return self.bottom - self.top

	def mid_point(self):
		return (self.left + self.right) // 2, (self.top + self.bottom) // 2


class _FakeElement:
	"""Stands for a Pywinauto wrapper returned by 'find_elements' in the tests of '_find'."""
	def __init__(self, name, left=0, top=0, width=50, height=20, visible=True, parent=None):
		self.name = name
		self._rectangle = _FakeRectangle(left, top, width, height)
		self._visible = visible
		self._parent = parent

	def parent(self):
		return self._parent

	def rectangle(self):
		return self._rectangle

	def is_visible(self):
		return self._visible


class _FakeFindElements:
	"""
	Replaces 'find_elements': each call returns the next list of 'results', the last one is repeated.
	"""
	def __init__(self, *results):
		self.results = list(results)
		self.call_count = 0

	def __call__(self, full_element_path=None):
		result = self.results[min(self.call_count, len(self.results) - 1)]
		self.call_count += 1
		return list(result)


@pytest.fixture
def fake_find_elements(monkeypatch):
	"""Returns a function that replaces 'find_elements' in the player by a '_FakeFindElements'."""
	def install(*results, search_again_pause=0.05):
		fake = _FakeFindElements(*results)
		monkeypatch.setattr(player, "find_elements", fake)
		monkeypatch.setattr(player, "_SEARCH_AGAIN_PAUSE", search_again_pause)  # short pause: fast tests
		monkeypatch.setattr(player, "get_wrapper_path", lambda wrapper: wrapper.name)
		return fake
	return install


def test_find_returns_a_unique_element_with_a_single_search(fake_find_elements):
	"""Tests that '_find' returns at once the element when only one element matches the path."""
	ok = _FakeElement("Window||Window->OK||Button")
	fake = fake_find_elements([ok])
	assert player._find("Window||Window->OK||Button", timeout=2) is ok
	assert fake.call_count == 1


def test_find_searches_again_when_several_elements_match(fake_find_elements):
	"""
	Tests that '_find' searches again when several elements match a path without index:
	the element is returned as soon as it is the only one that matches the path.
	"""
	ok = _FakeElement("Window||Window->OK||Button")
	closing = _FakeElement("Closing window||Window->OK||Button")
	fake = fake_find_elements([ok, closing], [ok, closing], [ok])
	assert player._find("RegEx: .*||Window->OK||Button", timeout=5) is ok
	assert fake.call_count == 3


def test_find_uses_the_result_of_the_last_search(fake_find_elements):
	"""Tests that '_find' returns the element when it becomes the only one at the last search, done at the timeout."""
	ok = _FakeElement("Window||Window->OK||Button")
	closing = _FakeElement("Closing window||Window->OK||Button")
	# Searches at about 0 s, 0.2 s, 0.4 s and 0.5 s (the timeout)
	fake_find_elements([ok, closing], [ok, closing], [ok, closing], [ok], search_again_pause=0.2)
	assert player._find("RegEx: .*||Window->OK||Button", timeout=0.5) is ok


def test_find_fails_when_several_elements_still_match_after_the_timeout(fake_find_elements):
	"""Tests that '_find' raises FailedSearch when several elements still match the path after the timeout."""
	ok_1 = _FakeElement("Window 1||Window->OK||Button")
	ok_2 = _FakeElement("Window 2||Window->OK||Button")
	fake = fake_find_elements([ok_1, ok_2])
	with pytest.raises(FailedSearch) as exc_info:
		player._find("RegEx: .*||Window->OK||Button", timeout=0.5)
	message = str(exc_info.value)
	assert "2 undiscriminated elements" in message
	assert ok_1.name in message and ok_2.name in message
	assert fake.call_count > 1  # the search was done again while waiting


def test_find_pauses_between_two_searches_when_several_elements_match(fake_find_elements):
	"""
	Tests that '_find' pauses between two searches when several elements match the path,
	and that the default pause is 1 second.
	"""
	assert player._SEARCH_AGAIN_PAUSE == 1.0
	ok_1 = _FakeElement("Window 1||Window->OK||Button")
	ok_2 = _FakeElement("Window 2||Window->OK||Button")
	fake = fake_find_elements([ok_1, ok_2], search_again_pause=0.3)
	t0 = time.time()
	with pytest.raises(FailedSearch):
		player._find("RegEx: .*||Window->OK||Button", timeout=1.0)
	duration = time.time() - t0
	# Searches at about 0 s, 0.3 s, 0.6 s, 0.9 s and 1.0 s (the last pause is shortened to respect the timeout)
	assert 2 <= fake.call_count <= 6
	assert 0.9 <= duration < 2.0


def test_find_fails_when_the_elements_disappear_after_matching(fake_find_elements):
	"""Tests that '_find' raises FailedSearch when the elements that matched the path are no longer found."""
	ok_1 = _FakeElement("Window 1||Window->OK||Button")
	ok_2 = _FakeElement("Window 2||Window->OK||Button")
	fake_find_elements([ok_1, ok_2], [])
	with pytest.raises(FailedSearch, match="No element found"):
		player._find("RegEx: .*||Window->OK||Button", timeout=0.3)


def _two_rows_of_buttons():
	"""Returns 4 fake buttons arranged in 2 rows of 2 buttons: [[a, b], [c, d]]."""
	a = _FakeElement("Window||Window->OK||Button", left=0, top=0)
	b = _FakeElement("Window||Window->OK||Button", left=100, top=0)
	c = _FakeElement("Window||Window->OK||Button", left=0, top=40)
	d = _FakeElement("Window||Window->OK||Button", left=100, top=40)
	return a, b, c, d


def test_find_with_index(fake_find_elements):
	"""Tests that '_find' returns the element at the row and column of the [] operator."""
	a, b, c, d = _two_rows_of_buttons()
	fake_find_elements([d, b, c, a])
	assert player._find("Window||Window->OK||Button#[0,0]", timeout=2) is a
	assert player._find("Window||Window->OK||Button#[0,1]", timeout=2) is b
	assert player._find("Window||Window->OK||Button#[1,0]", timeout=2) is c
	assert player._find("Window||Window->OK||Button#[1,1]", timeout=2) is d


@pytest.mark.parametrize("index", ["#[2,0]", "#[0,2]", "#[7,7]"])
def test_find_fails_when_the_index_is_out_of_range(fake_find_elements, index):
	"""
	Tests that '_find' raises FailedSearch, not IndexError, when the row or the column
	of the [] operator is out of range after the timeout.
	"""
	fake = fake_find_elements(list(_two_rows_of_buttons()))
	with pytest.raises(FailedSearch) as exc_info:
		player._find("Window||Window->OK||Button" + index, timeout=0.3)
	message = str(exc_info.value)
	assert "out of range" in message
	assert "Window||Window->OK||Button" + index in message
	assert "2 row(s)" in message
	assert fake.call_count > 1  # the search was done again while waiting


def test_find_waits_until_the_index_is_in_range(fake_find_elements):
	"""Tests that '_find' searches again when the index is out of range (a list that is still being filled)."""
	a, b, c, d = _two_rows_of_buttons()
	fake = fake_find_elements([a, b], [a, b], [a, b, c, d])
	assert player._find("Window||Window->OK||Button#[1,0]", timeout=5) is c
	assert fake.call_count == 3


def test_find_fails_when_no_element_is_visible_with_an_index(fake_find_elements):
	"""Tests that '_find' raises FailedSearch when the [] operator is used and no found element is visible."""
	hidden_1 = _FakeElement("Window||Window->OK||Button", left=0, visible=False)
	hidden_2 = _FakeElement("Window||Window->OK||Button", left=100, visible=False)
	fake_find_elements([hidden_1, hidden_2])
	with pytest.raises(FailedSearch, match="out of range"):
		player._find("Window||Window->OK||Button#[0,0]", timeout=0.3)


def test_find_fails_when_the_column_is_out_of_range_with_a_reference_element(monkeypatch):
	"""
	Tests that '_find' raises FailedSearch, not IndexError, when the column is out of range
	and the row is given by a reference element: "...#[UIPath of the reference element,column]".
	"""
	a, b, c, d = _two_rows_of_buttons()
	label = _FakeElement("Label||Text", left=200, top=35, height=30)  # on the row of c and d

	def find_elements(full_element_path=None):
		return [label] if full_element_path == "Label||Text" else [a, b, c, d]

	monkeypatch.setattr(player, "find_elements", find_elements)
	monkeypatch.setattr(player, "get_wrapper_path", lambda wrapper: wrapper.name)
	monkeypatch.setattr(player, "_SEARCH_AGAIN_PAUSE", 0.05)
	assert player._find("Window||Window->OK||Button#[Label||Text,1]", timeout=2) is d
	with pytest.raises(FailedSearch, match="out of range"):
		player._find("Window||Window->OK||Button#[Label||Text,5]", timeout=0.3)


def test_get_sorted_region_without_visible_element():
	"""Tests that 'get_sorted_region' returns no row when there is no element or no visible element."""
	assert get_sorted_region([]) == (0, 0, [])
	assert get_sorted_region([_FakeElement("OK||Button", visible=False)]) == (0, 0, [])


@pytest.mark.parametrize("encoding, expected", [("utf-8", "🔎"), ("cp1252", "?")])
def test_find_prints_its_progress_whatever_the_encoding_of_stdout(monkeypatch, fake_find_elements, encoding, expected):
	"""
	Tests that '_find' does not raise UnicodeEncodeError when the encoding of stdout cannot display
	the progress symbols (output redirected to a file or a pipe on Windows): ASCII symbols are printed instead.
	"""
	ok = _FakeElement("Window||Window->OK||Button")
	fake_find_elements([ok])
	stdout = io.TextIOWrapper(io.BytesIO(), encoding=encoding)
	monkeypatch.setattr(sys, "stdout", stdout)
	assert player._find("Window||Window->OK||Button", timeout=2) is ok
	stdout.flush()
	assert stdout.buffer.getvalue().decode(encoding) == expected


def test_find_works_without_stdout(monkeypatch, fake_find_elements):
	"""Tests that '_find' works when there is no stdout (sys.stdout is None with pythonw.exe or a windowed application)."""
	ok = _FakeElement("Window||Window->OK||Button")
	fake_find_elements([ok])
	monkeypatch.setattr(sys, "stdout", None)
	assert player._find("Window||Window->OK||Button", timeout=2) is ok


################################################################################################
#                           Tests of the other functions of the player                         #
################################################################################################

@pytest.fixture
def no_sleep(monkeypatch):
	"""Replaces 'time.sleep' by a function that does not wait."""
	monkeypatch.setattr(time, "sleep", lambda seconds: None)


@pytest.fixture
def native_window_handle(monkeypatch):
	"""Sets the native window handle used by the searches to 1234, and restores it after the test."""
	monkeypatch.setattr(core, "native_window_handle", 1234)
	return 1234


def _assert_message_is_a_sentence(exception):
	"""Asserts that the exception was built with one string, not with several arguments displayed as a tuple."""
	assert len(exception.args) == 1
	assert isinstance(exception.args[0], str)
	assert not str(exception).startswith("(")


# ---------------------------------------- menu_click ----------------------------------------

class _FakeMenu:
	"""Replaces 'find_all', 'move', 'click' and the mouse cursor position in the tests of 'menu_click'."""
	def __init__(self, monkeypatch, items_by_name, cursor_pos=(100, 100)):
		self.items_by_name = items_by_name
		self.clicked_positions = []
		monkeypatch.setattr(player, "win32api_GetCursorPos", lambda: cursor_pos)
		monkeypatch.setattr(player, "find_all", self.find_all)
		monkeypatch.setattr(player, "move", lambda *args, **kwargs: None)
		monkeypatch.setattr(player, "click", lambda position, **kwargs: self.clicked_positions.append(position))

	def find_all(self, element_path=None, timeout=None):
		for name, items in self.items_by_name.items():
			if name + "||MenuItem" in element_path:
				return list(items)
		return []


def _menu_item(name, mid_x, mid_y):
	"""Returns a fake menu item of 50 x 20 pixels whose middle is (mid_x, mid_y)."""
	menu = _FakeElement("||Menu", left=0, top=0, width=1000, height=1000)
	return _FakeElement(name + "||MenuItem", left=mid_x - 25, top=mid_y - 10, width=50, height=20, parent=menu)


def test_menu_click_clicks_on_the_nearest_item(monkeypatch, no_sleep, native_window_handle):
	"""
	Tests that 'menu_click' clicks on the item that is the nearest to the mouse cursor
	when several items have the same name: the vertical distance must be taken into account.
	"""
	far_below = _menu_item("Open", 100, 500)  # same x as the mouse cursor, but 400 pixels below
	near = _menu_item("Open", 150, 110)
	menu = _FakeMenu(monkeypatch, {"Open": [far_below, near]}, cursor_pos=(100, 100))
	assert player.menu_click("Open") is near
	assert menu.clicked_positions == [(150, 110)]


def test_menu_click_restores_the_path_and_the_window_handle(monkeypatch, no_sleep, native_window_handle):
	"""Tests that 'menu_click' clicks on each item and restores the current UIPath and the native window handle."""
	file_item, new_item = _menu_item("File", 20, 10), _menu_item("New", 30, 40)
	menu = _FakeMenu(monkeypatch, {"File": [file_item], "New": [new_item]})
	with UIPath("Window||Window"):
		assert player.menu_click("File->New") is file_item
		assert UIPath.get_full_path() == "Window||Window"
	assert menu.clicked_positions == [(20, 10), (30, 40)]
	assert core.get_native_window_handle() == native_window_handle


def test_menu_click_fails_when_a_menu_item_is_not_found(monkeypatch, no_sleep, native_window_handle):
	"""
	Tests that 'menu_click' raises FailedSearch, not IndexError, when a menu item is not found,
	and that it restores the current UIPath and the native window handle.
	"""
	_FakeMenu(monkeypatch, {"File": [_menu_item("File", 20, 10)]})
	with UIPath("Window||Window"):
		with pytest.raises(FailedSearch, match="New") as exc_info:
			player.menu_click("File->New")
		assert UIPath.get_full_path() == "Window||Window"
	_assert_message_is_a_sentence(exc_info.value)
	assert core.get_native_window_handle() == native_window_handle


# --------------------------------------- drag_and_drop ---------------------------------------

def test_drag_and_drop_releases_the_button_when_the_destination_is_not_found(monkeypatch):
	"""Tests that 'drag_and_drop' releases the mouse button when the destination element is not found."""
	mouse_events = []

	def move(element_path, **kwargs):
		if element_path == "Destination||Pane":
			raise FailedSearch("No element found")

	monkeypatch.setattr(player, "move", move)
	monkeypatch.setattr(player, "win32api_mouse_event", lambda event, *args: mouse_events.append(event))
	with pytest.raises(FailedSearch):
		player.drag_and_drop("Source||Pane", "Destination||Pane")
	assert mouse_events == [player.MOUSEEVENTF_LEFTDOWN, player.MOUSEEVENTF_LEFTUP]


def test_drag_and_drop_presses_and_releases_the_button(monkeypatch):
	"""Tests that 'drag_and_drop' presses the button on the source and releases it on the destination."""
	mouse_events = []
	destination = _FakeElement("Destination||Pane")
	monkeypatch.setattr(player, "move", lambda element_path, **kwargs: destination if element_path == "Destination||Pane" else None)
	monkeypatch.setattr(player, "win32api_mouse_event", lambda event, *args: mouse_events.append(event))
	assert player.drag_and_drop("Source||Pane", "Destination||Pane", button=player.ButtonLocation.right) is destination
	assert mouse_events == [player.MOUSEEVENTF_RIGHTDOWN, player.MOUSEEVENTF_RIGHTUP]


# ---------------------------------------- select_file ----------------------------------------

class _FakeFileDialog:
	"""Replaces 'find', the clicks and 'send_keys' in the tests of 'select_file'."""
	def __init__(self, monkeypatch, current_folder):
		self.current_folder = current_folder
		self.sent_keys = []
		monkeypatch.setattr(player, "find", lambda *args, **kwargs: self)
		monkeypatch.setattr(player, "click", lambda *args, **kwargs: None)
		monkeypatch.setattr(player, "double_left_click", lambda *args, **kwargs: None)
		monkeypatch.setattr(player, "send_keys", lambda str_keys, **kwargs: self.sent_keys.append(str_keys))

	def set_focus(self):
		pass

	def get_value(self):
		return self.current_folder


def test_select_file_does_not_type_the_folder_when_it_is_the_current_one(monkeypatch):
	"""Tests that 'select_file' does not type the folder when the file dialog is already in the folder of the file."""
	dialog = _FakeFileDialog(monkeypatch, current_folder="C:/Users/me/Documents")
	player.select_file("Open||Window", "C:/Users/me/Documents/file.txt")
	assert dialog.sent_keys == ["{ENTER}", "file.txt{ENTER}"]


@pytest.mark.skipif(sys.platform != "win32", reason="Windows paths")
def test_select_file_compares_the_folders_as_windows_paths(monkeypatch):
	"""Tests that 'select_file' ignores the case and the kind of slashes when it compares the folders."""
	dialog = _FakeFileDialog(monkeypatch, current_folder="c:\\users\\me\\documents")
	player.select_file("Open||Window", "C:/Users/me/Documents/file.txt")
	assert dialog.sent_keys == ["{ENTER}", "file.txt{ENTER}"]


def test_select_file_types_the_folder_when_it_is_not_the_current_one(monkeypatch):
	"""Tests that 'select_file' types the folder when the file dialog is in another folder."""
	dialog = _FakeFileDialog(monkeypatch, current_folder="C:/Users/me/Pictures")
	player.select_file("Open||Window", "C:/Users/me/Documents/file.txt")
	assert dialog.sent_keys == [str(pathlib.Path("C:/Users/me/Documents")), "{ENTER}", "file.txt{ENTER}"]


def test_select_file_types_the_folder_when_slow_path_typing_is_forced(monkeypatch):
	"""Tests that 'select_file' always types the folder when 'force_slow_path_typing' is True."""
	dialog = _FakeFileDialog(monkeypatch, current_folder="C:/Users/me/Documents")
	player.select_file("Open||Window", "C:/Users/me/Documents/file.txt", force_slow_path_typing=True)
	assert dialog.sent_keys == [str(pathlib.Path("C:/Users/me/Documents")), "{ENTER}", "file.txt{ENTER}"]


# ------------------------------------- start_application -------------------------------------

class _FakeWindow:
	"""Stands for a top level window in the tests of 'start_application'."""
	def __init__(self, handle):
		self.handle = handle


class _FakePywinauto:
	"""
	Replaces the 'pywinauto' module in the tests of 'start_application'.
	Each call to 'Desktop().windows()' returns the next list of 'windows_lists', the last one is repeated.
	"""
	def __init__(self, *windows_lists, max_call_count=5):
		self.windows_lists = windows_lists
		self.max_call_count = max_call_count
		self.call_count = 0

	def Desktop(self, **kwargs):
		return self

	def Application(self, **kwargs):
		return self

	def start(self, **kwargs):
		pass

	def windows(self):
		self.call_count += 1
		if self.call_count > self.max_call_count:
			raise RuntimeError("start_application is waiting for a new window in an endless loop")
		return list(self.windows_lists[min(self.call_count, len(self.windows_lists)) - 1])


def test_start_application_returns_the_new_window(monkeypatch, no_sleep, native_window_handle):
	"""Tests that 'start_application' returns the application with the handle of the window that appeared."""
	old_window, new_window = _FakeWindow(1), _FakeWindow(2)
	monkeypatch.setattr(player, "pywinauto", _FakePywinauto([old_window], [old_window], [old_window, new_window]))
	application = player.start_application("app.exe")
	assert application.native_window_handle == 2
	assert core.get_native_window_handle() == 2


def test_start_application_fails_when_no_window_appears(monkeypatch, no_sleep, native_window_handle):
	"""Tests that 'start_application' raises FailedSearch after the timeout when no new window appears."""
	monkeypatch.setattr(player, "pywinauto", _FakePywinauto([_FakeWindow(1)]))
	with pytest.raises(FailedSearch, match="app.exe") as exc_info:
		player.start_application("app.exe", timeout=0)
	_assert_message_is_a_sentence(exc_info.value)
	assert core.get_native_window_handle() == native_window_handle


# ------------------------------------- messages of errors -------------------------------------

def test_connect_application_fails_with_a_sentence_when_the_window_is_not_found(monkeypatch, no_sleep):
	"""Tests the message of the FailedSearch raised by 'connect_application' when the main window is not found."""
	monkeypatch.setattr(player, "find_main_windows", lambda main_window_uipath="*": [])
	with pytest.raises(FailedSearch, match="Window not found") as exc_info:
		player.connect_application(main_window_uipath="Not found||Window", timeout=0)
	_assert_message_is_a_sentence(exc_info.value)
	assert "Not found||Window" in str(exc_info.value)


def test_wait_is_ready_fails_with_a_sentence_when_the_element_is_never_ready(monkeypatch):
	"""Tests the message of the TimeoutError raised when an element is never enabled."""
	class DisabledElement:
		def is_enabled(self):
			return False

		def is_visible(self):
			return True

	monkeypatch.setattr(player, "get_wrapper_path", lambda wrapper: "Window||Window->Disabled||Button")
	with pytest.raises(TimeoutError, match="Disabled") as exc_info:
		player.wait_is_ready_try1(DisabledElement(), timeout=0.05)
	_assert_message_is_a_sentence(exc_info.value)


# ------------------------------- find_all, focus_on_application -------------------------------

def test_find_all_returns_an_empty_list_when_no_element_is_found(fake_find_elements, no_sleep):
	"""Tests that 'find_all' returns an empty list, without raising an exception, when no element is found."""
	fake_find_elements([])
	assert player.find_all("Not found||Window->OK||Button", timeout=0.2) == []


def test_find_all_returns_all_the_elements(fake_find_elements):
	"""Tests that 'find_all' returns all the elements that match the path."""
	ok_1, ok_2 = _FakeElement("Window||Window->OK||Button"), _FakeElement("Window||Window->OK||Button")
	fake_find_elements([ok_1, ok_2])
	assert player.find_all("Window||Window->OK||Button", timeout=2) == [ok_1, ok_2]


@pytest.fixture
def fake_window_functions(monkeypatch, no_sleep, native_window_handle):
	"""Replaces the win32gui functions used by 'focus_on_application'. Returns the handles put in the foreground."""
	handles = []
	monkeypatch.setattr(player, "win32gui_IsIconic", lambda handle: False)
	monkeypatch.setattr(player, "win32gui_ShowWindow", lambda handle, command: None)
	monkeypatch.setattr(player, "win32gui_SetWindowPos", lambda handle, *args: handles.append(handle))
	return handles


def test_focus_on_application_with_an_application(fake_window_functions):
	"""Tests 'focus_on_application' with a UIApplication object."""
	player.focus_on_application(player.UIApplication(app=None, native_window_handle=111))
	assert core.get_native_window_handle() == 111
	assert fake_window_functions == [111, 111]


def test_focus_on_application_with_a_window(fake_window_functions):
	"""Tests 'focus_on_application' with the Pywinauto wrapper of a main window, as its documentation allows."""
	player.focus_on_application(_FakeWindow(handle=222))
	assert core.get_native_window_handle() == 222
	assert fake_window_functions == [222, 222]


def test_focus_on_application_with_none(fake_window_functions):
	"""Tests that 'focus_on_application(None)' clears the native window handle used by the searches."""
	player.focus_on_application(None)
	assert core.get_native_window_handle() is None
	assert fake_window_functions == []


# ------------------------------------------ playback ------------------------------------------

@pytest.fixture
def playback_environment(monkeypatch):
	"""Restores 'sys.path' after 'playback', and counts the prompts 'Press Enter to continue...'."""
	prompts = []
	monkeypatch.setattr(sys, "path", list(sys.path))
	monkeypatch.setattr(builtins, "input", lambda prompt="": prompts.append(prompt))
	return prompts


class _FakeStdin:
	def __init__(self, interactive):
		self.interactive = interactive

	def isatty(self):
		return self.interactive


def test_playback_returns_true_when_the_script_succeeds(playback_environment):
	"""Tests that 'playback' returns True when the script is played back without error."""
	assert player.playback(str_code="result = 1 + 1") is True
	assert playback_environment == []


@pytest.mark.parametrize("stdin", [_FakeStdin(interactive=False), None])
def test_playback_does_not_wait_for_enter_when_nobody_can_press_it(monkeypatch, capsys, playback_environment, stdin):
	"""
	Tests that 'playback' returns False and prints the error when the script fails, and that it does not
	wait for the Enter key when the input is not a console (scheduled task, CI, pythonw.exe).
	"""
	monkeypatch.setattr(sys, "stdin", stdin)
	assert player.playback(str_code="raise ValueError('error raised by the script')") is False
	assert "ValueError: error raised by the script" in capsys.readouterr().err
	assert playback_environment == []


def test_playback_waits_for_enter_in_a_console(monkeypatch, capsys, playback_environment):
	"""Tests that 'playback' still waits for the Enter key after an error when it is run in a console."""
	monkeypatch.setattr(sys, "stdin", _FakeStdin(interactive=True))
	assert player.playback(str_code="raise ValueError('error raised by the script')") is False
	assert "ValueError: error raised by the script" in capsys.readouterr().err
	assert playback_environment == ["Press Enter to continue..."]
