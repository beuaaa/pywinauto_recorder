import builtins
import io
import pathlib
import re
import sys
import time
import warnings
import pytest
from pywinauto_recorder.core import get_entry_list, get_entry, match_entry_list, get_sorted_region
from pywinauto_recorder import core, player, recorder
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


@pytest.mark.parametrize("template, expected", [
	("Calculator||Window->Number pad||Group->One", True),
	("Calculator||Window->*->One", True),
	("*->One", True),
	("Calculator->*->One", True),
	("Calculator||Window->*->Number pad->One||Button", True),
	("Calculator||Window->*->Number pad||Group->One", True),
	("Calculator||Window->*->One||", True),
	("Calculator||Window->*->Two", False),
	("Calculator||Window->*->One||Text", False),
	("Calculator||Window->*->||Button", False),
	("Calculator||Window->*->Number pad->Two", False),
])
def test_match_entry_list_with_an_entry_without_control_type(template, expected):
	"""
	An entry without control type matches the elements that have its name, whatever their control type.

	Before, it matched no element after a '*' (found while documenting the paths, issue #35).
	"""
	element = "Calculator||Window->Number pad||Group->One||Button"
	assert match_entry_list(get_entry_list(element), get_entry_list(template)) == expected


@pytest.mark.parametrize("template, expected", [
	("*", True),
	("Calculator||Window->*", True),
	("Calculator||Window->*->*", True),
	("Calculator||Window->Number pad||Group->One||Button->*", True),
	("Calculator||Window->Number pad||Group->One||Button->*->*", True),
	("Calculator||Window->*->Number pad||Group->*->One||Button", True),
	("Calculator||Window", False),
	("Calculator||Window->Number pad||Group", False),
	("Calculator||Window->*->Number pad||Group", False),
	("Calculator||Window->Number pad||Group->One||Button->Two||Button", False),
	("Calculator||Window->*->Two||Button->*", False),
])
def test_match_entry_list_with_stars(template, expected):
	"""A '*' entry matches any number of entries, zero included, at any place of the path."""
	element = "Calculator||Window->Number pad||Group->One||Button"
	assert match_entry_list(get_entry_list(element), get_entry_list(template)) == expected



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
#                                  Tests of the path dictionary                                #
################################################################################################

TESTS_DIR = pathlib.Path(__file__).resolve().parent


@pytest.fixture
def empty_dictionary(monkeypatch):
	"""Replaces the path dictionary of the player by an empty one, restored after the test."""
	monkeypatch.setattr(player, "_dictionary", {})


def _dictionary_files(folder, keys, definitions, encoding="utf8"):
	"""Writes a key file and a definition file, and returns their names."""
	filename_key, filename_def = str(folder / "test.key"), str(folder / "test.def")
	pathlib.Path(filename_key).write_bytes(keys.encode(encoding))
	pathlib.Path(filename_def).write_bytes(definitions.encode(encoding))
	return filename_key, filename_def


def test_dictionary_of_the_calculator(empty_dictionary):
	"""Tests the dictionary of the example of the documentation, tests/Calculator.key and tests/Calculator.def."""
	player.load_dictionary(str(TESTS_DIR / "Calculator.key"), str(TESTS_DIR / "Calculator.def"))
	assert player.shortcut("Calculator") == "Calculator||Window->*"
	assert player.shortcut("Number pad") == "Number pad||Group"
	assert player.shortcut("1") == "One||Button"
	assert player.full_definition("1") == "Calculator||Window->*->Number pad||Group->One||Button"
	assert player.full_definition("=") == "Calculator||Window->*->Standard operators||Group->Equals||Button"
	assert player.full_definition("Number pad") == "Calculator||Window->*->Number pad||Group"
	# Before, the path of a key of the first level was repeated: "Calculator||Window->*->Calculator||Window->*"
	assert player.full_definition("Calculator") == "Calculator||Window->*"
	with UIPath(player.shortcut("Calculator")):
		with UIPath(player.shortcut("Number pad")):
			assert UIPath.get_full_path(player.shortcut("1")) == player.full_definition("1")


def test_dictionary_with_empty_lines_a_byte_order_mark_and_windows_line_breaks(empty_dictionary, tmp_path):
	"""
	Tests the files that Notepad can write: with a byte order mark and Windows line breaks.

	The empty lines are ignored, like the spaces at the end of the lines and the indentation of the key file.
	"""
	keys = "﻿Notepad\r\n\r\n    Text  \r\n\tMenu bar\r\n"
	definitions = "﻿RegEx: .* - Notepad||Window\r\n\r\n\tText editor||Document \r\n\tApplication||MenuBar\r\n"
	player.load_dictionary(*_dictionary_files(tmp_path, keys, definitions))
	assert player._dictionary == {
		"Notepad": ("", "RegEx: .* - Notepad||Window"),
		"Text": ("RegEx: .* - Notepad||Window", "Text editor||Document"),
		"Menu bar": ("RegEx: .* - Notepad||Window", "Application||MenuBar")}


def test_dictionaries_are_added(empty_dictionary, tmp_path):
	"""Tests that the keys of a dictionary are added to the keys already loaded, and that an unknown key is an error."""
	player.load_dictionary(*_dictionary_files(tmp_path, "A\n\tB\n", "A||Window\n\tB||Button\n"))
	player.load_dictionary(*_dictionary_files(tmp_path, "C\n", "C||Window\n"))
	assert (player.full_definition("B"), player.full_definition("C")) == ("A||Window->B||Button", "C||Window")
	with pytest.raises(KeyError):
		player.shortcut("D")


def test_dictionary_encoding(empty_dictionary, tmp_path):
	"""Tests that the files are read with the given encoding."""
	player.load_dictionary(*_dictionary_files(tmp_path, "Zéro\n", "Zéro||Button\n", encoding="cp1252"), encoding="cp1252")
	assert player.shortcut("Zéro") == "Zéro||Button"


@pytest.mark.parametrize("keys, definitions, message", [
	("A\n\tB\n", "A||Window\n", "test.key has 2 lines, and the definition file .*test.def has 1 lines"),
	("A\n\n", "A||Window\n\tB||Button\n", "Line 2 is empty in only one of the files"),
	("A\n\tB\n", "A||Window\n    B||Button\n", "test.def, line 2: the indentation must be made of tabs"),
	("A\n\tB\n", "A||Window\n\t\tB||Button\n", "test.def, line 2: the indentation has 2 tabs, but at most 1 are allowed"),
	("\tA\n", "\tA||Window\n", "test.def, line 1: the indentation has 1 tabs, but at most 0 are allowed"),
	("A\tB\n", "A||Window\n", "test.key, line 1: a tab separates the levels"),
	("A\n", "A||Window\tB||Button\n", "test.def, line 1: a tab separates the levels"),
])
def test_dictionary_errors(empty_dictionary, tmp_path, keys, definitions, message):
	"""Tests the errors of the files of a dictionary. Before, their keys were loaded with wrong paths, or not loaded."""
	with pytest.raises(ValueError, match=message):
		player.load_dictionary(*_dictionary_files(tmp_path, keys, definitions))


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
#                               Tests of the searches in a fake desktop                        #
################################################################################################

class _TreeElement:
	"""
	Stands for a UI Automation element_info of a fake desktop, in the tests of the searches.

	The requests to the application are counted in 'requests'.
	"""

	requests = 0

	def __init__(self, name, control_type, parent=None):
		self._name, self._control_type, self._parent = name, control_type, parent
		self.children = []
		if parent is not None:
			parent.children.append(self)

	def _read(self, value):
		_TreeElement.requests += 1
		return value

	name = property(lambda self: self._read(self._name))
	control_type = property(lambda self: self._read(self._control_type))
	automation_id = property(lambda self: self._read(self._name + "_id"))
	parent = property(lambda self: self._read(self._parent))

	def descendants(self):
		"""Returns the descendants of the element in the order of the tree, with one request like FindAll."""
		_TreeElement.requests += 1
		descendants, stack = [], list(reversed(self.children))
		while stack:
			element = stack.pop()
			descendants.append(element)
			stack.extend(reversed(element.children))
		return descendants


class _TreeWrapper:
	"""Stands for the Pywinauto wrapper of a _TreeElement."""

	def __init__(self, element_info):
		self.element_info = element_info

	def __eq__(self, other):
		return self.element_info is getattr(other, "element_info", other)

	def __hash__(self):
		return id(self.element_info)


def _tree_element_matches(element, title=None, title_re=None, control_type=None, visible_only=True, enabled_only=False):
	"""Returns True if an element of the fake desktop matches the criteria of a search of Pywinauto 0.6."""
	return (
		(title is None or element._name == title) and (title_re is None or re.match(title_re, element._name) is not None)
		and (control_type is None or element._control_type == control_type))


class _FakeDesktop:
	"""Replaces the desktop of Pywinauto and its searches, in the tests of the searches."""

	def __init__(self, monkeypatch):
		self.root = _TreeElement("Desktop", "Pane")
		monkeypatch.setattr(core, "desktop", self)
		monkeypatch.setattr(core, "findwindows", self)
		monkeypatch.setattr(core, "UIAWrapper", _TreeWrapper)

	def windows(self, active_only=False, handle=None, **criteria):
		"""Replaces Desktop.windows: returns the windows that match the criteria."""
		return [_TreeWrapper(window) for window in self.root.children if _tree_element_matches(window, **criteria)]

	def find_elements(self, backend=None, parent=None, top_level_only=True, **criteria):
		"""Replaces findwindows.find_elements: returns the windows, or descendants of parent, that match criteria."""
		parent = getattr(parent, "element_info", parent)
		elements = list(self.root.children) if top_level_only or parent is None else parent.descendants()
		return [element for element in elements if _tree_element_matches(element, **criteria)]

	def element(self, names):
		"""Returns the element at the end of a path of names, like 'Calculator->->Number pad->One'."""
		element = self.root
		for name in names.split("->"):
			element = next(child for child in element.children if child._name == name)
		return element


@pytest.fixture
def desktop(monkeypatch):
	"""Replaces the desktop by a fake desktop: a Calculator window, with its buttons in two groups, and a Notepad window."""
	monkeypatch.setattr(core, "native_window_handle", None)
	for name in ("last_parent_element_info", "last_parent_element_info_path", "last_top_level_parent"):
		monkeypatch.setattr(core, name, None)
	fake_desktop = _FakeDesktop(monkeypatch)
	calculator = _TreeElement("Calculator", "Window", fake_desktop.root)
	custom = _TreeElement("", "Custom", calculator)
	number_pad = _TreeElement("Number pad", "Group", custom)
	for name in ("One", "Two", "Three"):
		_TreeElement(name, "Button", number_pad)
	operators = _TreeElement("Standard operators", "Group", custom)
	for name in ("Plus", "Equals"):
		_TreeElement(name, "Button", operators)
	_TreeElement("Untitled - Notepad", "Window", fake_desktop.root)
	return fake_desktop


def _names(wrappers):
	"""Returns the names of the elements of the wrappers."""
	return [wrapper.element_info._name for wrapper in wrappers]


def test_get_element_info_path_from_an_element_of_the_window(desktop):
	"""
	Tests that the path kept for the parent of the last element is used only for the paths from the same element.

	Before, the path of an element from its window was used for the path of its sibling from an element of the window.
	"""
	calculator, number_pad = desktop.element("Calculator"), desktop.element("Calculator->->Number pad")
	one, two = desktop.element("Calculator->->Number pad->One"), desktop.element("Calculator->->Number pad->Two")
	assert core.get_element_info_path(one, calculator) == "Calculator||Window->||Custom->Number pad||Group->One||Button"
	assert core.get_element_info_path(two, number_pad) == "Number pad||Group->Two||Button"
	assert core.get_element_info_path(two, calculator) == "Calculator||Window->||Custom->Number pad||Group->Two||Button"


def test_get_element_info_path_reads_the_ancestors_of_the_elements_of_a_parent_once(desktop):
	"""Tests that the names of the ancestors are not read again for the path of an element of the same parent."""
	calculator = desktop.element("Calculator")
	one, two = desktop.element("Calculator->->Number pad->One"), desktop.element("Calculator->->Number pad->Two")
	core.get_element_info_path(one, calculator)
	_TreeElement.requests = 0
	assert core.get_element_info_path(two, calculator) == "Calculator||Window->||Custom->Number pad||Group->Two||Button"
	assert _TreeElement.requests == 4  # the name, the control type and the parent of Two, and the id of its parent


def test_get_element_info_path_of_a_window_from_itself(desktop):
	"""
	Tests that the path of a window from itself does not change the path kept for the parent of the last element.

	Before, the path kept became the path of the window without its last character: the path of the next element of the
	same parent was wrong, and the element was not found.
	"""
	calculator = desktop.element("Calculator")
	one, two = desktop.element("Calculator->->Number pad->One"), desktop.element("Calculator->->Number pad->Two")
	core.get_element_info_path(one, calculator)
	assert core.get_element_info_path(calculator, calculator) == "Calculator||Window"
	assert core.get_element_info_path(two, calculator) == "Calculator||Window->||Custom->Number pad||Group->Two||Button"


def test_get_element_info_path_after_an_element_whose_name_has_a_path_separator(desktop):
	"""
	Tests the path of an element after the path of its sibling whose name has '->'.

	Before, the path kept for their parent ended with the beginning of the name of the sibling.
	"""
	calculator, number_pad = desktop.element("Calculator"), desktop.element("Calculator->->Number pad")
	arrow = _TreeElement("x->y", "Button", number_pad)
	two = desktop.element("Calculator->->Number pad->Two")
	path_of_arrow = "Calculator||Window->||Custom->Number pad||Group->x->y||Button"
	assert core.get_element_info_path(arrow, calculator) == path_of_arrow
	assert core.get_element_info_path(two, calculator) == "Calculator||Window->||Custom->Number pad||Group->Two||Button"


def test_find_elements_after_the_search_of_a_window_with_a_regular_expression(desktop):
	"""
	Tests the search of an element after the search of its window with a regular expression.

	Before, the search of the window changed the path kept for the parent of the last element found, and the next
	element of the same parent was not found.
	"""
	assert _names(core.find_elements("Calculator||Window->*->One||Button")) == ["One"]
	assert _names(core.find_elements("RegEx: Calc.*||Window")) == ["Calculator"]
	assert _names(core.find_elements("Calculator||Window->*->Two||Button")) == ["Two"]


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

class _FakeMouse:
	"""
	Replaces 'move', the mouse events, the mouse cursor position and 'time.sleep' in the tests of 'drag_and_drop'.
	'log' is the sequence of what 'drag_and_drop' did.
	"""
	SCREEN_SIZE = (1920, 1080)

	def __init__(self, monkeypatch, positions):
		self.positions = positions  # {element path: position of the element}
		self.cursor_pos = (0, 0)
		self.log = []
		monkeypatch.setattr(player, "move", self.move)
		monkeypatch.setattr(player, "win32api_mouse_event", self.mouse_event)
		monkeypatch.setattr(player, "win32api_GetCursorPos", lambda: self.cursor_pos)
		monkeypatch.setattr(player, "win32api_GetSystemMetrics", lambda index: self.SCREEN_SIZE[index])
		monkeypatch.setattr(time, "sleep", lambda seconds: self.log.append(("pause", seconds)))

	def move(self, element_path, duration=None, **kwargs):
		if element_path not in self.positions:
			raise FailedSearch("No element found")
		self.log.append(("move to", element_path))
		if duration != -1:
			self.cursor_pos = self.positions[element_path]
		return element_path

	def mouse_event(self, event, x=0, y=0):
		if event == player.MOUSEEVENTF_MOVE | player.MOUSEEVENTF_ABSOLUTE:
			# x and y are normalized absolute coordinates: 0..65535
			self.cursor_pos = (round(x * (self.SCREEN_SIZE[0] - 1) / 65535), round(y * (self.SCREEN_SIZE[1] - 1) / 65535))
			self.log.append(("mouse at", self.cursor_pos))
		else:
			self.log.append(event)

	def positions_after_the_arrival(self):
		"""Returns the positions of the mouse cursor set by the mouse events sent after the move to the destination."""
		return [entry[1] for entry in self.log if isinstance(entry, tuple) and entry[0] == "mouse at"]


def test_drag_and_drop_sequence(monkeypatch):
	"""
	Tests the sequence of 'drag_and_drop': the button is pressed on the source, then there is a pause, the move to
	the destination, a small move back and forth on the path that was followed, and a pause before the release.
	Some applications ignore the drop when the button is released as soon as the mouse arrives on the destination.
	"""
	mouse = _FakeMouse(monkeypatch, {"Source||Pane": (100, 100), "Destination||Pane": (300, 100)})
	pause = player._DRAG_AND_DROP_PAUSE
	assert pause > 0
	assert player.drag_and_drop("Source||Pane", "Destination||Pane", duration=0) == "Destination||Pane"
	assert mouse.log == [
		("move to", "Source||Pane"),
		player.MOUSEEVENTF_LEFTDOWN,
		("pause", pause),
		("move to", "Destination||Pane"),
		("mouse at", (298, 100)),
		("mouse at", (300, 100)),
		("pause", pause),
		player.MOUSEEVENTF_LEFTUP]


def test_drag_and_drop_moves_back_and_forth_on_the_path(monkeypatch):
	"""Tests that the small move before the release goes 2 pixels back towards the source, whatever the direction."""
	mouse = _FakeMouse(monkeypatch, {"Source||Pane": (500, 500), "Destination||Pane": (530, 540)})
	player.drag_and_drop("Source||Pane", "Destination||Pane", duration=0, button=player.ButtonLocation.right)
	assert mouse.positions_after_the_arrival() == [(529, 538), (530, 540)]
	assert mouse.log[1] == player.MOUSEEVENTF_RIGHTDOWN
	assert mouse.log[-1] == player.MOUSEEVENTF_RIGHTUP
	assert mouse.cursor_pos == (530, 540)


def test_drag_and_drop_does_not_move_back_and_forth_when_the_mouse_does_not_move(monkeypatch):
	"""
	Tests that 'drag_and_drop' does not add a move when the destination is the source,
	or when the duration is -1 (the mouse cursor doesn't move).
	"""
	mouse = _FakeMouse(monkeypatch, {"Source||Pane": (100, 100), "Destination||Pane": (100, 100)})
	player.drag_and_drop("Source||Pane", "Destination||Pane", duration=0)
	assert mouse.positions_after_the_arrival() == []

	mouse = _FakeMouse(monkeypatch, {"Source||Pane": (100, 100), "Destination||Pane": (300, 100)})
	player.drag_and_drop("Source||Pane", "Destination||Pane", duration=-1)
	assert mouse.positions_after_the_arrival() == []
	assert mouse.log[-1] == player.MOUSEEVENTF_LEFTUP


def test_drag_and_drop_releases_the_button_when_the_destination_is_not_found(monkeypatch):
	"""Tests that 'drag_and_drop' releases the mouse button when the destination element is not found."""
	mouse = _FakeMouse(monkeypatch, {"Source||Pane": (100, 100)})
	with pytest.raises(FailedSearch):
		player.drag_and_drop("Source||Pane", "Destination||Pane", duration=0)
	assert [entry for entry in mouse.log if not isinstance(entry, tuple)] == [player.MOUSEEVENTF_LEFTDOWN, player.MOUSEEVENTF_LEFTUP]
	assert mouse.log[-1] == player.MOUSEEVENTF_LEFTUP


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

@pytest.mark.parametrize("cmd_line, exe_name", [
	('"C:\\Program Files\\App\\app.exe" -x', "app.exe"),
	("C:\\Program Files\\App\\App.EXE -x", "app.exe"),
	("notepad.exe file.txt", "notepad.exe"),
	("calc", "calc.exe"),
	('"C:\\Tools\\tool" -x', "tool.exe"),
	('python "C:\\Scripts\\run.exe"', "python.exe"),
	("", None),
])
def test_exe_name(cmd_line, exe_name):
	"""Tests the name of the executable file found in a command line."""
	assert player._exe_name(cmd_line) == exe_name


class _FakeProcess:
	"""Stands for a psutil.Process."""
	def __init__(self, fake_psutil, pid, name, create_time=0.0, parent_pid=None, running=True):
		self.fake_psutil = fake_psutil
		self.pid = pid
		self.info = {"name": name, "create_time": create_time}
		self.parent_pid = parent_pid
		self.running = running

	def children(self, recursive=False):
		children = [p for p in self.fake_psutil.processes.values() if p.parent_pid == self.pid and p.running]
		if recursive:
			for child in list(children):
				children += child.children(recursive=True)
		return children

	def is_running(self):
		return self.running

	def name(self):
		return self.info["name"]


class _FakePsutil:
	"""Replaces the 'psutil' module: 'processes' are the processes of the system, by ID."""
	class Error(Exception):
		pass

	class NoSuchProcess(Error):
		pass

	def __init__(self):
		self.processes = {}

	def add(self, pid, name, create_time=0.0, parent_pid=None, running=True):
		self.processes[pid] = _FakeProcess(self, pid, name, create_time, parent_pid, running)

	def Process(self, pid):
		process = self.processes.get(pid)
		if process is None or not process.running:
			raise self.NoSuchProcess(pid)
		return process

	def process_iter(self, attrs=None):
		return [p for p in self.processes.values() if p.running]


def test_candidate_pids(monkeypatch):
	"""
	Tests that the candidate processes are the launcher, its descendants and the processes of the same executable
	started after the launcher, but not the older processes of this executable nor the other processes.
	"""
	fake_psutil = _FakePsutil()
	fake_psutil.add(100, "launcher.exe", create_time=1000.0)
	fake_psutil.add(101, "helper.exe", create_time=1000.5, parent_pid=100)
	fake_psutil.add(102, "worker.exe", create_time=1000.6, parent_pid=101)
	fake_psutil.add(200, "app.exe", create_time=1001.0)  # started by the launcher, but not its child
	fake_psutil.add(201, "App.exe", create_time=900.0)  # an older instance of the application
	fake_psutil.add(300, "other.exe", create_time=1001.0)
	monkeypatch.setattr(player, "psutil", fake_psutil)
	assert player._candidate_pids(100, "app.exe", t_start=1000.0) == {100, 101, 102, 200}
	assert player._candidate_pids(100, None, t_start=1000.0) == {100, 101, 102}


def test_candidate_pids_when_the_launcher_has_exited(monkeypatch):
	"""Tests the candidate processes when the launcher has already exited."""
	fake_psutil = _FakePsutil()
	fake_psutil.add(100, "launcher.exe", running=False)
	fake_psutil.add(200, "app.exe", create_time=1001.0)
	monkeypatch.setattr(player, "psutil", fake_psutil)
	assert player._candidate_pids(100, "app.exe", t_start=1000.0) == {100, 200}


class _FakeDesktopForStart:
	"""
	Replaces pywinauto.Application, findwindows.find_windows, GetWindowThreadProcessId, psutil and time
	in the tests of 'start_application'.
	Each search of the windows returns the next list of 'windows_lists' (the last one is repeated). A window is a tuple
	(handle, pid of the process that owns it, title).
	"""
	def __init__(self, monkeypatch, *windows_lists, launcher_pid=100, launcher_running=True):
		self.windows_lists = windows_lists
		self.search_count = 0
		self.find_windows_kwargs = []
		self.fake_psutil = _FakePsutil()
		self.fake_psutil.add(launcher_pid, "launcher.exe", create_time=1000.0, running=launcher_running)
		self.launcher_pid = launcher_pid
		self.now = 1000.0
		self.owners = {}
		self.started_cmd_lines = []
		monkeypatch.setattr(player, "psutil", self.fake_psutil)
		monkeypatch.setattr(player, "pywinauto", self)
		monkeypatch.setattr(player, "findwindows", self)
		monkeypatch.setattr(player, "win32process_GetWindowThreadProcessId", self.get_window_thread_process_id)
		monkeypatch.setattr(time, "time", lambda: self.now)
		monkeypatch.setattr(time, "sleep", self.sleep)

	def sleep(self, seconds):
		self.now += seconds

	# --- pywinauto.Application
	def Application(self, backend=None):
		self.process = None
		return self

	def start(self, cmd_line=None, **kwargs):
		self.started_cmd_lines.append(cmd_line)
		self.process = self.launcher_pid

	# --- findwindows.find_windows
	def find_windows(self, **kwargs):
		self.find_windows_kwargs.append(kwargs)
		assert kwargs.get("top_level_only", True) is True
		assert kwargs.get("visible", kwargs.get("visible_only")) is True
		windows = self.windows_lists[min(self.search_count, len(self.windows_lists) - 1)]
		self.search_count += 1
		for handle, pid, title in windows:
			self.owners[handle] = pid
		title_re = kwargs.get("name_re", kwargs.get("title_re"))
		return [handle for handle, pid, title in windows if title_re is None or re.match(title_re, title)]

	def get_window_thread_process_id(self, handle):
		if handle not in self.owners:
			raise Exception("Invalid window handle")
		return 1, self.owners[handle]


def test_start_application_returns_the_new_window_of_the_application(monkeypatch, native_window_handle):
	"""
	Tests that 'start_application' returns the first new window of the started process, and ignores the windows
	that were already there and the new windows of other applications.
	"""
	desktop = _FakeDesktopForStart(
		monkeypatch,
		[(1, 100, "Old window of the application"), (2, 300, "Other application")],  # before the start
		[(1, 100, "Old window of the application"), (2, 300, "Other application")],
		[(3, 300, "Notification"), (1, 100, "Old window of the application"), (2, 300, "Other application")],
		[(3, 300, "Notification"), (4, 100, "Application"), (1, 100, "Old window of the application")])
	application = player.start_application("app.exe -x")
	assert desktop.started_cmd_lines == ["app.exe -x"]
	assert application.native_window_handle == 4
	assert application.app is desktop
	assert core.get_native_window_handle() == 4


def test_start_application_with_a_launcher_that_exits(monkeypatch, native_window_handle):
	"""
	Tests 'start_application' with a launcher that starts the real executable then exits:
	the window of the executable started after the launcher is returned, not the window of an older instance.
	"""
	desktop = _FakeDesktopForStart(monkeypatch, [(1, 201, "Old instance")], [(1, 201, "Old instance"), (2, 200, "App")])
	desktop.fake_psutil.add(201, "app.exe", create_time=900.0)
	desktop.fake_psutil.add(200, "app.exe", create_time=1000.5)
	desktop.fake_psutil.processes[100].running = False
	assert player.start_application('"C:\\Program Files\\App\\app.exe"').native_window_handle == 2


def test_start_application_with_a_uwp_application(monkeypatch, native_window_handle):
	"""
	Tests 'start_application' with a UWP application like the Calculator: the launcher exits at once and the
	main window belongs to ApplicationFrameHost.exe, a process of the system that is not related to the launcher.
	"""
	desktop = _FakeDesktopForStart(
		monkeypatch,
		[(1, 500, "Calculator")],  # a first Calculator, owned by ApplicationFrameHost.exe
		[(2, 500, "Calculator"), (1, 500, "Calculator")],
		launcher_running=False)
	desktop.fake_psutil.add(500, "ApplicationFrameHost.exe", create_time=10.0)
	assert player.start_application("calc").native_window_handle == 2


def test_start_application_ignores_uwp_windows_while_the_launcher_is_running(monkeypatch, native_window_handle):
	"""
	Tests that 'start_application' does not take a new window of ApplicationFrameHost.exe while the launcher is still
	running: it is the window of another UWP application, not the window of the started application.
	"""
	desktop = _FakeDesktopForStart(
		monkeypatch,
		[],
		[(1, 500, "Settings")],
		[(1, 500, "Settings"), (2, 100, "Application")])
	desktop.fake_psutil.add(500, "ApplicationFrameHost.exe", create_time=10.0)
	assert player.start_application("app.exe").native_window_handle == 2


def test_start_application_with_title_re(monkeypatch, native_window_handle):
	"""Tests that 'start_application' returns the first new window whose title matches title_re (not the splash screen)."""
	desktop = _FakeDesktopForStart(
		monkeypatch,
		[],
		[(1, 100, "Splash screen")],
		[(1, 100, "Splash screen")],
		[(2, 100, "My App - Document 1"), (1, 100, "Splash screen")])
	assert player.start_application("app.exe", title_re="My App - .*").native_window_handle == 2
	assert any(kwargs.get("name_re", kwargs.get("title_re")) == "My App - .*" for kwargs in desktop.find_windows_kwargs)


def test_start_application_ignores_a_window_destroyed_during_the_search(monkeypatch, native_window_handle):
	"""Tests that 'start_application' ignores a window that is destroyed while it is searched."""
	desktop = _FakeDesktopForStart(monkeypatch, [], [(1, 100, "Destroyed"), (2, 100, "Application")])
	real_find_windows = desktop.find_windows

	def find_windows(**kwargs):
		handles = real_find_windows(**kwargs)
		desktop.owners.pop(1, None)  # window 1 is destroyed before its process is read
		return handles

	monkeypatch.setattr(player, "findwindows", type("FindWindows", (), {"find_windows": staticmethod(find_windows)}))
	assert player.start_application("app.exe").native_window_handle == 2


def test_start_application_fails_when_no_window_appears(monkeypatch, native_window_handle):
	"""Tests that 'start_application' raises TimeoutError after the timeout when no window of the application appears."""
	desktop = _FakeDesktopForStart(monkeypatch, [(1, 300, "Other application")], [(1, 300, "Other application"), (2, 300, "Other")])
	with pytest.raises(TimeoutError, match="app.exe") as exc_info:
		player.start_application("app.exe", timeout=3)
	_assert_message_is_a_sentence(exc_info.value)
	assert 3.0 <= desktop.now - 1000.0 <= 4.0  # it waited the timeout
	assert core.get_native_window_handle() == native_window_handle


def test_start_application_fails_at_once_when_timeout_is_0(monkeypatch, native_window_handle):
	"""Tests that 'start_application' searches the window once, then raises TimeoutError, when timeout is 0."""
	desktop = _FakeDesktopForStart(monkeypatch, [])
	with pytest.raises(TimeoutError):
		player.start_application("app.exe", timeout=0)
	assert desktop.search_count == 2  # the windows before the start, then one search


class _FakeWindow:
	"""Stands for a top level window in the test of 'start_application' with main_window_uipath."""
	def __init__(self, handle):
		self.handle = handle


class _FakePywinautoWithDesktop:
	"""
	Replaces the 'pywinauto' module in the test of 'start_application' with main_window_uipath.
	Each call to 'Desktop().windows()' returns the next list of 'windows_lists', the last one is repeated.
	"""
	def __init__(self, *windows_lists):
		self.windows_lists = windows_lists
		self.call_count = 0

	def Desktop(self, **kwargs):
		return self

	def Application(self, **kwargs):
		return self

	def start(self, **kwargs):
		pass

	def windows(self):
		self.call_count += 1
		return list(self.windows_lists[min(self.call_count, len(self.windows_lists)) - 1])


def test_start_application_with_main_window_uipath(monkeypatch, no_sleep, native_window_handle):
	"""
	Tests that 'start_application' with main_window_uipath connects to the main window found with this UIPath,
	excluding the main windows that were there before the start.
	"""
	old_window, splash_screen = _FakeWindow(1), _FakeWindow(2)
	connect_kwargs = []

	def connect_application(**kwargs):
		connect_kwargs.append(kwargs)
		return player.UIApplication(app=None, native_window_handle=3)

	monkeypatch.setattr(player, "pywinauto", _FakePywinautoWithDesktop([old_window], [old_window, splash_screen]))
	monkeypatch.setattr(player, "connect_application", connect_application)
	application = player.start_application("app.exe", timeout=7, main_window_uipath="App||Window")
	assert application.native_window_handle == 3
	assert core.get_native_window_handle() == 3
	assert connect_kwargs == [{"exclude_main_windows": [old_window], "main_window_uipath": "App||Window", "timeout": 7}]


def test_start_application_with_main_window_uipath_fails_when_no_window_appears(monkeypatch, no_sleep, native_window_handle):
	"""Tests that 'start_application' with main_window_uipath raises TimeoutError when no new window appears."""
	monkeypatch.setattr(player, "pywinauto", _FakePywinautoWithDesktop([_FakeWindow(1)]))
	with pytest.raises(TimeoutError, match="app.exe") as exc_info:
		player.start_application("app.exe", timeout=0, main_window_uipath="App||Window")
	_assert_message_is_a_sentence(exc_info.value)


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


################################################################################################
#                                   Tests of bin/VersionInfo.rc                                #
################################################################################################

def test_version_info_rc_has_the_version_of_the_package():
	"""
	Tests that bin/VersionInfo.rc gives the version of the package, read the same way as in pywinauto_recorder.py.
	pywinauto_recorder.py reads this file on the 'master' branch to tell the users that a new version is available.
	"""
	import pywinauto_recorder
	version_info = pathlib.Path(pywinauto_recorder.__file__).resolve().parent.parent / "bin" / "VersionInfo.rc"
	line = version_info.read_bytes().decode('utf-8')
	latest_version = line.split('ProductVersion')[1].split('"')[2]
	assert latest_version == pywinauto_recorder.__version__


def test_tray_icon_message_handlers_return_an_int(monkeypatch):
	"""
	Tests that the handlers of the messages of the tray icon return an int: the result of the window procedure. Before,
	restart, destroy and command returned None, which recent versions of pywin32 refuse: "WNDPROC return value cannot be
	converted to LRESULT" was printed, twice when quitting for example.
	"""
	import importlib.util
	import os
	import types
	launcher_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "pywinauto_recorder.py")
	spec = importlib.util.spec_from_file_location("pywinauto_recorder_launcher", launcher_path)
	launcher = importlib.util.module_from_spec(spec)
	spec.loader.exec_module(launcher)
	calls = []
	monkeypatch.setattr(launcher, "win32gui", types.SimpleNamespace(
		NIM_DELETE=2, LOWORD=lambda value: value & 0xFFFF,
		Shell_NotifyIcon=lambda message, nid: calls.append("Shell_NotifyIcon"),
		PostQuitMessage=lambda exit_code: calls.append("PostQuitMessage")))
	tray = object.__new__(launcher.SysTrayIcon)
	tray.hwnd = 42
	tray.on_quit = lambda tray_: calls.append("on_quit")
	tray.menu_actions_by_id = {1024: lambda tray_: calls.append("action")}
	tray.default_menu_index = 1  # FIRST_ID + 1 = 1024
	tray.refresh_icon = lambda: calls.append("refresh_icon")
	tray.show_menu = lambda: calls.append("show_menu")
	results = [
		tray.restart(42, 0, 0, 0),
		tray.command(42, 0, 1024, 0),
		tray.notify(42, 0, 0, launcher.win32con.WM_LBUTTONDBLCLK),
		tray.notify(42, 0, 0, launcher.win32con.WM_RBUTTONUP),
		tray.destroy(42, 0, 0, 0)]
	assert all(isinstance(result, int) for result in results), results
	assert calls == ["refresh_icon", "action", "action", "show_menu", "on_quit", "Shell_NotifyIcon", "PostQuitMessage"]


################################################################################################
#                         Tests of the keyboard events written by the recorder                 #
################################################################################################

class _Rectangle(tuple):
	"""Replaces the rectangle of an element: (left, top, right, bottom)."""

	left = property(lambda self: self[0])
	top = property(lambda self: self[1])
	right = property(lambda self: self[2])
	bottom = property(lambda self: self[3])

	def mid_point(self):
		"""Returns the center of the rectangle."""
		return (self.left + self.right) // 2, (self.top + self.bottom) // 2

	def width(self):
		"""Returns the width of the rectangle."""
		return self.right - self.left

	def height(self):
		"""Returns the height of the rectangle."""
		return self.bottom - self.top


def _move(x, y):
	"""Returns a mouse move event."""
	return recorder.mouse.MoveEvent(x, y, 0.0)


def _button(event_type, time_=1.0, button="left"):
	"""Returns a mouse button event: "down" or "up"."""
	return recorder.mouse.ButtonEvent(event_type, button, time_)


def _element(path, left, top, right, bottom):
	"""Returns the event of an element found under the mouse cursor by the loop of the recorder."""
	return recorder.ElementEvent(None, _Rectangle((left, top, right, bottom)), path)


_A = _element("W||Window->A||Button", 0, 0, 102, 42)
_B = _element("W||Window->B||Button", 200, 0, 302, 42)


_MODIFIER_SCAN_CODES = {"shift": 42, "ctrl": 29, "alt": 56, "alt gr": 541, "windows": 91}
"""The scan codes of the modifier keys in the keyboard events of Windows (541: the left ctrl sent with alt gr)."""


def _key(event_type, name, scan_code=0, is_keypad=False):
	"""Returns a keyboard event, like the 'keyboard' library gives it on Windows."""
	return recorder.keyboard.KeyboardEvent(event_type, scan_code, name=name, time=0.0, is_keypad=is_keypad)


def _press(name, scan_code=0, is_keypad=False):
	"""Returns the press and the release of a key."""
	return [_key("down", name, scan_code, is_keypad), _key("up", name, scan_code, is_keypad)]


def _typed(text):
	"""Returns the presses and the releases of the keys that type a text."""
	return [event for character in text for event in _press("space" if character == " " else character)]


def _held(modifier, events, name=None):
	"""Returns events while a modifier key is held. name: the name given by Windows in its language, like "maj"."""
	return [_key("down", name or modifier, _MODIFIER_SCAN_CODES[modifier])] + events + [
		_key("up", name or modifier, _MODIFIER_SCAN_CODES[modifier])]


_KEYBOARD_CASES = [
	pytest.param(_typed("2"), ['"2"'], id="key alone (issue #60)"),
	pytest.param(_typed("hello"), ['"hello"'], id="text without ENTER"),
	pytest.param(_held("shift", _typed("H")) + _typed("ello"), ['"Hello"'], id="shift"),
	pytest.param(_held("shift", _typed("H"), name="maj") + _typed("i"), ['"Hi"'], id="shift in French"),
	pytest.param(
		_typed("hello") + _press("tab", 15) + _typed("world") + _press("enter", 28), ['"hello{TAB}world{ENTER}"'],
		id="tab and enter"),
	pytest.param(_held("ctrl", _typed("s")), ['"{VK_CONTROL down}s{VK_CONTROL up}"'], id="ctrl+s"),
	pytest.param(_held("alt", _press("f4", 62)), ['"{VK_MENU down}{F4}{VK_MENU up}"'], id="alt+f4"),
	pytest.param(_held("shift", _press("tab", 15)), ['"{VK_SHIFT down}{TAB}{VK_SHIFT up}"'], id="shift+tab"),
	pytest.param(
		_held("ctrl", _held("shift", _press("esc", 1))),
		['"{VK_CONTROL down}{VK_SHIFT down}{ESC}{VK_SHIFT up}{VK_CONTROL up}"'], id="ctrl+shift+esc"),
	pytest.param(_typed("50%") + _press("enter", 28), ['"50{%}{ENTER}"'], id="percent"),
	pytest.param(
		_typed('{a}^b~c(d)e+f"g\\h'), ['"{{}a{}}{^}b{~}c(d)e+f\\"g\\\\h"'], id="special characters"),
	pytest.param(
		_press("suppr", 83) + _press("haut", 72) + _press("fin", 79) + _press("pg. suiv.", 81),
		['"{DELETE}{UP}{END}{PGDN}"'], id="keys named in French"),
	pytest.param(
		_typed("a") + _held("alt gr", _typed("@"), name="ctrl") + _typed("b.fr"), ['"a@b.fr"'], id="alt gr"),
	pytest.param(
		_typed("ab") + _held("ctrl", _typed("a")) + _held("shift", _typed("C")) + _typed("d"),
		['"ab{VK_CONTROL down}a{VK_CONTROL up}"', '"Cd"'], id="text after a shortcut"),
	pytest.param(_press("alt", 56) + _typed("f"), ['"{VK_MENU}f"'], id="alt alone"),
	pytest.param(_press("windows gauche", 91), ['"{LWIN}"'], id="windows alone"),
	pytest.param(
		_held("windows", _typed("e"), name="left windows"), ['"{VK_LWIN down}e{VK_LWIN up}"'], id="windows+e"),
	pytest.param(_press("ctrl", 29) + _press("shift", 42) + _typed("x"), ['"x"'], id="ctrl and shift alone"),
	pytest.param(
		_typed("a") + [_key("down", "ctrl", 29), _key("down", "alt", 56), _key("down", "s", 31), _key("up", "ctrl", 29),
		_key("up", "alt", 56)] + _typed("b"), ['"ab"'], id="shortcut of the recorder"),
	pytest.param(_typed("helo") + _press("backspace", 14) + _typed("lo"), ['"hello"'], id="backspace"),
	pytest.param(_press("backspace", 14) + _typed("ok"), ['"{BACKSPACE}ok"'], id="backspace of a previous text"),
	pytest.param(_press("caps lock", 58) + _typed("AB") + _press("verr.maj", 58), ['"AB"'], id="caps lock"),
	pytest.param(
		_press("1", 79, True) + _press("decimal", 83, True) + _press("end", 79, True), ['"1{VK_DECIMAL}{END}"'],
		id="numeric keypad"),
	pytest.param(
		_typed("a b") + _held("ctrl", _typed(" ")), ['"a b{VK_CONTROL down}{SPACE}{VK_CONTROL up}"'], id="space"),
	pytest.param(_typed("a") + _press("volume up", 48) + _typed("b"), ['"ab"'], id="unknown key"),
]
"""Keyboard events, and the strings of send_keys written in the script."""


@pytest.mark.parametrize("events, expected", _KEYBOARD_CASES)
def test_process_events_keyboard(events, expected):
	"""
	Tests the strings of send_keys written for keyboard events.

	Before, a text was lost when it was not followed by ENTER, like the keys that type no character, and the keys that
	type no character were written with names that send_keys does not know, like {f4} (issue #60).
	"""
	events = list(events)
	recorder._process_events(events)
	assert events == [recorder.SendKeysEvent(line) for line in expected]


def test_process_events_key_between_clicks():
	"""Tests that a key pressed between 2 clicks is written, like a digit typed in the Calculator (issue #60)."""
	events = [_move(10, 10), _A, _button("down"), _button("up", 1.1)] + _typed("2") + [
		_move(210, 10), _B, _button("down", 2.0), _button("up", 2.1)]
	recorder._process_events(events)
	assert [e for e in events if isinstance(e, (recorder.ClickEvent, recorder.SendKeysEvent))] == [
		recorder.ClickEvent("left", 1, "W||Window->A||Button", -0.82, -0.55, 1.1),
		recorder.SendKeysEvent('"2"'),
		recorder.ClickEvent("left", 1, "W||Window->B||Button", -0.82, -0.55, 2.1)]


def test_process_events_ctrl_click():
	"""
	Tests that CTRL stays pressed during a click, and that a shortcut typed after the click presses CTRL again.

	send_keys sends a character as a virtual key only after the press or the release of a key in the same string.
	"""
	events = [_key("down", "ctrl", 29), _move(10, 10), _A, _button("down"), _button("up", 1.1)] + _typed("c") + [
		_key("up", "ctrl", 29)]
	recorder._process_events(events)
	assert [e for e in events if isinstance(e, (recorder.ClickEvent, recorder.SendKeysEvent))] == [
		recorder.SendKeysEvent('"{VK_CONTROL down}"'),
		recorder.ClickEvent("left", 1, "W||Window->A||Button", -0.82, -0.55, 1.1),
		recorder.SendKeysEvent('"{VK_CONTROL down}c{VK_CONTROL up}"')]


def test_process_events_releases_the_modifier_keys_at_the_end():
	"""Tests that a modifier key still pressed at the end of the recording is released at the end of the script."""
	events = [_key("down", "ctrl", 29), _move(10, 10), _A, _button("down"), _button("up", 1.1)]
	recorder._process_events(events)
	assert [e for e in events if isinstance(e, (recorder.ClickEvent, recorder.SendKeysEvent))] == [
		recorder.SendKeysEvent('"{VK_CONTROL down}"'),
		recorder.ClickEvent("left", 1, "W||Window->A||Button", -0.82, -0.55, 1.1),
		recorder.SendKeysEvent('"{VK_CONTROL up}"')]


def test_process_events_keyboard_with_mouse_moves():
	"""Tests that the mouse moves between keyboard events do not cut the text in several strings of send_keys."""
	events = _typed("ab") + [_move(10, 10), _A] + _typed("c")
	recorder._process_events(events)
	assert events == [recorder.SendKeysEvent('"abc"'), _move(10, 10), _A]


@pytest.mark.skipif(sys.platform != "win32", reason="pywinauto needs Windows")
def test_send_keys_strings_are_parsed_by_pywinauto(monkeypatch):
	"""Tests that pywinauto parses the strings of send_keys written by the recorder, escaped by player.send_keys."""
	import ast
	import pywinauto.keyboard
	sent = []
	monkeypatch.setattr(pywinauto.keyboard, "send_keys", lambda keys, **kwargs: sent.append(keys))
	for case in _KEYBOARD_CASES:
		for line in case.values[1]:
			player.send_keys(ast.literal_eval(line))
	assert len(sent) == sum(len(case.values[1]) for case in _KEYBOARD_CASES)
	for keys in sent:
		assert pywinauto.keyboard.parse_keys(keys, with_spaces=True, with_tabs=True, with_newlines=True)
