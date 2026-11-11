from pywinauto_recorder.core import get_entry_list, get_entry, match_entry_list
from pywinauto_recorder.player import UIPath


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
