from pywinauto_recorder.core import get_entry_list, get_entry, match_entry_list


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

