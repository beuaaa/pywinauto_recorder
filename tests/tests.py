# -*- coding: utf-8 -*-

import pytest
from pywinauto_recorder.core import *


################################################################################################
#                                   Tests of low level functions                               #
################################################################################################

def test_get_entry_list():
	""" Tests get_entry_list(element_path) """
	element_path = "D:\\Name||Window->||Pane->Property||Group->"
	element_path = element_path + "Label:||Text#[Name||Window->||Pane->Property||Group->Label:||Text,0]%(0,0)"
	entry_list = get_entry_list(element_path)
	assert entry_list[0] == 'D:\\Name||Window'
	assert entry_list[1] == '||Pane'
	assert entry_list[2] == 'Property||Group'
	assert entry_list[3] == 'Label:||Text#[Name||Window->||Pane->Property||Group->Label:||Text,0]%(0,0)'


def test_get_entry_list_with_asterisk():
	""" Tests get_entry_list(element_path) element_path has some asterisks"""
	element_path = "Name||Window->*->||Group->Property||%(0,0)"
	entry_list2 = get_entry_list(element_path)
	assert entry_list2[1] == '*'
	assert entry_list2[2] == '||Group'
	assert entry_list2[3] == 'Property||%(0,0)'


def test_get_entry_elements():
	""" Tests get_entry(entry) """
	entry_list = [
		'Name:||Type#[0,0]%(2,-24)', '||Type#[0,0]%(2,-24)', 'Name:||#[0,0]%(2,-24)',
		'Name:||Type#[0,0]', '||Type#[0,0]', 'Name:||#[0,0]',
		'Name:||Type', '||Type', 'Name:||',
		'Name||Type', '||Type', 'Name||',
		'Name||Type#[y_name:||y_type,0]', '||Type#[y_name:||y_type,0]', 'Name||#[y_name:||y_type,0]',
		'Name||Type#[y_name||y_type,0]', '||Type#[y_name||y_type,0]', 'Name||#[y_name||y_type,0]',
	]

	for i, entry in enumerate(entry_list):
		str_name, str_type, y_x, dx_dy = get_entry(entry)
		if i % 3 == 0:
			if i < 9:
				assert str_name == 'Name:'
			else:
				assert str_name == 'Name'
			assert str_type == 'Type'
		if i % 3 == 1:
			assert str_name == ''
			assert str_type == 'Type'
		if i % 3 == 2:
			if i < 9:
				assert str_name == 'Name:'
			else:
				assert str_name == 'Name'
			assert str_type is None
		if i < 3:
			assert dx_dy == (2, -24)
		else:
			assert dx_dy is None
		if i < 6:
			assert y_x == [0, 0]
		elif i >= 15:
			assert y_x == ['y_name||y_type', 0]
		elif i >= 12:
			assert y_x == ['y_name:||y_type', 0]
		else:
			assert y_x is None


def test_match_entry_list():
	""" Tests match_entry_list(entry_list, template_list, regex_title=False) """

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

