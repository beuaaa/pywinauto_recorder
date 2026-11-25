# -*- coding: utf-8 -*-
"""
Tests that need no application under test (they run in CI on Windows runners).

They check that Pywinauto Recorder calls pywinauto 1.0 with keywords it accepts:
pytest.ini turns pywinauto's "[pywinauto>=0.7.0] Keyword is renamed" warnings into
errors, so any old keyword (title, title_re, process, visible_only...) fails here.
"""

import pytest
import pywinauto
from pywinauto import findwindows
from pywinauto_recorder import core
from pywinauto_recorder.player import _translate_pywinauto_kwargs, find_main_windows

UNLIKELY_NAME = "Pywinauto Recorder CI - no such window 9f3c2a"


def test_pywinauto_1_0_is_installed():
	assert int(pywinauto.__version__.split(".")[0]) >= 1


@pytest.mark.parametrize("old, new", [
	({'title': 'Calculator'}, {'name': 'Calculator'}),
	({'title_re': 'Calc.*'}, {'name_re': 'Calc.*'}),
	({'process': 1234}, {'pid': 1234}),
	({'visible_only': True}, {'visible': True}),
	({'visible_only': False}, {}),
	({'enabled_only': True}, {'enabled': True}),
	({'enabled_only': False}, {}),
	({'handle': 42, 'timeout': 5}, {'handle': 42, 'timeout': 5}),
	({'name': 'Calculator', 'path': 'calc.exe'}, {'name': 'Calculator', 'path': 'calc.exe'}),
])
def test_translate_pywinauto_kwargs(old, new):
	assert _translate_pywinauto_kwargs(old) == new


@pytest.mark.parametrize("kwargs", [
	{'title': 'a', 'name': 'b'},
	{'name': 'b', 'title': 'a'},
	{'process': 1, 'pid': 2},
	{'visible_only': True, 'visible': True},
])
def test_translate_pywinauto_kwargs_conflict(kwargs):
	with pytest.raises(ValueError):
		_translate_pywinauto_kwargs(kwargs)


def test_find_window_candidates_exact_and_regex():
	assert core.find_window_candidates(UNLIKELY_NAME + "||Window") == []
	assert core.find_window_candidates("RegEx: " + UNLIKELY_NAME + ".*||Window") == []
	assert isinstance(core.find_window_candidates("*"), list)


def test_find_main_windows():
	assert isinstance(find_main_windows(), list)


def test_find_elements_not_found():
	assert core.find_elements(UNLIKELY_NAME + "||Window->Nothing||Button") == []


def test_find_elements_keywords_accepted():
	# Same keyword combinations as core.find_elements()
	findwindows.find_elements(
		name=UNLIKELY_NAME, control_type="Window", backend="uia", visible=True, parent=None, top_level_only=True)
	findwindows.find_elements(
		name_re=UNLIKELY_NAME, control_type="Window", backend="uia", visible=True, parent=None, top_level_only=True)


def test_element_info_properties_used_by_recorder():
	windows = core.desktop.windows()
	if not windows:
		pytest.skip("No top level window in this session")
	element_info = windows[0].element_info
	for prop in ("name", "control_type", "auto_id", "handle", "parent", "process_id"):
		getattr(element_info, prop)
	descendants = windows[0].descendants(name=UNLIKELY_NAME, control_type="Button")
	assert descendants == []
	assert isinstance(core.get_wrapper_path(windows[0]), str)


def test_visible_top_windows_uses_the_keywords_of_pywinauto_1_0():
	"""
	Tests that the search of the windows by 'start_application' calls pywinauto 1.0 without any renamed keyword.

	'_visible_top_windows' catches the exceptions, so the warnings are recorded instead of being raised.
	"""
	import warnings
	from pywinauto_recorder import player
	with warnings.catch_warnings(record=True) as caught:
		warnings.simplefilter("always")
		handles = player._visible_top_windows()
		player._visible_top_windows(title_re=UNLIKELY_NAME)
	assert not [w for w in caught if "Keyword is renamed" in str(w.message)]
	assert isinstance(handles, list)
	assert player._visible_top_windows(title_re=UNLIKELY_NAME) == []


def test_candidate_pids_with_the_real_processes():
	"""Tests '_candidate_pids' with psutil and the process running the tests."""
	import os
	import time
	import psutil
	from pywinauto_recorder import player
	me = psutil.Process(os.getpid())
	pids = player._candidate_pids(me.parent().pid, me.name().lower(), t_start=time.time() - 3600 * 24 * 365)
	assert os.getpid() in pids
