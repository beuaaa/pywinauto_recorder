# -*- coding: utf-8 -*-
"""
Tests of the Pywinauto Recorder application (pywinauto_recorder/app.py).

They need no application under test and no user interaction, so they run in the CI.
"""

import pathlib

import pytest

import pywinauto_recorder
from pywinauto_recorder import app, recorder

PACKAGE_DIR = pathlib.Path(pywinauto_recorder.__file__).resolve().parent


def test_app_has_main():
	"""Tests that the application is started by pywinauto_recorder.app.main."""
	assert callable(app.main)


def test_icons_are_in_the_package():
	"""Tests that the icons are in the package, where the application and the recorder search them."""
	icons_dir = PACKAGE_DIR / "Icons"
	for name in ("IconPyRec.ico", "record.ico", "stop.ico", "play.ico", "folder.ico", "search.ico", "light-on.ico",
	             "settings.ico", "check.ico", "cross.ico", "power.ico", "help.ico", "comments.ico", "favourite.ico",
	             "paste.ico"):
		assert (icons_dir / name).is_file(), name
	assert pathlib.Path(recorder.IconSet.path_icons).resolve() == PACKAGE_DIR


def test_command_pywinauto_recorder():
	"""Tests that installing the package installs the command 'pywinauto-recorder'."""
	from importlib import metadata
	try:
		distribution = metadata.distribution("pywinauto_recorder")
	except metadata.PackageNotFoundError:
		pytest.skip("The package is not installed")
	entry_points = [entry_point for entry_point in distribution.entry_points
	                if entry_point.group == "console_scripts" and entry_point.name == "pywinauto-recorder"]
	assert len(entry_points) == 1
	assert entry_points[0].value == "pywinauto_recorder.app:main"


def test_python_m_pywinauto_recorder():
	"""Tests that the package can be started with 'python -m pywinauto_recorder'."""
	assert (PACKAGE_DIR / "__main__.py").is_file()
	assert "from pywinauto_recorder.app import main" in (PACKAGE_DIR / "__main__.py").read_text()


################################################################################################
#                                     Modes of the recorder                                    #
################################################################################################

@pytest.mark.parametrize("mode, value", [
	(recorder.RecorderMode.INITIALIZING, "Initializing"),
	(recorder.RecorderMode.RECORD, "Record"),
	(recorder.RecorderMode.PLAY, "Play"),
	(recorder.RecorderMode.INFO, "Info"),
	(recorder.RecorderMode.STOP, "Stop"),
	(recorder.RecorderMode.QUIT, "Quit"),
])
def test_recorder_mode_is_equal_to_its_value(mode, value):
	"""Tests that a mode is equal to its former string value, so that code comparing it with a string keeps working."""
	assert isinstance(mode, str)
	assert mode == value and value == mode
	assert mode in (value,)
	assert recorder.RecorderMode(value) is mode


class _FakeRecorder:
	"""Has the attribute used by the 'mode' property of the Recorder, without starting a recorder."""
	_mode = None


@pytest.mark.parametrize("value", ["Record", recorder.RecorderMode.RECORD])
def test_recorder_mode_setter_accepts_a_string_or_a_mode(value):
	"""Tests that the mode of the recorder can be set with a RecorderMode or with its value."""
	fake_recorder = _FakeRecorder()
	recorder.Recorder.mode.fset(fake_recorder, value)
	assert recorder.Recorder.mode.fget(fake_recorder) is recorder.RecorderMode.RECORD


def test_recorder_mode_setter_refuses_an_unknown_mode():
	"""Tests that the mode of the recorder cannot be set with an unknown value."""
	with pytest.raises(ValueError):
		recorder.Recorder.mode.fset(_FakeRecorder(), "Pause")


def test_recorder_mode_is_exported_by_the_package():
	"""Tests that RecorderMode can be imported from the package."""
	assert pywinauto_recorder.RecorderMode is recorder.RecorderMode
