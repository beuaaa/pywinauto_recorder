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
