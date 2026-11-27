#!/usr/bin/env python
# -*- coding: utf-8 -*-

# print('__file__={0:<35} | __name__={1:<20} | __package__={2:<20}'.format(__file__, __name__, str(__package__)))

"""
Pywinauto recorder records user interface actions and saves them in a Python script.

The generated Python script plays back user interface actions in the order in which the user recorded them.

Pywinauto recorder uses accessibility technologies via the Pywinauto_ library.
"""

__version__ = "1.0.0.dev0"

import os

import pywinauto as _pywinauto


def _check_pywinauto_version(pywinauto_version, pywinauto_folder=''):
	"""
	Raises an ImportError if the installed pywinauto is older than pywinauto 1.0.

	Pywinauto recorder searches the elements with the keywords of pywinauto 1.0 (name, name_re...). With pywinauto 0.6,
	each search raised a TypeError, and the player searched again and again until the timeout.

	:param pywinauto_version: The version of pywinauto, pywinauto.__version__
	:param pywinauto_folder: The folder of pywinauto, given in the message
	"""
	major = str(pywinauto_version).split('.', maxsplit=1)[0]
	if major.isdigit() and int(major) < 1:
		raise ImportError(
			"Pywinauto recorder {} needs pywinauto 1.0 or later, but pywinauto {} is installed{}. Install Pywinauto "
			"recorder with pip, which installs pywinauto 1.0, preferably in a virtual environment.".format(
				__version__, pywinauto_version, " in " + pywinauto_folder if pywinauto_folder else ""))


_PYWINAUTO_FILE = getattr(_pywinauto, '__file__', None)
_check_pywinauto_version(
	getattr(_pywinauto, '__version__', ''), os.path.dirname(_PYWINAUTO_FILE) if isinstance(_PYWINAUTO_FILE, str) else '')

# The API of the package: the functions of the player, and the recorder
from .player import *  # noqa: F401,F403
from .recorder import Recorder, RecorderMode  # noqa: F401


