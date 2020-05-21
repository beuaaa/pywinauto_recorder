#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
    Pywinauto recorder records user interface actions and saves them in a Python script.
    The generated Python script plays back user interface actions in the order in which the user recorded them.

    Pywinauto recorder uses accessibility technologies via the Pywinauto_ library.
"""

__version__ = "0.1.0"

from recorder import Recorder
from player import *
