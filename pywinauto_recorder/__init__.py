#!/usr/bin/env python
# -*- coding: utf-8 -*-

# print('__file__={0:<35} | __name__={1:<20} | __package__={2:<20}'.format(__file__, __name__, str(__package__)))

"""
    Pywinauto recorder records user interface actions and saves them in a Python script.
    The generated Python script plays back user interface actions in the order in which the user recorded them.

    Pywinauto recorder uses accessibility technologies via the Pywinauto_ library.
"""

__version__ = "0.6.6"

from .player import *
from .recorder import Recorder


