# -*- coding: utf-8 -*-

import os
import sys

sys.path.insert(0, os.path.abspath('..'))
sys.path.insert(0, os.path.abspath('../pywinauto_recorder'))


extensions = ['sphinx.ext.autodoc',
              'sphinx.ext.autosummary',
              'sphinx_autodoc_typehints',
              'sphinx.ext.viewcode'
              ]

autodoc_mock_imports = ["pywinauto", "win32api", "win32gui", "win32con", "win32process", "psutil", "threading", "overlay_arrows_and_more",
                        "keyboard", "mouse", "collections", "pyperclip", "codecs", "pytest", "os", "time", "PIL",
                        "cachetools", "thefuzz", "easyocr", "numpy", "OCRWrapper"]

autosummary_generate = True
# autosummary_imported_members = True
autosummary_ignore_module_all = False
templates_path = ['_templates']
exclude_patterns = ['_build', '_templates']
source_suffix = '.rst'
master_doc = 'index'
html_theme = 'sphinx_rtd_theme'


# General information about the project.
project = u'pywinauto_recorder'
# The name of this setting of Sphinx is also the name of a built-in function of Python
copyright = u"2022, David Pratmarty"  # pylint: disable=redefined-builtin
author = u"David Pratmarty"