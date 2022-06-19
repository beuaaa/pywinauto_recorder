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

autodoc_mock_imports = ["pywinauto", "win32api", "win32gui", "win32con", "threading", "overlay_arrows_and_more",
                        "keyboard", "mouse", "collections", "pyperclip", "codecs", "pytest", "os", "time", "PIL"]

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
copyright = u"2021, David Pratmarty"
author = u"David Pratmarty"