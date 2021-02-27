# -*- coding: utf-8 -*-

import os
import sys

sys.path.insert(0, os.path.abspath('..'))
sys.path.insert(0, os.path.abspath('../pywinauto_recorder'))

extensions = ['sphinx.ext.autodoc', 'sphinx.ext.autosummary', ]

autosummary_generate = True

templates_path = ['_templates']

master_doc = 'index'

