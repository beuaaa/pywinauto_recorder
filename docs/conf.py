# -*- coding: utf-8 -*-

import os
import sys

sys.path.insert(0, os.path.abspath('../'))
sys.path.insert(0, os.path.abspath('../../'))
sys.path.insert(0, os.path.abspath('../pywinauto_recorder/'))

extensions = ['sphinx.ext.autodoc', 'sphinx.ext.autosummary', 'sphinxcontrib.apidoc',]

apidoc_module_dir = '../pywinauto_recorder'
apidoc_output_dir = 'reference/api'
apidoc_excluded_paths = ['tests']
apidoc_separate_modules = True

autosummary_generate = True

templates_path = ['_templates']

master_doc = 'index'

