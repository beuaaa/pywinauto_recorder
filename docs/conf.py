# -*- coding: utf-8 -*-

import os
import sys

sys.path.insert(0, os.path.abspath('..'))


import pywinauto_recorder

extensions = ['sphinx.ext.autodoc', 'sphinx.ext.autosummary', ]


autosummary_generate = True

templates_path = ['_templates']

source_suffix = '.rst'

master_doc = 'index'


# General information about the project.
project = u'pywinauto_recorder'
copyright = u"2021, David Pratmarty"
author = u"David Pratmarty"