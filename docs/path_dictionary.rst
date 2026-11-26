Path dictionary
===============

A path dictionary gives names, the **keys**, to the paths of the elements. A script uses the keys instead of the paths:
the paths are written once, in the dictionary. When the application changes, only the dictionary is changed, not the
scripts. For example, with the same key file and another definition file, the same scripts automate another version of
the application, or the application in another language.

The dictionary files
--------------------

A path dictionary is made of 2 text files that have the same lines:

- the **key file** (``.key``) gives the key of an element on each line;
- the **definition file** (``.def``) gives, on the same line, the **definition** of this element: its path relative to
  its parent, written like in a script.

The tabs at the beginning of the lines of the definition file give the hierarchy of the elements: the parent of an
element is on the last line above it that has one tab less. The key file is indented the same way, so that the 2 files
can be compared line by line, but its indentation is ignored.

For example, `Calculator.key`_ and `Calculator.def`_ describe the number pad and the operators of the Calculator of
Windows:

.. list-table::
    :header-rows: 1
    :widths: 35 65

    * - Calculator.key
      - Calculator.def
    * - .. code-block:: text

            Calculator
                Number pad
                    0
                    1
                    2
                    3
                    4
                    5
                    6
                    7
                    8
                    9
                Standard operators
                    /
                    *
                    -
                    +
                    =

      - .. code-block:: text

            Calculator||Window->*
                Number pad||Group
                    Zero||Button
                    One||Button
                    Two||Button
                    Three||Button
                    Four||Button
                    Five||Button
                    Six||Button
                    Seven||Button
                    Eight||Button
                    Nine||Button
                Standard operators||Group
                    Divide by||Button
                    Multiply by||Button
                    Minus||Button
                    Plus||Button
                    Equals||Button

.. note::  The indentation of the definition file must be made of tab characters, one tab per level. This page displays
    the tabs as spaces: to start from files that have the tabs, download `Calculator.key`_ and `Calculator.def`_.

Using a dictionary in a script
------------------------------

:func:`~pywinauto_recorder.player.load_dictionary` loads the 2 files. Then
:func:`~pywinauto_recorder.player.shortcut` returns the definition of a key, to use in nested ``UIPath`` blocks that
follow the hierarchy of the dictionary:

.. code-block:: python

    from pywinauto_recorder.player import load_dictionary, shortcut, UIPath, click

    load_dictionary("Calculator.key", "Calculator.def")
    with UIPath(shortcut("Calculator")):
        with UIPath(shortcut("Number pad")):
            click(shortcut("1"))
        with UIPath(shortcut("Standard operators")):
            click(shortcut("+"))
        with UIPath(shortcut("Number pad")):
            click(shortcut("2"))
        with UIPath(shortcut("Standard operators")):
            click(shortcut("="))

Here ``shortcut("1")`` returns ``"One||Button"``, so the first click is on the element
``"Calculator||Window->*->Number pad||Group->One||Button"``.

:func:`~pywinauto_recorder.player.full_definition` returns the full path of a key: the definitions of its parents and
its definition, joined by ``->``. It does not need ``UIPath`` blocks:

.. code-block:: python

    from pywinauto_recorder.player import load_dictionary, full_definition, click

    load_dictionary("Calculator.key", "Calculator.def")
    click(full_definition("1"))  # Calculator||Window->*->Number pad||Group->One||Button
    click(full_definition("+"))
    click(full_definition("2"))
    click(full_definition("="))

Writing a dictionary
--------------------

The definitions are the entries of the paths written by the recorder, for example in a recorded script, or in the code
that CTRL+SHIFT+f copies for the element under the mouse cursor: each entry, or group of entries joined by ``->``, goes
on its own line, indented under the line of its parent.

The files follow these rules:

- The 2 files have the same number of lines: line *n* of the key file is the key of line *n* of the definition file.
  The empty lines are ignored, but a line must be empty in both files.
- The indentation of the definition file is made of tabs, one tab per level, and a line is at most one level deeper
  than the line above it. The spaces and the tabs at the end of the lines are ignored.
- A definition is written like a path in a script, but without the quotes, and without the backslash before ``\``
  and ``"``. It can use regular expressions, like ``RegEx: .* - Notepad||Window``, and ``*``. The ``[]`` operator,
  like ``#[0,1]``, must be at the end of a path: only the definition of an element without children can use it.
- A key or a definition cannot contain a tab, because the tabs separate the levels. For an element whose name contains
  a tab, use a regular expression, for example ``RegEx: Stop recording.*||MenuItem``.
- The files are read with the UTF-8 encoding, with or without a byte order mark. The parameter ``encoding`` of
  :func:`~pywinauto_recorder.player.load_dictionary` gives another encoding, for example ``"cp1252"``.
- The names of the files are relative to the current folder. When Pywinauto recorder replays a script file, the
  current folder is the folder of the script.

:func:`~pywinauto_recorder.player.load_dictionary` raises a ``ValueError`` when the files do not follow these rules:
its message gives the file and the line.

Several dictionaries can be loaded: their keys are added to the same dictionary, and a key that is loaded again gets its
new definition. :func:`~pywinauto_recorder.player.shortcut` and :func:`~pywinauto_recorder.player.full_definition` raise
a ``KeyError`` for an unknown key.

.. _Calculator.key: https://raw.githubusercontent.com/beuaaa/pywinauto_recorder/master/tests/Calculator.key
.. _Calculator.def: https://raw.githubusercontent.com/beuaaa/pywinauto_recorder/master/tests/Calculator.def
