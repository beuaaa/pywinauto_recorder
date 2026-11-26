Element paths and UIPath
========================

The functions of the player that act on an element, like :func:`~pywinauto_recorder.player.find`,
:func:`~pywinauto_recorder.player.click` or :func:`~pywinauto_recorder.player.move`, identify it with a **path**: the
names and the control types of the element and of its ancestors, from its window. The recorder writes these paths in
the scripts. This page explains all the ways to identify an element with a path: to read a recorded script, to make it
more robust, or to write a script.

The path of an element
----------------------

A path is a list of **entries** joined by ``->``. An entry gives the name and the control type of an element, separated
by ``||``:

.. code-block:: python

    click("Calculator||Window->*->Number pad||Group->One||Button")

- The first entry is the window of the element: here the window named "Calculator", whose control type is "Window".
- The last entry is the element: here the button named "One".
- The entries between them are the ancestors of the element. Here ``*`` replaces some of them (see below).

The names and the control types are the ones of UI Automation, the accessibility technology used by Pywinauto: Button,
Edit, Text, Window, Pane, Group, List, ListItem, MenuItem, and so on. They are compared exactly, with the case and the
spaces:

- an entry without name, like ``||Pane``, designates an element whose name is empty;
- an entry without ``||``, like ``One``, designates an element that has this name, whatever its control type.

To see the path of an element, move the mouse cursor over it while Pywinauto recorder is in display information mode:
the tooltip displays its path. CTRL+SHIFT+f copies the code that finds this element to the clipboard. In a script, a path
is a Python string: the recorder writes the special characters of the names with a backslash, like ``\"`` for a quote or
``\t`` for a tab.

UIPath blocks
-------------

A ``UIPath`` block gives the beginning of the paths used inside it: these paths are relative to the path of the block.
The blocks can be nested. The recorder writes the window, and the part of the paths that the next actions have in
common, in ``UIPath`` blocks:

.. code-block:: python

    from pywinauto_recorder.player import UIPath, click, find

    with UIPath("Calculator||Window"):
        with UIPath("*->Number pad||Group"):
            click("One||Button")  # Calculator||Window->*->Number pad||Group->One||Button
            click("Two||Button")
        find().set_focus()  # the element of the block: the window "Calculator"

- Without path, or with an empty path, a function designates the element of the block: here ``find()`` returns the
  window.
- A path that ends with ``~Absolute_UIPath`` is a full path: the path of the blocks is not added to it. It designates an
  element of another window inside a block, for example the destination of a drag and drop:
  ``drag_and_drop("One||Button", "Other window||Window->*->Destination||List~Absolute_UIPath")``.
- ``UIPath(path, absolute_path=True)`` starts a block whose path replaces the path of the blocks around it.

The windows and the elements found in a block are kept until the end of the block, to make the next searches faster:
the windows of the block are found by its first search, and the element of an inner block is found among the ancestors
of the first element found in this block. An element of the block is then searched among their descendants, instead of
all the windows of the desktop. When it is not found there, it is searched in all the windows of the desktop. So a
window opened after the first search of a block is searched only for the elements that are not in the windows already
found. :func:`~pywinauto_recorder.player.find_all` and the ``[]`` operator always search all the windows. With
``PlayerSettings.keep_block_elements = False``, each search is done in all the windows of the desktop.

Several elements with the same path
-----------------------------------

:func:`~pywinauto_recorder.player.find`, and the functions that use it, need one element. When several elements match
a path, the search is done again until the timeout, because the user interface may be changing, and then a
``FailedSearch`` exception lists the elements found. :func:`~pywinauto_recorder.player.find_all` returns all the
elements that match a path, and :func:`~pywinauto_recorder.player.exists` returns ``None`` instead of raising an
exception when no element is found.

To designate one element among several, add ancestors to the path, use a regular expression, or use the ``[]``
operator.

Wildcards and regular expressions
---------------------------------

``*`` is an entry that replaces any number of entries, zero included:

- ``Calculator||Window->*->One||Button`` designates the button "One" anywhere in the window "Calculator";
- ``*->One||Button``: as the first entry, ``*`` designates any window;
- ``Calculator||Window->*`` designates all the elements of the window, for
  :func:`~pywinauto_recorder.player.find_all`.

An entry that begins with ``RegEx:`` and a space has a name that is a regular expression (see the module re_ of
Python):

- the regular expression must match the beginning of the name: ``RegEx: Display is||Text`` designates the text
  "Display is 3". Add ``$`` to match the whole name: ``RegEx: .* - Notepad$||Window``;
- the control type is still compared exactly: ``RegEx: .*||Button`` designates any button;
- the special characters of regular expressions are escaped with a backslash, which is doubled in a Python string,
  unless the string is a raw string: ``r"RegEx: .* - Notepad\+\+$||Window"``.

A name that changes, like the title of a window that contains the name of a file, is a typical use of a regular
expression.

.. _re: https://docs.python.org/3/library/re.html

The [] operator
---------------

When several elements match a path, the recorder adds the ``[]`` operator at the end of the path, to designate one of
them. It has 2 forms:

.. code-block:: python

    with UIPath("Calculator||Window->*->Number pad||Group"):
        click("RegEx: .*||Button#[0,1]")              # row 0, column 1: "Eight"
        click("RegEx: .*||Button#[Seven||Button,1]")  # column 1, row of "Seven": "Eight"

``#[row,column]``: the elements that match the path are sorted in rows, from the top, and in columns, from the left.
Here, the first row of the number pad is "Seven", "Eight", "Nine". The elements smaller than 8 pixels, and the elements
that are not visible, are not sorted. An element is in the row of the first element of the row when their tops differ
by less than 20 pixels, or by less than their height when all the elements have the same height. The indices change
when the elements move, for example when the window is resized: a reference element is more robust.

``#[path of a reference element,column]``: the element is in the row of a reference element, at the given column. The
path of the reference element must designate one element only: it is relative to the ``UIPath`` blocks too, unless it
ends with ``~Absolute_UIPath``. The recorder writes this form in Smart mode: in a table, the reference element is often
a cell whose text identifies the row.

The ``[]`` operator is always at the end of a path. When only one element matches the path, this element is
returned, and the indices are not checked.

While Pywinauto recorder displays information or records, the element under the mouse cursor is colored green when its
path is unique. Otherwise, the element designated with the ``[]`` operator is colored orange, the other elements that
match the path are colored red, and the reference element is colored blue (see :doc:`pywinauto_recorder_exe`).

The position in the element
---------------------------

``%(dx,dy)`` at the end of a path gives the position of the mouse cursor in the element, for
:func:`~pywinauto_recorder.player.move`, :func:`~pywinauto_recorder.player.click` and the other functions that move the
mouse. ``dx`` and ``dy`` are percentages of the half width and of the half height of the element, from its center:
``%(0,0)`` is the center, ``%(-100,-100)`` the top left corner, and ``%(100,0)`` the middle of the right side. The
recorder writes them when the option "% relative coordinates" of its menu is checked, to click where the user clicked:
``click("Address and search bar||Edit%(-87.50,-20.00)")``. :func:`~pywinauto_recorder.player.find` ignores them. In an
absolute path, ``%(dx,dy)`` is before ``~Absolute_UIPath``.

Other ways to identify an element
---------------------------------

- A Pywinauto wrapper, returned by :func:`~pywinauto_recorder.player.find` for example, can replace a path:
  ``click(wrapper)``.
- Coordinates of the screen, in pixels, can replace a path in the functions that move the mouse: ``move((66, 66))``
  or ``click((66, 66))``.
- :func:`~pywinauto_recorder.player.menu_click` has its own path, made of the names of the menu items:
  ``menu_click("File->New")``.
- :func:`~pywinauto_recorder.player.focus_on_application` restricts the search to the window of an application.
- A path dictionary gives names to the paths, so that the scripts do not repeat them (see :doc:`path_dictionary`).
