Pywinauto_recorder for Windows
==============================

- "pywinauto_recorder.exe_" is a standalone application, it's the compiled version of "pywinauto_recorder.py" for 64-bit Windows.

.. _pywinauto_recorder.exe: https://raw.githubusercontent.com/beuaaa/pywinauto_recorder/master/bin/Pywinauto_recorder_installer.exe

.. image:: https://raw.githubusercontent.com/beuaaa/pywinauto_recorder/master/Images/Download.png?sanitize=true
            :target: https://raw.githubusercontent.com/beuaaa/pywinauto_recorder/master/bin/Pywinauto_recorder_installer.exe
            :width: 200 px
            :align: center
            :alt: Download installer

- "pywinauto_recorder.py_" is the main source, you will find the Pyhon code downloading the github clone.

.. _pywinauto_recorder.py: https://github.com/beuaaa/pywinauto_recorder/archive/master.zip

.. image:: https://raw.githubusercontent.com/beuaaa/pywinauto_recorder/master/Images/Download.png?sanitize=true
            :target: https://github.com/beuaaa/pywinauto_recorder/archive/master.zip
            :width: 200 px
            :align: center
            :alt: Download GitHub clone

Usage
-----
- Double click on "pywinauto_recorder.exe_" or run "python.exe pywinauto_recorder.py_" to start the recorder.
- The recorder is started in display information mode, a tray  icon is added in the right-side of the Windows Taskbar.
- Press CTRL+SHIFT+f to copy the code that finds the element colored green or orange to the clipboard.
- Press CTRL+ALT+r to switch to "Record" mode.
- If the element below the mouse cursor can be uniquely identified, it will turn green or orange.
- You can then click or perform another action on the user interface and it will be recorded in the generated Python script.
- Repeat this process performing a few actions on the user interface and when you're done press CTRL+ALT+r to end recording.
- The generated Python script is saved in the "Pywinauto recorder" folder in your home folder and copied in the clipboard.
- Click on "Quit" in the tray menu.
- To replay a Python script, you can drag and drop it to "pywinauto_recorder.exe_"

.. raw:: html

    <div style="position: relative; padding-bottom: 56.25%; height: 0; overflow: hidden; max-width: 100%; height: auto;">
        <iframe src="https://www.youtube.com/embed/GQiT5w4dzgE" frameborder="0" allowfullscreen style="position: absolute; top: 0; left: 0; width: 100%; height: 100%;"></iframe>
    </div>

.. warning::  Pywinauto_recorder_exe does not work on all PCs. If it doesn't work use the Python version. Help me to find a solution by filling this form_.

.. _form: https://docs.google.com/forms/d/e/1FAIpQLSdvxXJCYfoFUaTVHCDzGxkMbc_8qq68pTb_7hPbaPUDyYlOeQ/viewform


Icons
-----
Some transparent icons are displayed at the top left of the screen:
 - an icon corresponds to Record/Stop mode. Press CTRL+ALT+r to switch.
 - a magnifiying glass icon cooresponds to Display information mode.
 - a bulb icon corresponds to Smart mode. Press CTRL+ALT+S to activate it.
 - another icon displays a green bar at each iteration of the loop. It allows you to see how fast the loop is running.

More explanations
^^^^^^^^^^^^^^^^^
The main of "Pywinauto recorder" is an infinite loop where at each iteration it:
 (1) finds the path of the element under the mouse cursor. The path is formed by the window_text and control_type pair of the element and all its ancestors.
 (2) searches for an unambiguous path, if found, it colors the element region green or orange.
 (3) records a user action in a file involving the last recognized unique path.

.. note::  To reflect the position of the mouse cursor as closely as possible, an offset is added to the user actions recorded in the generated Python script. This offset is proportional to the size of the element and relative to the center of the element.

If the path of the element under the mouse cursor is not ambiguous, the region of the element is colored green. Otherwise two strategies are used to try to disambiguate the path in the following order:
 (1) All elements having the same path are ordered in a 2D array. The path of the element region under the mouse cursor is disambiguated by adding a row index and a column index so that it is colored orange. The other element regions are colored red
 (2) When Smart mode is enabled, an element whose path is unambiguous is searched on the same line on the left, if found its region is colored blue and the element under the mouse cursor is colored orange.
