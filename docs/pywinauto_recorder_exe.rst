Pywinauto_recorder.exe for Windows
==================================
"pywinauto_recorder.exe_" is a standalone application, it's the compiled version for Windows.

.. _pywinauto_recorder.exe: https://raw.githubusercontent.com/beuaaa/pywinauto_recorder/master/bin/pywinauto_recorder.exe

.. image:: https://raw.githubusercontent.com/beuaaa/pywinauto_recorder/master/Images/Download.png?sanitize=true
            :target: https://raw.githubusercontent.com/beuaaa/pywinauto_recorder/master/bin/pywinauto_recorder.exe
            :width: 200 px
            :align: center

Usage
-----
- Double click on "pywinauto_recorder.exe_" to start the recorder.
- When the recorder is started, it is in "Pause" mode.
- Press CTRL+SHIFT+f to copy the code that finds the element colored green or orange to the clipboard.
- Press CTRL+ALT+r to switch to "Record" mode.
- If the element below the mouse cursor can be uniquely identified, it will turn green or orange.
- You can then click or perform another action on the user interface and it will be recorded in the generated Python script.
- Repeat this process performing a few actions on the user interface and when you're done press CTRL+ALT+r to return to end recording.
- Eventually, press CTRL+ALT+q to exit the recorder.
- The generated Python script is saved in the "Record files" folder and copied in the clipboard.
- To replay a Python script, you can drag and drop it to "pywinauto_recorder.exe_"

.. image:: https://raw.githubusercontent.com/beuaaa/pywinauto_recorder/master/Images/Tutorial1.png?sanitize=true
            :target: https://www.youtube.com/watch?v=-7W-rOvUjdE&list=PLV9GWm_Y6wCAW5aaDGSNJwwenmRk6dI-u
            :width: 1207 px
            :align: center
Icons
-----
Two transparent icons are displayed at the top left of the screen:
 - the first icon corresponds to Record/Pause mode. Press CTRL+ALT+r to switch.
 - the second icon displays a green bar at each iteration of the loop. It allows you to see how fast the loop is running.

More explanations
^^^^^^^^^^^^^^^^^
The main of "Pywinauto recorder" is an infinite loop where at each iteration it:
 (1) finds the path of the element under the mouse cursor. The path is formed by the window_text and control_type pair of the element and all its ancestors.
 (2) searches for an unambiguous path, if found, it colors the element region green or orange.
 (3) records a user action in a file involving the last recognized unique path.

.. note::  To reflect the position of the mouse cursor as closely as possible, an offset is added to the user actions recorded in the generated Python script. This offset is proportional to the size of the element and relative to the center of the element.

If the path of the element under the mouse cursor is not ambiguous, the region of the element is colored green. Otherwise two strategies are used to try to disambiguate the path in the following order:
 (1) All elements having the same path are ordered in a 2D array. The path of the element region under the mouse cursor is disambiguated by adding a row index and a column index so that it is colored orange. The other element regions are colored red
 (2) An element whose path is unambiguous is searched on the same line on the left, if found its region is colored blue and the element under the mouse cursor is colored orange.
