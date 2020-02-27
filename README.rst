
******************
Pywinauto recorder
******************

**WARNING:**
This recorder is still at a very early stage of development.


Description
###########
"Pywinauto recorder" records user interface actions and saves them in a Python script.
The generated Python script plays back user interface actions in the order in which the user recorded them.

"Pywinauto recorder" uses accessibility technologies via the Pywinauto_ library.

.. _Pywinauto: https://github.com/pywinauto/pywinauto/

Installation
############
 pip install pywinauto

 pip install keyboard

 pip install mouse

 pip install overlay_arrows_and_more

 Download https://github.com/beuaaa/pywinauto_recorder/archive/master.zip

 Unzip master.zip

Usage
#####

In the folder pywinauto_recorder you will find "Recorder.bat". Modify the path of the Python interpreter if necessary.

Recorder.bat:

.. code-block:: bat

    if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit
    python.exe recorder.py

- Double click on Recorder.bat to start the recorder.
- When the recorder is started, it is in "Pause" mode.
- Press ALT+r to switch to "Recording" mode.
- If the item below the mouse cursor can be uniquely identified, it will turn green.
- You can then click or perform another action on the user interface and the action is recorded in the generated Python script.
- Repeat this process performing a few actions on the user interface and when you're done press ALT+r to return to "Pause" mode.
- Eventually, press ALT+q to exit the recorder.
- The generated Python script is saved in the "Record files" folder.
- To replay actions of a Python script, you can drag and drop it to Drag_n_drop_to_replay.bat. Modify the path of the Python interpreter if necessary.

More explanations
#################

The main of "Pywinauto recorder" is an infinite loop where, at each iteration, it:
 - finds the path of the element under the mouse cursor
 - if this path is unique (unambiguous), it greens the region of the element
 - records a user action in a file involving the last recognized unique path
 - the icon, after the Recording/Pause mode icon, displays a green bar at each iteration of the loop. It allows you to see how fast the loop is running.

.. note:: When it cannot find a unique path for an element it reds all the elements with the same path and it searchs a unique path in the ancestors then it greens the ancestor.
The mouse coordinates recorded are relative to the center of the element recognized with a unique path

Functions
**********************

To be completed
