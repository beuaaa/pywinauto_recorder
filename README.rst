
******************
Pywinauto recorder
******************

**WARNING:**
This recorder is still at a very early stage of development.


Description
###########
The "Pywinauto recorder" records user interface actions and saves them in a Python script.
The generated Python script plays back user interface actions in the order in which the user recorded them.

The "Pywinauto recorder" uses accessibility technologies via the Pywinauto library.

Installation
############
 pip install overlay_arrows_and_more

 Download https://github.com/beuaaa/pywinauto_recorder/archive/master.zip

 Unzip master.zip

Usage
#####

In the folder pywinauto_recoder you will find Recorder.bat. Modify the path of the Python interpreter if necessary.

Recorder.bat:

.. code-block:: bat

    if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit
    python.exe recorder.py

Double click on Recorder.bat to start the recorder.

When the recorder is started, it is in "Pause" mode.
Press ALT+r to switch to "Recording" mode.
If the item below the mouse cursor can be uniquely identified, it will turn green.
The user can then click or perform another action on the user interface and the action is recorded in the generated Python script.
Perform a few actions on the user interface to finish.
Press ALT+r to return to "Pause" mode.
Eventually, press ALT+q to exit the recoder.
The generated Python script is generated in the "Record files" folder.
To replay a file, drag and drop it to Drag_n_drop_to_replay.bat.


Functions
**********************

To be completed