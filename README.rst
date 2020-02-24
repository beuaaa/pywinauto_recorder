
******************
Pywinauto recorder
******************

**WARNING:**
This recorder is still at a very early stage of development.


Description
###########
The 'Pywinauto recorder' records UI actions and save them in a Python script.
The generated Python script allows to playback the UI actions in the respective order that user recorded them.

When the recorder is started, it is in Pause mode. To enter in Recording mode you have to press ATL+r.
Then the recorder generates a path to the element under the mouse cursor, if this path is unique it
paints in green the element. Then the user can click or do another action on the UI.
To stop recording the user have to press ALT+r

'Pywinauto recorder' uses accessibility technologies via Pywinauto library


Installation
############
 pip install overlay_arrows_and_more
 Download https://github.com/beuaaa/pywinauto_recorder/archive/master.zip


Usage
#####

Unzip master.zip
In the folder pywinauto_recoder you will find Recorder.bat. Modify the path of the Python interpreter if necessary.

.. code-block:: bat

    if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit
    python.exe recorder.py

Double click on Recorer.bat to start the recorder.


Functions
**********************

To be completed