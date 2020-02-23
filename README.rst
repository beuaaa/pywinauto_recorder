
******************
Pywinauto recorder
******************

**WARNING:**
This recorder is still at a very early stage of development.

Description
###########
The 'Pywinauto recorder' paints in green the element in under the mouse cursor if it has a unique path.
Then the user can do an action and it is recorded if the recording mode is activated.
To activate/deactivate the recording mode: ALT+R

Installation
############
 TODO: pip install pywinauto_recorder


Usage
#####

Recorder.bat:

.. code-block:: bat

    if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit
    python.exe recorder.py


Functions
**********************

To be completed