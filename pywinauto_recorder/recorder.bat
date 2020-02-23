if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit

..\Python\Pyportable-2.7.10rc1\python.exe recorder.py 


exit 