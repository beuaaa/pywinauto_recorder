REM COMPILE EXE
set PYTHON_EXE=%homedrive%%homepath%\AppData\Local\Programs\Python\Python38\python.exe
%PYTHON_EXE% -m nuitka --standalone --mingw64 ..\clear_comtypes_cache.py --show-scons

pause
