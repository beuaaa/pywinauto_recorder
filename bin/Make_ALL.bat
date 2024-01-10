REM Get version from __init__.py
for /f "tokens=2 delims==" %%a in ('type ..\pywinauto_recorder\__init__.py^|find "__version__ = "') do (
  set version=%%a & goto :continue
)
:continue
set version=%version:~1%
echo %version%

REM COMPILE EXE
set PYTHON_EXE=%homedrive%%homepath%\AppData\Local\Programs\Python\Python310\python.exe
set CMD_ICON=--windows-icon-from-ico=..\pywinauto_recorder\Icons\IconPyRec.ico
set CMD_INFO=--windows-file-version=%version% --windows-product-version=%version% --windows-product-name="Pywinauto Recorder" --windows-company-name="Pywinauto Recorder"
%PYTHON_EXE% -m nuitka --standalone --mingw64 %CMD_ICON% %CMD_INFO% ..\pywinauto_recorder.py --nofollow-import-to=*.ocr_wrapper.py
REM %PYTHON_EXE% -m nuitka --standalone --mingw64 ..\clear_comtypes_cache.py --show-scons


REM CLEAN pywinauto_recorder.dist
cd pywinauto_recorder.dist

del /Q win32file.pyd
del /Q tcl86t.dll
del /Q tk86t.dll
del /Q sqlite3.dll
del /Q pyexpat.pyd
del /Q comctl32.dll
del /Q _bz2.pyd
del /Q _tkinter.pyd
del /Q _sqlite3.pyd
del /Q _msi.pyd
del /Q _hashlib.pyd
del /Q _elementtree.pyd
del /Q gdiplus.dll
del /Q ucrtbase.dll
del /Q _asyncio.pyd
del /Q _decimal.pyd
del /Q _lzma.pyd
del /Q _overlapped.pyd
del /Q _queue.pyd
rmdir certifi /s /q
rmdir numpy /s /q
rmdir PIL /s /q
rmdir win32com /s /q


REM Copy Icons\*.ico in pywinauto_recorder.dist
MKDIR  Icons
xcopy /y ..\..\pywinauto_recorder\Icons\*.ico .\Icons


REM Copy clear_comtypes_cache.exe in pywinauto_recorder.dist
xcopy /y ..\clear_comtypes_cache.dist\*.exe .


cd ..


REM Make installer
"C:\Program Files (x86)\NSIS\Bin\makensis.exe" Pywinauto_recorder.nsi

pause