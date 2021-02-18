REM COMPILE EXE
C:\Users\oktalse\AppData\Local\Programs\Python\Python38\python.exe -m nuitka --standalone --mingw64 --windows-dependency-tool=peffile ..\pywinauto_recorder.py --show-progress --show-scons

REM CLEAN pywinauto_recorder.dist
cd pywinauto_recorder.dist

del /Q win32file.pyd
del /Q unicodedata.pyd
del /Q tcl86t.dll
del /Q tk86t.dll
del /Q sqlite3.dll
del /Q pyexpat.pyd
del /Q comctl32.dll
del /Q _bz2.pyd
del /Q _tkinter.pyd
del /Q _ssl.pyd
del /Q _sqlite3.pyd
del /Q _msi.pyd
del /Q _hashlib.pyd
del /Q _elementtree.pyd
del /Q gdiplus.dll
del /Q libcrypto-1_1.dll
del /Q libssl-1_1.dll
del /Q ucrtbase.dll
del /Q vcruntime140.dll
del /Q _asyncio.pyd
del /Q _decimal.pyd
del /Q _lzma.pyd
del /Q _overlapped.pyd
del /Q _queue.pyd
rmdir certifi /s /q
rmdir numpy /s /q
rmdir PIL /s /q
rmdir win32com /s /q


REM Add pywinauto_recorder\*.png in pywinauto_recorder.dist
MKDIR  pywinauto_recorder
xcopy /y ..\..\pywinauto_recorder\*.png .\pywinauto_recorder
MKDIR  Icons
xcopy /y ..\..\Icons\*.ico .\Icons

cd ..


REM Update icons
RENAME .\pywinauto_recorder.dist\pywinauto_recorder.exe pywinauto_recorder_original.exe
"C:\Program Files (x86)\Resource Hacker\ResourceHacker.exe" -script icon_script_in_dist.txt
DEL .\pywinauto_recorder.dist\pywinauto_recorder_original.exe


REM Make installer
"C:\Program Files (x86)\NSIS\Bin\makensis.exe" Pywinauto_recorder.nsi

pause