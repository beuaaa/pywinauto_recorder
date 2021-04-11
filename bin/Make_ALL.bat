REM COMPILE EXE
C:\Users\oktalse\AppData\Local\Programs\Python\Python38\python.exe -m nuitka --standalone --mingw64 ..\pywinauto_recorder.py

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


REM Copy Icons\*.ico in pywinauto_recorder.dist
MKDIR  Icons
xcopy /y ..\..\pywinauto_recorder\Icons\*.ico .\Icons


REM Copy clear_comtypes_cache.exe
xcopy /y ..\clear_comtypes_cache.dist\*.exe .


cd ..

REM Create VersionInfo.rc 
for /f "tokens=2 delims==" %%a in ('type ..\pywinauto_recorder\__init__.py^|find "__version__ = "') do (
  set version=%%a & goto :continue
)
:continue
set version=%version:~1%
echo %version%

setlocal ENABLEDELAYEDEXPANSION
set word=^,
set str=%version:.=!word!%
set str=%str:" =%
set str=%str:"=%

@echo off
setlocal EnableDelayedExpansion
(
echo(
echo 1 VERSIONINFO
echo FILEVERSION 0^,2^,0^,0
echo PRODUCTVERSION %str%^,0
echo FILEOS 0x40004
echo FILETYPE 0x1
echo {
echo BLOCK "StringFileInfo"
echo {
echo 	BLOCK "000004B0"
echo 	{
echo 		VALUE "ProductName", "Pywinauto Recorder"
echo 		VALUE "ProductVersion", %version%
echo 	}
echo }
echo(
echo BLOCK "VarFileInfo"
echo {
echo 	VALUE "Translation", 0x0000 0x04B0  
echo }
echo }
) 1>"VersionInfo.rc"
@echo on

REM Compile VersioInfo.rc 
"C:\Program Files (x86)\Resource Hacker\ResourceHacker.exe" -open .\VersionInfo.rc -save .\VersionInfo.res -action compile -log con

REM Update pywinauto_recorder.exe with icons and version info
RENAME .\pywinauto_recorder.dist\pywinauto_recorder.exe pywinauto_recorder_original.exe
"C:\Program Files (x86)\Resource Hacker\ResourceHacker.exe" -script icon_script_in_dist.txt
DEL .\pywinauto_recorder.dist\pywinauto_recorder_original.exe
DEL VersionInfo.res

REM Make installer
"C:\Program Files (x86)\NSIS\Bin\makensis.exe" Pywinauto_recorder.nsi

pause