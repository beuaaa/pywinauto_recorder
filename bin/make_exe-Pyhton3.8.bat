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
del /Q api-ms-win-core-console-l1-1-0.dll
del /Q api-ms-win-core-datetime-l1-1-0.dll
del /Q api-ms-win-core-debug-l1-1-0.dll
del /Q api-ms-win-core-errorhandling-l1-1-0.dll
del /Q api-ms-win-core-file-l1-1-0.dll
del /Q api-ms-win-core-file-l1-2-0.dll
del /Q api-ms-win-core-file-l2-1-0.dll
del /Q api-ms-win-core-handle-l1-1-0.dll
del /Q api-ms-win-core-heap-l1-1-0.dll
del /Q api-ms-win-core-interlocked-l1-1-0.dll
del /Q api-ms-win-core-libraryloader-l1-1-0.dll
del /Q api-ms-win-core-localization-l1-2-0.dll
del /Q api-ms-win-core-memory-l1-1-0.dll
del /Q api-ms-win-core-namedpipe-l1-1-0.dll
del /Q api-ms-win-core-processenvironment-l1-1-0.dll
del /Q api-ms-win-core-processthreads-l1-1-0.dll
del /Q api-ms-win-core-processthreads-l1-1-1.dll
del /Q api-ms-win-core-profile-l1-1-0.dll
del /Q api-ms-win-core-rtlsupport-l1-1-0.dll
del /Q api-ms-win-core-string-l1-1-0.dll
del /Q api-ms-win-core-synch-l1-1-0.dll
del /Q api-ms-win-core-synch-l1-2-0.dll
del /Q api-ms-win-core-sysinfo-l1-1-0.dll
del /Q api-ms-win-core-timezone-l1-1-0.dll
del /Q api-ms-win-core-util-l1-1-0.dll
del /Q api-ms-win-crt-conio-l1-1-0.dll
del /Q api-ms-win-crt-convert-l1-1-0.dll
del /Q api-ms-win-crt-environment-l1-1-0.dll
del /Q api-ms-win-crt-filesystem-l1-1-0.dll
del /Q api-ms-win-crt-heap-l1-1-0.dll
del /Q api-ms-win-crt-locale-l1-1-0.dll
del /Q api-ms-win-crt-math-l1-1-0.dll
del /Q api-ms-win-crt-multibyte-l1-1-0.dll
del /Q api-ms-win-crt-process-l1-1-0.dll
del /Q api-ms-win-crt-runtime-l1-1-0.dll
del /Q api-ms-win-crt-stdio-l1-1-0.dll
del /Q api-ms-win-crt-string-l1-1-0.dll
del /Q api-ms-win-crt-time-l1-1-0.dll
del /Q api-ms-win-crt-utility-l1-1-0.dll
del /Q api-ms-win-eventing-provider-l1-1-0.dll
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

REM Add pywinauto_recorder\*.png in pywinauto_recorder.dist
MKDIR  pywinauto_recorder
xcopy /y ..\..\pywinauto_recorder\*.png .\pywinauto_recorder

cd ..

REM ZIP pywinauto_recorder.dist
del /Q pywinauto_recorder.dist.7z
"C:\Program Files\7-Zip\7z.exe" a pywinauto_recorder.dist.7z pywinauto_recorder.dist

REM Create sfx
copy /b 7zsd_All_x64.sfx + config_sfx.txt + pywinauto_recorder.dist.7z pywinauto_recorder_original.exe
del /Q pywinauto_recorder.dist.7z

REM Update icons
"C:\Program Files (x86)\Resource Hacker\ResourceHacker.exe" -script icon_script.txt
del /Q pywinauto_recorder_original.exe

pause