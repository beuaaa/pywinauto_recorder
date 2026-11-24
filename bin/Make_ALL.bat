@echo off
REM Builds Pywinauto Recorder with Nuitka and creates its installer: bin\Pywinauto_recorder_installer.exe
REM
REM Nuitka creates the installer with NSIS. NSIS does not need to be installed: Nuitka downloads it
REM if makensis.exe is not found in PATH. Nuitka also downloads the MinGW64 C compiler if needed.
REM
REM The build uses its own Python environment, bin\venv (ignored by git), created with BUILD_PYTHON:
REM    set BUILD_PYTHON=py -3.12      (default)
REM With Python 3.13 or later, MinGW64 does not work: install "Visual Studio Build Tools" (MSVC)
REM and remove --mingw64 below.

cd /d "%~dp0"
if not defined BUILD_PYTHON set "BUILD_PYTHON=py -3.12"

REM Get version from __init__.py
for /f "tokens=2 delims==" %%a in ('type ..\pywinauto_recorder\__init__.py ^| find "__version__ = "') do set "version=%%~a"
set "version=%version: =%"
set "version=%version:"=%"
REM Nuitka only accepts numbers in the version information: "1.0.0.dev0" -> "1.0.0"
for /f "tokens=1 delims=abcdefghijklmnopqrstuvwxyz" %%v in ("%version%") do set "numeric_version=%%v"
if "%numeric_version:~-1%"=="." set "numeric_version=%numeric_version:~0,-1%"
echo Pywinauto Recorder %version% (version of the executable: %numeric_version%)

REM BUILD ENVIRONMENT
if not exist venv\Scripts\python.exe %BUILD_PYTHON% -m venv venv
if not exist venv\Scripts\python.exe (echo Cannot create the build environment with "%BUILD_PYTHON%". & pause & exit /b 1)
set "PYTHON_EXE=venv\Scripts\python.exe"
%PYTHON_EXE% -m pip install --upgrade pip
%PYTHON_EXE% -m pip install --upgrade .. nuitka || (echo Cannot install the dependencies. & pause & exit /b 1)

REM COMPILE EXE AND CREATE INSTALLER
set OPTIONS=--mode=standalone --mingw64 --assume-yes-for-downloads --windows-console-mode=force
set OPTIONS=%OPTIONS% --windows-icon-from-ico=..\pywinauto_recorder\Icons\IconPyRec.ico
set OPTIONS=%OPTIONS% --product-name="Pywinauto Recorder" --company-name="Pywinauto Recorder"
set OPTIONS=%OPTIONS% --file-version=%numeric_version% --product-version=%numeric_version%
set OPTIONS=%OPTIONS% --file-description="Pywinauto Recorder %version%"
REM The icons are searched in <folder of pywinauto_recorder.exe>\Icons
set OPTIONS=%OPTIONS% --include-data-dir=..\pywinauto_recorder\Icons=Icons
REM Packages that Pywinauto Recorder does not need (formerly deleted after the build)
set OPTIONS=%OPTIONS% --nofollow-import-to=numpy --nofollow-import-to=PIL --nofollow-import-to=tkinter
set OPTIONS=%OPTIONS% --nofollow-import-to=certifi --nofollow-import-to=win32com
REM comtypes imports some of its modules by name (importlib) when it generates the COM code, and this
REM generated code imports other comtypes modules: Nuitka cannot see these imports, so the whole
REM comtypes package is included, except its tests ("No module named 'comtypes.stream'" otherwise)
set OPTIONS=%OPTIONS% --include-package=comtypes --nofollow-import-to=comtypes.test
REM Installer: per-user install without administrator rights, like the former NSIS script
set OPTIONS=%OPTIONS% --windows-create-installer --windows-installer-output=Pywinauto_recorder_installer.exe
set OPTIONS=%OPTIONS% --windows-installer-mode=user
set OPTIONS=%OPTIONS% --windows-installer-install-dir="$LOCALAPPDATA\Programs\Pywinauto recorder"
set OPTIONS=%OPTIONS% --windows-installer-shortcuts=start-menu,desktop

%PYTHON_EXE% -m nuitka %OPTIONS% ..\pywinauto_recorder.py || (echo The build failed. & pause & exit /b 1)

echo.
echo Installer created: %cd%\Pywinauto_recorder_installer.exe
pause
