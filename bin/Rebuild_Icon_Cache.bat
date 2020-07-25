:: Created by: Shawn Brink
:: http://www.sevenforums.com
:: Tutorial:  http://www.sevenforums.com/tutorials/49819-icon-cache-rebuild.html

@echo off
set iconcache=%localappdata%\IconCache.db

echo The Explorer process must be killed to delete the Icon DB. 
echo.
echo Please SAVE ALL OPEN WORK before continuing.
echo.
pause
echo.
If exist "%iconcache%" goto delID
echo.
echo Icon DB has already been deleted. 
echo.
pause
exit /B

:delID
echo Attempting to delete Icon DB...
echo.
ie4uinit.exe -ClearIconCache
taskkill /IM explorer.exe /F 
del "%iconcache%" /A
del "%localappdata%\Microsoft\Windows\Explorer\iconcache*" /A 
echo.
echo Icon DB has been successfully deleted. Please "restart your PC" now to rebuild your icon cache.
echo.
start explorer.exe
pause
exit /B