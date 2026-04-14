@echo off
echo Installing dependencies...
pip install pyinstaller pystray Pillow -q
if %errorlevel% neq 0 ( echo FAILED & pause & exit /b 1 )

echo Building exe...
python -m PyInstaller --clean meeting_reminder.spec
if %errorlevel% neq 0 ( echo FAILED & pause & exit /b 1 )

echo.
echo Done! exe is in dist folder.
echo Next: open installer.iss with Inno Setup and press Ctrl+F9
echo.
pause
