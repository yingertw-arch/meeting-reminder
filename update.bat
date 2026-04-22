@echo off
cd /d "%~dp0"

echo Enter version (e.g. v1.1.0):
set /p VERSION=

echo Enter update description (e.g. fix bug, add feature):
set /p MSG=

git add .
git commit -m "%VERSION%: %MSG%"
git push origin main

echo.
echo Done! Pushed to GitHub.
echo Remember to create a new Release on GitHub if needed.
pause
