@echo off
echo === WorkBuddy EXE Builder ===
echo.

:: Collect static files (Django admin CSS/JS etc.)
echo [1/2] Collecting static files...
python manage.py collectstatic --noinput
if errorlevel 1 (
    echo ERROR: collectstatic failed.
    pause
    exit /b 1
)

:: Build the exe
echo.
echo [2/2] Building WorkBuddy.exe with PyInstaller...
pyinstaller workbuddy.spec --clean
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    pause
    exit /b 1
)

echo.
echo === Build complete! ===
echo Output: dist\WorkBuddy.exe
echo.
echo To distribute: copy dist\WorkBuddy.exe to any Windows machine with Outlook installed.
echo The db.sqlite3 will be created next to the .exe on first run.
pause
