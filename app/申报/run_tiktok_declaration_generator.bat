@echo off
setlocal
cd /d "%~dp0"
set "PYFILE=tiktok_declaration_generator.py"

echo Running TikTok declaration generator...

where python.exe >nul 2>nul
if not errorlevel 1 (
    python.exe "%PYFILE%"
) else (
    py -3 "%PYFILE%"
)

if errorlevel 1 (
    echo.
    echo Failed. Please open ..\..\data\申报\tiktok_declaration_generator.log for details.
    pause
    exit /b 1
)

echo.
echo Done. Check the output Excel file in ..\..\data\申报.
pause
endlocal
