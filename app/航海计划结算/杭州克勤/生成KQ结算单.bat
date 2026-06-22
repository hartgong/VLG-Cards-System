@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "BUNDLED_PY=C:\Users\feiyu\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if exist "%BUNDLED_PY%" (
  "%BUNDLED_PY%" "%~dp0kq_settlement_generator.py"
  pause
  exit /b %ERRORLEVEL%
)

python "%~dp0kq_settlement_generator.py"
if %ERRORLEVEL% EQU 0 (
  pause
  exit /b 0
)

py -3 "%~dp0kq_settlement_generator.py"
pause
