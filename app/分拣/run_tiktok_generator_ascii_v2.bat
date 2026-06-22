@echo off
cd /d "%~dp0"

python "%~dp0tiktok_shipping_list_generator.py"
if errorlevel 1 goto failed

echo DONE
echo Output folder: ..\..\data\分拣
pause
exit /b 0

:failed
echo FAILED
echo See log: ..\..\data\分拣\tiktok_shipping_list_generator.log
pause
exit /b 1
