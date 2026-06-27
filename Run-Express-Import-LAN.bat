@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo   โปรแกรมนำเข้าข้อมูล - โหมดแชร์ในวง LAN
echo   (เครื่องอื่นเปิดเบราว์เซอร์ไปที่ IP ที่โปรแกรมแสดง)
echo ============================================================
set PY=
where py >nul 2>nul && set PY=py
if not defined PY where python >nul 2>nul && set PY=python
if not defined PY goto NOPY
%PY% -c "import xlrd, openpyxl" 2>nul || %PY% -m pip install xlrd openpyxl
%PY% express_gui.py --lan
goto END
:NOPY
echo [!] ไม่พบ Python  ติดตั้งจาก https://www.python.org/downloads/  (ติ๊ก Add to PATH)
:END
pause
