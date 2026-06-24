@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo   โปรแกรมนำเข้าข้อมูลศูนย์บริการ - Express
echo ============================================================
set PY=
where py >nul 2>nul && set PY=py
if not defined PY where python >nul 2>nul && set PY=python
if not defined PY goto NOPY
echo ตรวจสอบ/ติดตั้งไลบรารี xlrd ...
%PY% -c "import xlrd" 2>nul || %PY% -m pip install xlrd
echo เปิดโปรแกรม... (ปิดหน้าต่างนี้เพื่อหยุดโปรแกรม)
%PY% express_gui.py
goto END
:NOPY
echo [!] ไม่พบ Python บนเครื่องนี้
echo     ติดตั้งจาก https://www.python.org/downloads/
echo     ** ตอนติดตั้งให้ติ๊ก  Add Python to PATH **
:END
pause
