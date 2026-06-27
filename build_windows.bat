@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo   สร้างไฟล์ ExpressImport.exe   (Win7: ต้องใช้ Python 3.8 แบบ 32-bit) (ไม่ต้องลง Python บนเครื่องปลายทาง)
echo ============================================================
set PY=
where py >nul 2>nul && set PY=py
if not defined PY where python >nul 2>nul && set PY=python
if not defined PY goto NOPY
echo ติดตั้ง pyinstaller + xlrd ...
%PY% -m pip install "pyinstaller==5.13.2" xlrd openpyxl
echo กำลัง build ...
%PY% -m PyInstaller --onefile --console --name ExpressImport --hidden-import xlrd --collect-all openpyxl express_gui.py
echo.
echo เสร็จแล้ว -^> dist\ExpressImport.exe
echo (นำ ExpressImport.exe ไปวางที่เครื่องไหนก็รันได้เลย ไม่ต้องลง Python)
goto END
:NOPY
echo [!] ไม่พบ Python  ติดตั้งจาก https://www.python.org/downloads/  (ติ๊ก Add to PATH)
:END
pause
