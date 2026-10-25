@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
python --version >nul 2>&1 || (echo يجب تثبيت Python 3.11 أو أحدث من https://www.python.org/downloads/windows/ وتفعيل Add Python to PATH. & pause & exit /b 1)
python -m pip install --upgrade pip
if errorlevel 1 goto error
python -m pip install -r requirements.txt
if errorlevel 1 goto error
python -m PyInstaller --noconfirm --clean --windowed --name FactureDesktop --onefile app.py
if exist dist\FactureDesktop.exe (echo. & echo تم إنشاء: %~dp0dist\FactureDesktop.exe & goto end)
:error
echo فشل البناء. راجع رسائل الخطأ أعلاه.
:end
pause
