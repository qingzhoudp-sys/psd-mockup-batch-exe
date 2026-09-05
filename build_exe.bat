@echo off
setlocal
cd /d "%~dp0"
py -m pip install --upgrade pyinstaller pywin32 tkinterdnd2
py -m PyInstaller --noconfirm --clean --onefile --windowed --name "PSD样机批量替换" --collect-submodules win32com --collect-all tkinterdnd2 --hidden-import pythoncom --hidden-import pywintypes app.py
echo.
echo EXE 已生成：dist\PSD样机批量替换.exe
pause
