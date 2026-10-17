@echo off
REM Build script para Windows (PyInstaller)
pyinstaller --onefile --windowed main.py --name "EditorPlanilha" --distpath dist
echo Build finalizado. O executável estará em dist\EditorPlanilha.exe
pause
