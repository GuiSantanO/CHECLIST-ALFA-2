@echo off
echo Installing dependencies...
pip install pyinstaller wmi psutil pillow pandas openpyxl customtkinter pywin32

echo Cleaning previous builds...
rmdir /s /q build dist
del /q *.spec

echo Building Executable...
python -m PyInstaller --noconsole --onefile --name="ChecklistApp" --collect-all customtkinter checklist_recondicionado.py

echo Build complete. Executable should be in the 'dist' folder.
pause
