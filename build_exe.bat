@echo off
REM Builds the Windows app "GtR-Search.exe" in the dist folder from gtr_gui.py.
REM Needs Python 3.8+ on this computer. Colleagues who only use the app do not need this.
cd /d "%~dp0"
python -m pip install --upgrade "pyinstaller==6.10.0"
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "GtR-Search" gtr_gui.py
echo.
echo Done: %~dp0dist\GtR-Search.exe
pause
