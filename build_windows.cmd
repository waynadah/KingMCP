@echo off
setlocal
cd /d "%~dp0"
if not exist main.py exit /b 1
if not exist desktop_server.py exit /b 1
if not exist voice_tools.py exit /b 1
python -m pip install -r requirements.txt || exit /b 1
python -m pip install "pyinstaller>=6,<7" || exit /b 1
python -m PyInstaller --noconfirm --clean --onedir --console --name DesktopAI-MCP --icon desktopai.ico --collect-all mcp --collect-all pywinauto --collect-all docx --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywintypes desktop_server.py || exit /b 1
python -m PyInstaller --noconfirm --clean --onedir --windowed --name DesktopAI --icon desktopai.ico --collect-all customtkinter --collect-all mcp --collect-all speech_recognition --collect-all pynput --hidden-import win32com.client --hidden-import win32crypt --hidden-import pythoncom --hidden-import pywintypes main.py || exit /b 1
robocopy "dist\DesktopAI-MCP" "dist\DesktopAI\server" /E /NFL /NDL /NJH /NJS
if errorlevel 8 exit /b 1
echo Build complete: dist\DesktopAI\DesktopAI.exe
echo Keep the entire DesktopAI folder together.
