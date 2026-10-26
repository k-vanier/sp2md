@echo off
rem Construit dist\sp2md.exe : un seul fichier, aucune installation de Python requise sur le poste cible.
cd /d "%~dp0"
if not exist .venv python -m venv .venv
.venv\Scripts\python -m pip install -q -r requirements.txt || exit /b 1
.venv\Scripts\pyinstaller --noconfirm --clean --onefile --name sp2md --console ^
  --collect-all markitdown --collect-all magika --collect-all playwright ^
  sp2md.py || exit /b 1
echo.
echo Application prete : %~dp0dist\sp2md.exe
