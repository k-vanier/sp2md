@echo off
rem Construit dist\sp2md.exe : un seul fichier, sans Python a installer sur le poste cible.
cd /d "%~dp0"
if not exist .venv python -m venv .venv
.venv\Scripts\python -m pip install -q -r requirements.txt || exit /b 1
rem MkDocs trouve theme, extensions et greffons par les metadonnees des paquets : on les embarque.
.venv\Scripts\pyinstaller --noconfirm --clean --onefile --name sp2md --console ^
  --collect-all markitdown --collect-all magika --collect-all playwright ^
  --collect-all mkdocs --collect-all material --collect-all pymdownx ^
  --copy-metadata mkdocs --copy-metadata mkdocs-material --copy-metadata pymdown-extensions ^
  --copy-metadata markdown --copy-metadata mkdocs-get-deps ^
  sp2md.py || exit /b 1
echo.
echo Application prete : %~dp0dist\sp2md.exe
