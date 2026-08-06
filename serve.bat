@echo off
REM Suivi Usine (statique) — serveur local pour tester avant de pousser sur GitHub.
REM Necessaire car les navigateurs bloquent fetch() sur un fichier ouvert en double-clic (file://).

cd /d "%~dp0"
echo Ouvre http://localhost:8000/ dans ton navigateur (Ctrl+C ici pour arreter)
python -m http.server 8000
