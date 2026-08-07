@echo off
cd /d "%~dp0"
start "" http://localhost:8080/escuelitas.html
python -m http.server 8080
