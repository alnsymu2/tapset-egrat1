@echo off
chcp 65001 >nul
cd /d "%~dp0dist"
echo مُتابع المعاملات يعمل على http://localhost:8000
python -m http.server 8000
pause
