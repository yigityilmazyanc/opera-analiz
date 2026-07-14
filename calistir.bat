@echo off
rem Opera Analiz - elle calistirma. Excel dosyasi KAPALI olmali.
chcp 65001 >nul
set PYTHONUTF8=1
cd /d %~dp0
where py >nul 2>nul && (set PY=py) || (set PY=python)
%PY% opera_guncelle.py
echo.
pause
