@echo off
rem Opera Analiz - elle calistirma. Excel dosyasi KAPALI olmali.
chcp 65001 >nul
set PYTHONUTF8=1
cd /d %~dp0
python opera_guncelle.py
echo.
pause
