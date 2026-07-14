@echo off
rem Opera + Piyasa verilerini gunceller, raporu PDF yapip e-posta ile gonderir.
rem Sadece guncelleme icin calistir.bat kullanin.
chcp 65001 >nul
set PYTHONUTF8=1
cd /d %~dp0
set PY=python
where py >nul 2>nul
if not errorlevel 1 set PY=py

echo === [1/3] Opera Analiz verileri guncelleniyor ===
%PY% opera_guncelle.py
if errorlevel 1 echo UYARI: Opera guncellenemedi - Excel acik olabilir, MEVCUT haliyle gonderilecek.

echo.
echo === [2/3] Piyasa Analiz verileri guncelleniyor ===
if not exist "%~dp0..\piyasa-analiz\piyasa_guncelle.py" goto piyasa_yok
pushd "%~dp0..\piyasa-analiz"
%PY% piyasa_guncelle.py
if errorlevel 1 echo UYARI: Piyasa guncellenemedi - Excel acik olabilir, MEVCUT haliyle gonderilecek.
popd
goto mail
:piyasa_yok
echo piyasa-analiz klasoru bulunamadi - Piyasa guncellemesi atlaniyor.

:mail
echo.
echo === [3/3] PDF donusumu + e-posta ===
%PY% mail_gonder.py
if errorlevel 1 goto hata
echo.
echo TAMAM - gonderildi.
pause
exit /b 0
:hata
echo.
echo HATA: gonderilemedi - yukaridaki mesaja bakin.
pause
exit /b 1
