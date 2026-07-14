@echo off
rem Opera + Piyasa verilerini gunceller, raporu PDF yapip e-posta ile gonderir.
rem Sadece guncelleme icin calistir.bat kullanin.
chcp 65001 >nul
set PYTHONUTF8=1
cd /d %~dp0
where py >nul 2>nul && (set PY=py) || (set PY=python)

echo === [1/3] Opera Analiz verileri guncelleniyor ===
%PY% opera_guncelle.py
if errorlevel 1 echo UYARI: Opera guncellenemedi (Excel acik olabilir) - MEVCUT haliyle gonderilecek.

echo.
echo === [2/3] Piyasa Analiz verileri guncelleniyor ===
if exist "%~dp0..\piyasa-analiz\piyasa_guncelle.py" (
  pushd "%~dp0..\piyasa-analiz"
  %PY% piyasa_guncelle.py
  if errorlevel 1 echo UYARI: Piyasa guncellenemedi (Excel acik olabilir) - MEVCUT haliyle gonderilecek.
  popd
) else (
  echo piyasa-analiz klasoru bulunamadi - Piyasa guncellemesi atlaniyor.
)

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
