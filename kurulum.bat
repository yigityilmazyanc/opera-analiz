@echo off
rem ================================================================
rem  Opera Analiz - TEK SEFERLIK KURULUM (Windows)
rem  Cift tikla: paketleri kurar, ayar dosyasini olusturup acar.
rem  On kosul: Python 3.11+ kurulu olmali (python.org - kurarken
rem  "Add python.exe to PATH" isaretli). Gerisini bu dosya yapar.
rem ================================================================
chcp 65001 >nul
cd /d %~dp0

set PY=python
where py >nul 2>nul
if not errorlevel 1 set PY=py

%PY% --version >nul 2>nul
if errorlevel 1 goto python_yok

echo === [1/2] Gerekli paketler kuruluyor (eptr2, pandas, openpyxl, pywin32) ===
%PY% -m pip install --quiet eptr2 pandas openpyxl pywin32
if errorlevel 1 goto pip_hata

echo.
echo === [2/2] Ayar dosyasi ===
if exist ayarlar.json goto ayar_var
copy ayarlar.ornek.json ayarlar.json >nul
echo ayarlar.json olusturuldu - simdi Not Defteri'nde acilacak.
echo Doldurun ve KAYDEDIN:
echo   api_kullanici / api_sifre : EPIAS Seffaflik hesabi (kayit.epias.com.tr)
echo   mail_gonderen             : gonderen Gmail adresi
echo   mail_sifre                : Gmail UYGULAMA sifresi (myaccount.google.com/apppasswords)
echo   mail_kime                 : rapor gidecek adres(ler)
start /wait notepad ayarlar.json
goto son

:ayar_var
echo ayarlar.json zaten var - dokunulmadi.
goto son

:python_yok
echo HATA: Python bulunamadi.
echo https://www.python.org/downloads/ adresinden kurun;
echo kurulumda "Add python.exe to PATH" kutusunu isaretleyin, sonra bu dosyayi yeniden calistirin.
pause
exit /b 1

:pip_hata
echo HATA: paket kurulumu basarisiz - internet baglantisini kontrol edin.
pause
exit /b 1

:son
echo.
echo KURULUM TAMAM.
echo   Veri guncelle           : calistir.bat
echo   Guncelle + PDF + e-posta: gonder.bat
echo   Excel dosyasi Downloads klasorunde "Opera Analiz AY YIL.xlsx" adiyla durmali.
pause
