@echo off
cd /d "%~dp0"
echo ============================================
echo    TRENDYOL SIPARIS HAZIRLAMA
echo ============================================
echo.
echo Program baslatiliyor...
echo Chrome'da acilacak: http://localhost:5000
echo.
echo Bu pencereyi KAPATMAYIN!
echo.

start http://localhost:5000
python app.py

pause
