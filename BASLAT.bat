@echo off
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
python "%~dp0app.py"

pause
