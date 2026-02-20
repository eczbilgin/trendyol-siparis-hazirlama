@echo off
cd /d "%~dp0"
echo ============================================
echo    TRENDYOL SIPARIS HAZIRLAMA
echo ============================================
echo.

:: Python kurulu mu kontrol et
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [HATA] Python bulunamadi!
    echo Python'u https://www.python.org/downloads/ adresinden indirip kurun.
    echo Kurulum sirasinda "Add Python to PATH" secenegini isaretleyin!
    echo.
    pause
    exit /b
)

:: Gerekli paketleri kur
echo Gerekli paketler kontrol ediliyor...
pip install -r requirements.txt -q
echo.

echo Program baslatiliyor...
echo Chrome'da acilacak: http://localhost:5000
echo.
echo Bu pencereyi KAPATMAYIN!
echo.

start http://localhost:5000
python app.py

pause
