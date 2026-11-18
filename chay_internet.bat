@echo off
chcp 65001 > nul
echo ========================================================
echo   KHOI DONG HE THONG DOC BAN DO RA INTERNET
echo ========================================================
echo 1. Dang khoi dong Web Server noi bo...
start "" python app.py
timeout /t 2 > nul
echo 2. Dang ket noi va tao duong link Internet (HTTPS)...
echo.
echo ========================================================
echo Hay chu y dong hien thi link dang: https://...app.lhr.life
echo Ban co the mo link do tren dien thoai hoac may tinh bat ky!
echo ========================================================
ssh -R 80:127.0.0.1:5000 -o StrictHostKeyChecking=no nokey@localhost.run
pause
