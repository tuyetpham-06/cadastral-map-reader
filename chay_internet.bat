@echo off
chcp 65001 > nul
echo ========================================================
echo   KHOI DONG HE THONG DOC BAN DO RA INTERNET
echo ========================================================
echo 1. Dang khoi dong Web Server...
start "" python app.py
timeout /t 2 > nul
echo 2. Dang tao duong link Internet (HTTPS) qua Cloudflare...
echo.
echo Hay chu y dong co link dang: https://...trycloudflare.com
echo Ban co the copy link do gui cho bat ky ai tren the gioi su dung!
echo ========================================================
.\cloudflared.exe tunnel --url http://127.0.0.1:5000
pause

