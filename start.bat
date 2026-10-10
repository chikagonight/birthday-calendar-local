@echo off
cd /d "%~dp0"
echo.
where node >nul 2>&1
if errorlevel 1 (
  echo Node.js не найден. Установите Node.js 20 или новее с https://nodejs.org/
  pause
  exit /b 1
)
start "Birthday Calendar Server" cmd /k "node server.js"
timeout /t 2 /nobreak >nul
start "" http://localhost:8080
echo Приложение запущено: http://localhost:8080
echo Для остановки закройте окно "Birthday Calendar Server".
timeout /t 4 /nobreak >nul
