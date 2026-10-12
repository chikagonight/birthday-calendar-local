@echo off
cd /d "%~dp0"
echo.
where node >nul 2>&1
if errorlevel 1 (
  where winget >nul 2>&1
  if not errorlevel 1 (
    echo Node.js не найден. Попытка установить Node.js LTS через winget...
    winget install -e --id OpenJS.NodeJS.LTS --accept-source-agreements --accept-package-agreements --silent --disable-interactivity
    if not errorlevel 1 (
      set "PATH=%PATH%;%ProgramFiles%\nodejs;%LOCALAPPDATA%\Programs\nodejs"
    )
  )
  where node >nul 2>&1
  if errorlevel 1 (
    echo Node.js не найден. Установите Node.js 20 или новее с https://nodejs.org/
    echo Если Node.js уже установлен, перезапустите start.bat.
    pause
    exit /b 1
  )
)
set PORT=8090
start "Birthday Calendar Server" cmd /k "set PORT=%PORT% && node server.js"
timeout /t 2 /nobreak >nul
start "" http://localhost:%PORT%
echo Приложение запущено: http://localhost:%PORT%
echo Для остановки закройте окно "Birthday Calendar Server".
timeout /t 4 /nobreak >nul
