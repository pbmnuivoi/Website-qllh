@echo off
setlocal
chcp 65001 >nul

echo ==================================================
echo   CAI NODE.JS LTS CHO SCHOOL APP
 echo ==================================================
echo.
where winget >nul 2>&1
if errorlevel 1 (
  echo Khong tim thay winget tren may nay.
  echo Hay tai Node.js LTS tai trang chinh thuc:
  start "" "https://nodejs.org/en/download/"
  pause
  exit /b 1
)

echo Dang cai Node.js LTS...
winget install --id OpenJS.NodeJS.LTS -e --accept-source-agreements --accept-package-agreements
if errorlevel 1 (
  echo.
  echo Cai dat khong thanh cong. Hay thu cai thu cong.
  start "" "https://nodejs.org/en/download/"
  pause
  exit /b 1
)

echo.
echo Da cai Node.js LTS. Hay dong cua so nay va chay START_SERVER.bat.
pause
