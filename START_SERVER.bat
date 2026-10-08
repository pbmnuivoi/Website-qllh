@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul 2>&1
cd /d "%~dp0"

echo ==================================================
echo   TRUONG TIEU HOC - CSDL EXCEL TONG
echo ==================================================
echo.
echo [1/4] Dang kiem tra moi truong chay...

REM --------------------------------------------------
REM Uu tien Python neu co san: khong phu thuoc npm/node.
REM --------------------------------------------------
set "PYEXE="
where py >nul 2>&1
if not errorlevel 1 set "PYEXE=py -3"
if not defined PYEXE (
    where python >nul 2>&1
    if not errorlevel 1 set "PYEXE=python"
)

if defined PYEXE (
    echo [OK] Tim thay Python: %PYEXE%
    %PYEXE% -c "import openpyxl" >nul 2>&1
    if not errorlevel 1 goto START_PYTHON
    echo [!] Python co san nhung thieu openpyxl.
    echo     Dang thu cai openpyxl...
    %PYEXE% -m pip install openpyxl --user
    %PYEXE% -c "import openpyxl" >nul 2>&1
    if not errorlevel 1 goto START_PYTHON
    echo [!] Khong cai duoc openpyxl bang pip.
)

REM --------------------------------------------------
REM Neu khong co Python thi thu Node.js.
REM --------------------------------------------------
where node >nul 2>&1
if errorlevel 1 goto INSTALL_HELP
where npm >nul 2>&1
if errorlevel 1 goto INSTALL_HELP

echo [OK] Tim thay Node.js + npm.
if not exist "node_modules\express" goto NODE_INSTALL
if not exist "node_modules\xlsx" goto NODE_INSTALL
goto START_NODE

:NODE_INSTALL
echo [2/4] Dang cai cac thu vien Node.js...
npm install --no-audit --no-fund
if errorlevel 1 goto INSTALL_HELP
goto START_NODE

:START_PYTHON
echo [2/4] Khoi dong server Python + Excel...
start "TRUONG TIEU HOC - SERVER" /min cmd /k "cd /d ""%~dp0"" && %PYEXE% server.py"
goto WAIT_SERVER

:START_NODE
echo [2/4] Khoi dong server Node.js + Excel...
start "TRUONG TIEU HOC - SERVER" /min cmd /k "cd /d ""%~dp0"" && node server.js"
goto WAIT_SERVER

:WAIT_SERVER
echo [3/4] Dang cho server san sang...
set "READY=0"
for /L %%N in (1,1,20) do (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $r=Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:3000/api/health' -TimeoutSec 1; if($r.StatusCode -eq 200){exit 0}else{exit 1} } catch { exit 1 }" >nul 2>&1
    if not errorlevel 1 (
        set "READY=1"
        goto SERVER_READY
    )
    timeout /t 1 /nobreak >nul
)

if "%READY%"=="0" (
    echo.
    echo [LOI] Server khong phan hoi tai http://127.0.0.1:3000/
    echo Hay xem cua so "TRUONG TIEU HOC - SERVER" de xem loi chi tiet.
    pause
    exit /b 1
)

:SERVER_READY
echo [4/4] SERVER DA SAN SANG!
echo.
echo ==================================================
echo   Ung dung: http://localhost:3000/
echo   Kiem tra: http://localhost:3000/api/health
echo ==================================================
echo.
echo Dang mo ung dung bang HTTP (KHONG mo file index.html truc tiep)...
start "" "http://localhost:3000/"
echo.
echo Cua so SERVER da duoc thu nho xuong thanh taskbar va chay ngam.
echo Cua so khoi dong se tu dong dong.
timeout /t 2 /nobreak >nul
exit /b 0

:INSTALL_HELP
echo.
echo ==================================================
echo KHONG CO MOI TRUONG CHAY SERVER PHU HOP
 echo ==================================================
echo.
echo Ung dung can mot trong hai:
echo   1. Python 3 + openpyxl (khuyen dung cho ban nay)
echo   2. Node.js LTS + npm
 echo.
echo Thu mo trang cai Node.js LTS...
start "" "https://nodejs.org/en/download/"
echo.
echo Sau khi cai xong, hay DONG cua so nay va chay lai START_SERVER.bat.
pause
exit /b 1
