@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY where py >nul 2>&1 && set "PY=py"

if not defined PY (
  echo.
  echo 이 컴퓨터에는 파이썬이 없습니다.
  echo 설치 화면이 열리면 아래만 하시면 됩니다.
  echo.
  echo   1. Add python.exe to PATH 에 체크
  echo   2. Install Now
  echo   3. 설치가 끝나면 이 파일을 다시 더블클릭
  echo.
  start https://www.python.org/downloads/
  pause
  exit /b 1
)

echo 프로그램을 준비합니다. 처음 한 번만 조금 걸릴 수 있습니다.
"%PY%" -m pip install -r requirements.txt -q
if errorlevel 1 (
  echo.
  echo 준비가 실패했습니다. 인터넷이 되는 곳에서 다시 눌러 주세요.
  pause
  exit /b 1
)

echo 실행합니다.
"%PY%" main.py
if errorlevel 1 pause
