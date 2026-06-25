@echo off
cd /d "%~dp0"
echo 자동 번역기를 시작합니다...
echo 브라우저가 열리면 문서 경로를 입력하세요.
echo 종료하려면 이 창에서 Ctrl + C 를 누르세요.
echo.
python "auto_translator.py"
if errorlevel 1 (
  echo.
  echo [오류] 실행에 실패했습니다. Python 또는 lxml 설치를 확인하세요.
  echo   설치 명령:  pip install lxml
  pause
)
