@echo off
chcp 65001 >nul
rem 다운로드 폴더 최근 7일 CSV 중 GMO 형식만 골라 1_Input 에 넣고 실행한다.
cd /d "%~dp0"
set "PY=%USERPROFILE%\anaconda3\envs\workspace-py311\python.exe"
if not exist "%PY%" set "PY=py"
"%PY%" -X utf8 tools\run_all.py --downloads --open %*
echo.
pause
