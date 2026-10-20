@echo off
chcp 65001 >nul
rem GMO 클렌징 한 번에: 넣기 -> 점검 -> 클렌징 -> 매핑 제안 -> 실행 요약
rem 받은 CSV 를 이 파일 위에 끌어다 놓으면 판별해서 1_Input 에 넣고 실행한다.
cd /d "%~dp0"
set "PY=%USERPROFILE%\anaconda3\envs\workspace-py311\python.exe"
if not exist "%PY%" set "PY=py"
"%PY%" -X utf8 tools\run_all.py --open %*
echo.
pause
