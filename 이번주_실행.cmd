@echo off
chcp 65001 >nul
rem 이번 ISO 주차 폴더(국가별 weekly\<nn>주차\MX, CE)에서 법인별 원본을 모아(zip 풀기 · 시트 찾기 · 열 맞추기)
rem 1_Input 에 넣고 -> 점검 -> 매핑 자동 추론 -> 클렌징 -> 제안 -> 요약까지 한 번에.
rem 다른 주차는: 이번주_실행.cmd 를 쓰지 말고  python tools\run_all.py --week 43
cd /d "%~dp0"
set "PY=%USERPROFILE%\anaconda3\envs\workspace-py311\python.exe"
if not exist "%PY%" set "PY=py"
"%PY%" -X utf8 tools\run_all.py --week now --open %*
echo.
pause
