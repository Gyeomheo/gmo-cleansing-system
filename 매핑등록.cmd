@echo off
chcp 65001 >nul
rem 2_Output\Reports\매핑등록.xlsx 에서 '등록' 칸에 O 한 행을 매핑 엑셀에 추가하고 다시 실행한다.
rem --remap: 숫자가 같아도 입력 기간 전체를 마스터에 다시 써서, 바뀐 매핑이 과거 행에도 반영되게 한다.
cd /d "%~dp0"
set "PY=%USERPROFILE%\anaconda3\envs\workspace-py311\python.exe"
if not exist "%PY%" set "PY=py"
"%PY%" -X utf8 tools\mapping.py apply && "%PY%" -X utf8 tools\run_all.py --remap
echo.
pause
