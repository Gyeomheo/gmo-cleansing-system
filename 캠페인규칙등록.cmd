@echo off
chcp 65001 >nul
rem 2_Output\Reports\캠페인규칙제안.xlsx 에서 '사용' 칸에 O 한 행을 3_Config\campaign_rules.xlsx 에 추가하고 다시 실행한다.
rem --remap: 숫자가 같아도 입력 기간 전체를 마스터에 다시 써서, 바뀐 분류가 과거 행에도 반영되게 한다.
cd /d "%~dp0"
set "PY=%USERPROFILE%\anaconda3\envs\workspace-py311\python.exe"
if not exist "%PY%" set "PY=py"
"%PY%" -X utf8 tools\campaign_rules.py apply && "%PY%" -X utf8 tools\run_all.py --remap
echo.
pause
