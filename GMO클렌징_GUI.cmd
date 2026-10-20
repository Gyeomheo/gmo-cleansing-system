@echo off
chcp 65001 >nul
rem GMO 클렌징 GUI: 주차 폴더 고르기 -> 수집 · 클렌징 -> 취합 검증 -> 취합본 만들기 ([전체 실행] 한 번)
cd /d "%~dp0"
set "PY=%USERPROFILE%\anaconda3\envs\workspace-py311\pythonw.exe"
if not exist "%PY%" set "PY=pyw"
start "" "%PY%" tools\gui.py
