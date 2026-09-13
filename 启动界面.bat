@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

rem 项目只使用一个环境：.venv（由「一键配置环境.bat」创建）
set "VENV_PY=.venv\Scripts\python.exe"

if not exist "%VENV_PY%" (
    echo 未找到项目环境 .venv，请先双击「一键配置环境.bat」完成配置。
    pause
    exit /b 1
)

"%VENV_PY%" "scripts\gui_app.py" %*
if errorlevel 1 pause
