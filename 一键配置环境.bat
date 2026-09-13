@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

rem 依次实测候选命令，能真正运行才采用（仅用 where 找到文件并不代表可用）
set "PYCMD="
py -3 -c "import sys" >nul 2>nul && set "PYCMD=py -3"
if not defined PYCMD (python -c "import sys" >nul 2>nul && set "PYCMD=python")
if not defined PYCMD (python3 -c "import sys" >nul 2>nul && set "PYCMD=python3")

if not defined PYCMD (
    echo.
    echo 未找到可用的 Python。
    echo 请安装 Python 3.10 或更高版本:
    echo   https://www.python.org/downloads/
    echo 安装时请勾选 Add python.exe to PATH
    echo.
    pause
    exit /b 1
)

rem 具体逻辑与中文提示都在 Python 脚本内，避免批处理的编码问题
%PYCMD% "scripts\setup_env.py"
if errorlevel 1 pause
