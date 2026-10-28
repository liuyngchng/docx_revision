@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
title 党政机关公文格式审核与修订系统

cd /d "%~dp0"

rem 查找 Python（优先 python，其次 py 启动器）
set "PY=python"
where python >nul 2>nul
if errorlevel 1 (
    where py >nul 2>nul
    if errorlevel 1 (
        echo [错误] 未找到 Python，请先安装 Python 3.10 及以上版本。
        echo        下载地址: https://www.python.org/downloads/
        pause
        exit /b 1
    )
    set "PY=py"
)

rem 启动服务（支持传入参数，如 start.bat --port 9090 --no-browser）
%PY% run.py %*

echo.
echo 程序已退出。
pause
