@echo off
chcp 65001 >nul
title PQR 文档批量转化工具
cd /d "%~dp0"

echo ========================================
echo   PQR 文档批量转化工具 启动中...
echo ========================================
echo.

python gui.py
if errorlevel 1 (
    echo.
    echo 启动失败，请检查：
    echo   1. Python 已安装并加入 PATH
    echo   2. 已安装 pywin32 （命令: pip install pywin32）
    echo.
    pause
)
