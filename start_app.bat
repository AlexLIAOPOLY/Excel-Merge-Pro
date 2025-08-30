@echo off
chcp 65001 >nul
title Excel合并系统启动器

echo ==========================================================
echo         Excel合并系统 - 自动化表格处理工具
echo ==========================================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo 正在启动Excel合并系统...
echo.

REM 运行Python启动脚本
python start_app.py

echo.
echo 按任意键退出...
pause >nul
