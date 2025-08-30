@echo off
chcp 65001 >nul
title Excel合并系统 - 简单启动版

echo ======================================
echo     Excel合并系统 - 简单启动版
echo ======================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误: 未找到Python，请先安装Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo ✅ 发现Python版本:
python --version

echo.
echo 📥 正在安装依赖到系统Python...
python -m pip install --upgrade pip --user >nul 2>&1
python -m pip install -r requirements.txt --user >nul 2>&1
echo ✅ 依赖安装完成

REM 检查应用文件
if exist "app_v2.py" (
    set APP_FILE=app_v2.py
) else if exist "app.py" (
    set APP_FILE=app.py
) else (
    echo ❌ 错误: 找不到应用文件
    echo.
    pause
    exit /b 1
)

echo.
echo ======================================
echo 🚀 启动Excel合并系统...
echo ======================================
echo 📍 本地访问地址: http://localhost:5002
echo ⚠️  按 Ctrl+C 停止服务器
echo.

set FLASK_ENV=development
python %APP_FILE%

echo.
echo 服务器已停止
echo.
pause
