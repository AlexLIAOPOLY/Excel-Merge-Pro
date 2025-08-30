#!/bin/bash

# Excel合并系统 - 简单启动版（不使用虚拟环境）
cd "$(dirname "$0")"

echo "======================================"
echo "    Excel合并系统 - 简单启动版"  
echo "======================================"
echo ""

# 检查Python
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
else
    echo "❌ 错误: 未找到Python"
    read -p "按Enter键退出..."
    exit 1
fi

echo "✅ 使用Python: $($PYTHON_CMD --version)"

# 直接安装依赖到系统Python
echo "📥 正在安装依赖到系统Python..."
$PYTHON_CMD -m pip install --upgrade pip --user
$PYTHON_CMD -m pip install -r requirements.txt --user
echo "✅ 依赖安装完成"

# 启动应用
if [ -f "app_v2.py" ]; then
    APP_FILE="app_v2.py"
elif [ -f "app.py" ]; then
    APP_FILE="app.py"
else
    echo "❌ 错误: 找不到应用文件"
    read -p "按Enter键退出..."
    exit 1
fi

echo ""
echo "======================================"
echo "🚀 启动Excel合并系统..."
echo "======================================"
echo "📍 本地访问地址: http://localhost:5002"
echo "⚠️  按 Ctrl+C 停止服务器"
echo ""

export FLASK_ENV=development
$PYTHON_CMD $APP_FILE

echo ""
echo "服务器已停止"
read -p "按Enter键关闭..."
