#!/bin/bash

# Excel合并系统 - Mac双击启动文件
# .command文件在Mac上双击即可运行

cd "$(dirname "$0")"

echo "======================================"
echo "    Excel合并系统 - 正在启动..."  
echo "======================================"
echo ""

# 检查Python
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
else
    echo "❌ 错误: 未找到Python，请先安装Python 3.8+"
    echo "安装方法: brew install python3"
    read -p "按Enter键退出..."
    exit 1
fi

echo "✅ 发现Python版本: $($PYTHON_CMD --version)"

# 创建虚拟环境
if [ ! -d "venv" ]; then
    echo "📦 正在创建虚拟环境..."
    $PYTHON_CMD -m venv venv
    echo "✅ 虚拟环境创建完成"
else
    echo "✅ 虚拟环境已存在"
fi

# 激活虚拟环境并安装依赖
echo "📥 正在安装依赖..."
source venv/bin/activate

# 检查虚拟环境是否正确激活
if [ ! -f "venv/bin/pip" ]; then
    echo "❌ 虚拟环境pip未找到，重新创建虚拟环境..."
    rm -rf venv
    $PYTHON_CMD -m venv venv
    source venv/bin/activate
fi

# 使用绝对路径调用pip
venv/bin/pip install --upgrade pip
venv/bin/pip install -r requirements.txt
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
source venv/bin/activate
venv/bin/python $APP_FILE

echo ""
echo "服务器已停止"
read -p "按Enter键关闭..."
