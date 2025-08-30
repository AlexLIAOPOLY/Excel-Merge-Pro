#!/bin/bash

echo "正在启动自动化表格处理系统..."

# 激活虚拟环境
source venv/bin/activate

# 检查依赖是否安装
if ! python3 -c "import flask" 2>/dev/null; then
    echo "正在安装依赖包..."
    pip install Flask==2.3.3 Flask-CORS==4.0.0 pandas==2.0.3 openpyxl==3.1.2 SQLAlchemy==2.0.21 Flask-SQLAlchemy==3.0.5 Werkzeug==2.3.7 xlrd==2.0.1
fi

# 启动应用
echo "启动Flask应用..."
python3 app.py