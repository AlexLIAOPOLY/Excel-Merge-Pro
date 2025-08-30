#!/bin/bash

# Excel合并系统启动脚本 (Mac/Linux)
# 自动安装依赖并启动服务

set -e  # 遇到错误时退出

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}=========================================================="
echo -e "        Excel合并系统 - 自动化表格处理工具"
echo -e "==========================================================${NC}"
echo ""

# 检查Python版本
check_python() {
    if command -v python3 &> /dev/null; then
        PYTHON_CMD="python3"
    elif command -v python &> /dev/null; then
        PYTHON_CMD="python"
    else
        echo -e "${RED}错误: 未找到Python，请先安装Python 3.8+${NC}"
        echo "Mac安装方法: brew install python3"
        echo "或访问: https://www.python.org/downloads/"
        exit 1
    fi
    
    # 检查Python版本
    PYTHON_VERSION=$($PYTHON_CMD --version 2>&1 | awk '{print $2}')
    echo -e "${GREEN}✓ 发现Python版本: $PYTHON_VERSION${NC}"
}

# 检查并创建虚拟环境
setup_venv() {
    if [ ! -d "venv" ]; then
        echo -e "${YELLOW}正在创建虚拟环境...${NC}"
        $PYTHON_CMD -m venv venv
        echo -e "${GREEN}✓ 虚拟环境创建完成${NC}"
    else
        echo -e "${GREEN}✓ 虚拟环境已存在${NC}"
    fi
}

# 激活虚拟环境并安装依赖
install_deps() {
    echo -e "${YELLOW}正在安装依赖包...${NC}"
    
    # 激活虚拟环境
    source venv/bin/activate
    
    # 升级pip并安装依赖
    pip install --upgrade pip
    pip install -r requirements.txt
    
    echo -e "${GREEN}✓ 依赖包安装完成${NC}"
}

# 启动应用
start_app() {
    # 检查应用文件
    if [ -f "app_v2.py" ]; then
        APP_FILE="app_v2.py"
    elif [ -f "app.py" ]; then
        APP_FILE="app.py"
    else
        echo -e "${RED}错误: 找不到应用主文件${NC}"
        exit 1
    fi
    
    echo -e "${BLUE}=========================================================="
    echo -e "准备就绪! 即将启动Excel合并系统..."
    echo -e "==========================================================${NC}"
    echo ""
    echo -e "${GREEN}启动后请在浏览器中访问: http://localhost:5002${NC}"
    echo -e "${YELLOW}按 Ctrl+C 可以停止服务器${NC}"
    echo ""
    
    # 激活虚拟环境并启动应用
    source venv/bin/activate
    export FLASK_ENV=development
    python $APP_FILE
}

# 主执行流程
main() {
    check_python
    setup_venv
    install_deps
    start_app
}

# 捕获Ctrl+C
trap 'echo -e "\n${YELLOW}应用已停止${NC}"; exit 0' INT

# 运行主函数
main
