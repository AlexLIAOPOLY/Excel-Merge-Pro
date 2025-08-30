#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel合并系统 - 一键启动脚本
支持Windows、Mac和Linux系统
"""

import os
import sys
import subprocess
import platform
import time
from pathlib import Path

class ExcelMergeStarter:
    def __init__(self):
        self.system = platform.system()
        self.project_root = Path(__file__).parent
        self.venv_path = self.project_root / 'venv'
        
    def print_banner(self):
        """显示启动横幅"""
        banner = """
==========================================================
        Excel合并系统 - 自动化表格处理工具
==========================================================
系统: {system}
项目路径: {path}
Python版本: {python_version}
==========================================================
        """.format(
            system=self.system,
            path=self.project_root,
            python_version=sys.version.split()[0]
        )
        print(banner)
        
    def check_python_version(self):
        """检查Python版本"""
        version = sys.version_info
        if version.major < 3 or (version.major == 3 and version.minor < 8):
            print("错误: 需要Python 3.8或更高版本")
            print(f"当前版本: Python {version.major}.{version.minor}.{version.micro}")
            sys.exit(1)
        print(f"✓ Python版本检查通过: {version.major}.{version.minor}.{version.micro}")
        
    def get_python_command(self):
        """获取Python命令"""
        commands = ['python3', 'python']
        for cmd in commands:
            try:
                result = subprocess.run([cmd, '--version'], 
                                     capture_output=True, text=True)
                if result.returncode == 0:
                    return cmd
            except FileNotFoundError:
                continue
        return 'python'
        
    def get_pip_command(self):
        """获取pip命令"""
        python_cmd = self.get_python_command()
        return f"{python_cmd} -m pip"
        
    def create_virtual_environment(self):
        """创建虚拟环境"""
        if self.venv_path.exists():
            print("✓ 虚拟环境已存在")
            return
            
        print("正在创建虚拟环境...")
        python_cmd = self.get_python_command()
        
        try:
            subprocess.run([python_cmd, '-m', 'venv', 'venv'], 
                         cwd=self.project_root, check=True)
            print("✓ 虚拟环境创建成功")
        except subprocess.CalledProcessError as e:
            print(f"错误: 创建虚拟环境失败: {e}")
            sys.exit(1)
            
    def get_activation_script(self):
        """获取虚拟环境激活脚本路径"""
        if self.system == "Windows":
            return self.venv_path / 'Scripts' / 'activate.bat'
        else:
            return self.venv_path / 'bin' / 'activate'
            
    def get_venv_python(self):
        """获取虚拟环境中的Python路径"""
        if self.system == "Windows":
            return self.venv_path / 'Scripts' / 'python.exe'
        else:
            return self.venv_path / 'bin' / 'python'
            
    def install_dependencies(self):
        """安装依赖"""
        print("正在安装依赖包...")
        
        venv_python = self.get_venv_python()
        requirements_file = self.project_root / 'requirements.txt'
        
        if not requirements_file.exists():
            print("错误: requirements.txt文件不存在")
            sys.exit(1)
            
        try:
            # 升级pip
            subprocess.run([str(venv_python), '-m', 'pip', 'install', '--upgrade', 'pip'], 
                         cwd=self.project_root, check=True)
            
            # 安装依赖
            subprocess.run([str(venv_python), '-m', 'pip', 'install', '-r', 'requirements.txt'], 
                         cwd=self.project_root, check=True)
            print("✓ 依赖包安装完成")
        except subprocess.CalledProcessError as e:
            print(f"错误: 安装依赖失败: {e}")
            print("尝试手动运行:")
            print(f"  {venv_python} -m pip install -r requirements.txt")
            sys.exit(1)
            
    def check_main_app(self):
        """检查主应用文件"""
        app_files = ['app_v2.py', 'app.py']
        for app_file in app_files:
            app_path = self.project_root / app_file
            if app_path.exists():
                return app_file
        
        print("错误: 找不到应用主文件 (app_v2.py 或 app.py)")
        sys.exit(1)
        
    def start_application(self):
        """启动应用"""
        app_file = self.check_main_app()
        venv_python = self.get_venv_python()
        
        print(f"正在启动应用: {app_file}")
        print("=" * 60)
        
        # 设置环境变量
        env = os.environ.copy()
        env['FLASK_ENV'] = 'development'
        env['PYTHONPATH'] = str(self.project_root)
        
        try:
            # 启动Flask应用
            subprocess.run([str(venv_python), app_file], 
                         cwd=self.project_root, env=env)
        except KeyboardInterrupt:
            print("\n\n应用已停止")
        except subprocess.CalledProcessError as e:
            print(f"错误: 启动应用失败: {e}")
            sys.exit(1)
            
    def run(self):
        """主运行函数"""
        try:
            self.print_banner()
            self.check_python_version()
            self.create_virtual_environment()
            self.install_dependencies()
            
            print("\n" + "=" * 60)
            print("准备就绪! 即将启动Excel合并系统...")
            print("=" * 60)
            print("\n启动后请在浏览器中访问: http://localhost:5002")
            print("按 Ctrl+C 可以停止服务器")
            print("\n")
            
            time.sleep(2)  # 给用户时间阅读信息
            
            self.start_application()
            
        except KeyboardInterrupt:
            print("\n\n用户取消启动")
            sys.exit(0)
        except Exception as e:
            print(f"意外错误: {e}")
            sys.exit(1)

def main():
    """主入口函数"""
    starter = ExcelMergeStarter()
    starter.run()

if __name__ == "__main__":
    main()
