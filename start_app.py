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
import socket
import webbrowser
import threading
from pathlib import Path

class ExcelMergeStarter:
    def __init__(self):
        self.system = platform.system()
        self.project_root = Path(__file__).parent
        self.venv_path = self.project_root / 'venv'
        self.default_port = 5002
        
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
        
    def is_port_available(self, port):
        """检查端口是否可用"""
        try:
            # 尝试连接到端口，如果连接成功说明端口被占用
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
                sock.settimeout(1)
                result = sock.connect_ex(('localhost', port))
                return result != 0  # 连接失败说明端口可用
        except Exception:
            # 如果出现异常，再尝试绑定测试
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
                    sock.bind(('localhost', port))
                    return True
            except OSError:
                return False
    
    def find_available_port(self, start_port=None):
        """查找可用端口"""
        if start_port is None:
            start_port = self.default_port
            
        port = start_port
        max_attempts = 100  # 最多尝试100个端口
        
        for attempt in range(max_attempts):
            if self.is_port_available(port):
                if attempt > 0:
                    print(f"  经过 {attempt + 1} 次尝试，找到可用端口: {port}")
                return port
            else:
                if attempt == 0:
                    print(f"  端口 {port} 被占用，正在寻找其他可用端口...")
                elif attempt < 5:  # 只显示前几次尝试
                    print(f"  端口 {port} 也被占用...")
            port += 1
            
        # 如果找不到可用端口，返回None
        return None
        
    def open_browser(self, url, delay=3):
        """延迟打开浏览器"""
        def delayed_open():
            time.sleep(delay)
            try:
                print(f"正在打开浏览器: {url}")
                
                # 尝试使用Chrome浏览器
                chrome_paths = {
                    'Darwin': [
                        '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
                        '/Applications/Chromium.app/Contents/MacOS/Chromium'
                    ],
                    'Windows': [
                        r'C:\Program Files\Google\Chrome\Application\chrome.exe',
                        r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
                    ],
                    'Linux': [
                        '/usr/bin/google-chrome',
                        '/usr/bin/google-chrome-stable',
                        '/usr/bin/chromium-browser',
                        '/usr/bin/chromium'
                    ]
                }
                
                chrome_found = False
                if self.system in chrome_paths:
                    for chrome_path in chrome_paths[self.system]:
                        if os.path.exists(chrome_path):
                            try:
                                subprocess.Popen([chrome_path, url])
                                print("✓ 已在Chrome浏览器中打开应用")
                                chrome_found = True
                                break
                            except:
                                continue
                
                # 如果找不到Chrome，使用系统默认浏览器
                if not chrome_found:
                    try:
                        webbrowser.open(url)
                        print("✓ 已在默认浏览器中打开应用")
                    except:
                        print("⚠ 无法自动打开浏览器，请手动访问: " + url)
                        
            except Exception as e:
                print(f"⚠ 打开浏览器时出错: {e}")
                print(f"请手动在浏览器中访问: {url}")
        
        # 在新线程中延迟打开浏览器
        browser_thread = threading.Thread(target=delayed_open, daemon=True)
        browser_thread.start()
        
    def start_application(self, port):
        """启动应用"""
        app_file = self.check_main_app()
        venv_python = self.get_venv_python()
        
        print(f"正在启动应用: {app_file}")
        print("=" * 60)
        
        # 设置环境变量
        env = os.environ.copy()
        env['FLASK_ENV'] = 'development'
        env['PYTHONPATH'] = str(self.project_root)
        env['PORT'] = str(port)  # 设置端口环境变量
        
        # 准备在应用启动后自动打开浏览器
        app_url = f"http://localhost:{port}"
        self.open_browser(app_url, delay=4)  # 延迟4秒打开浏览器，给Flask更多时间启动
        
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
            
            # 提前检查可用端口
            print("正在检查端口可用性...")
            available_port = self.find_available_port()
            
            if available_port is None:
                print("错误: 无法找到可用端口 (已尝试100个连续端口)")
                sys.exit(1)
                
            if available_port != self.default_port:
                print(f"✓ 默认端口 {self.default_port} 被占用，自动切换到端口 {available_port}")
            else:
                print(f"✓ 使用默认端口 {available_port}")
            
            print("\n" + "=" * 60)
            print("准备就绪! 即将启动Excel合并系统...")
            print("=" * 60)
            print(f"\n启动后将自动在Google Chrome浏览器中打开: http://localhost:{available_port}")
            print("如果浏览器没有自动打开，请手动访问上述地址")
            print("按 Ctrl+C 可以停止服务器")
            print("\n")
            
            time.sleep(2)  # 给用户时间阅读信息
            
            self.start_application(available_port)
            
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
