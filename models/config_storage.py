"""
API配置存储管理
提供简单的配置持久化功能
"""

import json
import os
from pathlib import Path
from datetime import datetime


class ConfigStorage:
    """配置存储管理器"""
    
    def __init__(self, config_file="api_config.json"):
        """
        初始化配置存储
        
        Args:
            config_file (str): 配置文件名称
        """
        self.config_dir = Path.cwd() / 'config'
        self.config_dir.mkdir(exist_ok=True)
        self.config_file = self.config_dir / config_file
        self._load_config()
    
    def _load_config(self):
        """加载配置文件"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            else:
                self.config = self._get_default_config()
                self._save_config()
        except Exception as e:
            print(f"[配置存储] 加载配置失败: {str(e)}")
            self.config = self._get_default_config()
    
    def _save_config(self):
        """保存配置到文件"""
        try:
            self.config['updated_at'] = datetime.now().isoformat()
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[配置存储] 保存配置失败: {str(e)}")
    
    def _get_default_config(self):
        """获取默认配置"""
        return {
            'api': {
                'provider': 'none',
                'url': '',
                'key': '',
                'model': 'none'
            },
            'created_at': datetime.now().isoformat(),
            'updated_at': datetime.now().isoformat()
        }
    
    def get_api_config(self):
        """获取API配置"""
        return self.config.get('api', {
            'provider': 'none',
            'url': '',
            'key': '',
            'model': 'none'
        })
    
    def set_api_config(self, provider, url='', key='', model=''):
        """
        设置API配置
        
        Args:
            provider (str): API提供商
            url (str): API地址
            key (str): API密钥
            model (str): 模型名称
        """
        self.config['api'] = {
            'provider': provider,
            'url': url,
            'key': key,
            'model': model
        }
        self._save_config()
        print(f"[配置存储] API配置已保存: {provider}")
    
    def update_api_config(self, **kwargs):
        """
        更新API配置（部分更新）
        
        Args:
            **kwargs: 要更新的配置项
        """
        if 'api' not in self.config:
            self.config['api'] = self._get_default_config()['api']
        
        self.config['api'].update(kwargs)
        self._save_config()
        print(f"[配置存储] API配置已更新: {kwargs}")
    
    def reset_config(self):
        """重置配置为默认值"""
        self.config = self._get_default_config()
        self._save_config()
        print("[配置存储] 配置已重置为默认值")


# 全局配置存储实例
_config_storage = None


def get_config_storage():
    """获取全局配置存储实例"""
    global _config_storage
    if _config_storage is None:
        _config_storage = ConfigStorage()
    return _config_storage


def get_api_config():
    """便捷函数：获取API配置"""
    storage = get_config_storage()
    return storage.get_api_config()


def set_api_config(provider, url='', key='', model=''):
    """便捷函数：设置API配置"""
    storage = get_config_storage()
    storage.set_api_config(provider, url, key, model)


if __name__ == "__main__":
    # 测试代码
    print("=== 配置存储测试 ===")
    
    storage = ConfigStorage('test_config.json')
    
    # 测试默认配置
    print("1. 默认配置:", storage.get_api_config())
    
    # 测试设置配置
    storage.set_api_config('ollama', 'http://localhost:11434', '', 'gemma2')
    print("2. 设置OLLAMA配置:", storage.get_api_config())
    
    # 测试部分更新
    storage.update_api_config(model='llama3.1')
    print("3. 更新模型后:", storage.get_api_config())
    
    # 测试非LLM配置
    storage.set_api_config('none')
    print("4. 非LLM配置:", storage.get_api_config())
    
    # 测试重置
    storage.reset_config()
    print("5. 重置后:", storage.get_api_config())