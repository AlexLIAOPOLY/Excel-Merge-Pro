"""
统一API管理器
支持多种AI提供商和非LLM模式的表格命名
"""

import json
from datetime import datetime
from .deepseek_api import DeepSeekAPIClient
from .ollama_api import OllamaAPIClient


class APIManager:
    """统一的API管理器，支持多种提供商"""
    
    def __init__(self, config=None):
        """
        初始化API管理器
        
        Args:
            config (dict): API配置信息，包含provider, url, key, model等
        """
        self.config = config or {}
        self.provider = self.config.get('provider', 'none')
        self.counter = 1  # 用于非LLM模式的计数器
    
    def generate_table_name(self, column_names, sample_data=None, filename=None):
        """
        生成表格名称（根据配置的提供商）
        
        Args:
            column_names (list): 表格列名列表
            sample_data (list): 样本数据
            filename (str): 文件名
            
        Returns:
            tuple: (成功标志, 表格名称, 消息)
        """
        print(f"[API管理器] 使用提供商: {self.provider}")
        
        try:
            if self.provider == 'none':
                return self._generate_no_llm_name()
            elif self.provider == 'ollama':
                return self._generate_ollama_name(column_names, sample_data)
            elif self.provider == 'deepseek':
                return self._generate_deepseek_name(column_names, sample_data)
            elif self.provider in ['openai', 'anthropic', 'gemini', 'zhipu', 'hunyuan', 'qwen', 'doubao', 'moonshot']:
                return self._generate_other_provider_name(column_names, sample_data)
            else:
                # 未知提供商，使用默认命名
                return self._generate_fallback_name(column_names, filename)
                
        except Exception as e:
            print(f"[API管理器] 生成表格名称时出错: {str(e)}")
            return self._generate_fallback_name(column_names, filename)
    
    def _generate_no_llm_name(self):
        """生成非LLM模式的名称：合并表1, 合并表2 ...

        规则：读取数据库中现有的“合并表N”序列，返回下一个未使用的序号，
        避免出现“合并表1_2”这类名称。
        """
        try:
            # 延迟导入以避免循环依赖
            from .database import TableGroup
            import re

            # 查询所有以“合并表”开头的分组名称
            existing_names = [g.group_name for g in TableGroup.query.all() if g.group_name]

            # 提取严格形如“合并表<number>”的编号
            used_numbers = set()
            pattern = re.compile(r'^合并表(\d+)$')
            for name in existing_names:
                m = pattern.match(name.strip())
                if m:
                    try:
                        used_numbers.add(int(m.group(1)))
                    except ValueError:
                        continue

            # 选择最小未使用的正整数（从1开始）
            next_num = 1
            while next_num in used_numbers:
                next_num += 1

            table_name = f"合并表{next_num}"
        except Exception:
            # 退化处理：使用进程内计数器，确保不中断
            table_name = f"合并表{self.counter}"
            self.counter += 1

        print(f"[API管理器] 非LLM模式生成名称: {table_name}")
        return True, table_name, "使用非LLM默认命名"
    
    def _generate_ollama_name(self, column_names, sample_data):
        """使用OLLAMA生成表格名称"""
        try:
            base_url = self.config.get('url', 'http://localhost:11434')
            model = self.config.get('model', 'gemma2')
            
            client = OllamaAPIClient(base_url=base_url, model=model)
            return client.generate_table_name(column_names, sample_data)
            
        except Exception as e:
            print(f"[API管理器] OLLAMA调用失败: {str(e)}")
            return self._generate_fallback_name(column_names)
    
    def _generate_deepseek_name(self, column_names, sample_data):
        """使用DeepSeek生成表格名称"""
        try:
            api_key = self.config.get('key')
            client = DeepSeekAPIClient(api_key=api_key)
            return client.generate_table_name(column_names, sample_data)
            
        except Exception as e:
            print(f"[API管理器] DeepSeek调用失败: {str(e)}")
            return self._generate_fallback_name(column_names)
    
    def _generate_other_provider_name(self, column_names, sample_data):
        """其他提供商的命名逻辑（待实现）"""
        print(f"[API管理器] {self.provider} 提供商暂未实现，使用默认命名")
        return self._generate_fallback_name(column_names)
    
    def _generate_fallback_name(self, column_names, filename=None):
        """生成默认表格名称"""
        from datetime import datetime
        
        # 根据列名推断表格类型
        name_keywords = {
            '员工信息': ['员工', '姓名', '工号', '部门', '职位'],
            '客户资料': ['客户', '公司', '联系人', '电话', '地址'],
            '订单记录': ['订单', '商品', '数量', '金额', '价格'],
            '库存管理': ['库存', '商品', '数量', '仓库', '入库'],
            '财务报表': ['金额', '收入', '支出', '预算', '成本'],
            '项目管理': ['项目', '进度', '负责人', '开始', '完成'],
            '销售数据': ['销售', '业绩', '客户', '提成', '目标'],
            '产品目录': ['产品', '型号', '价格', '规格', '品牌'],
            '合同信息': ['合同', '签订', '甲方', '乙方', '金额'],
            '学生成绩': ['学生', '成绩', '科目', '班级', '学号'],
            '设备清单': ['设备', '型号', '序列号', '状态', '位置']
        }
        
        # 检查列名中是否包含关键词
        column_text = ''.join(column_names).lower()
        
        for table_type, keywords in name_keywords.items():
            if sum(1 for keyword in keywords if keyword in column_text) >= 2:
                print(f"[API管理器] 根据列名推断为: {table_type}")
                return True, table_type, "根据列名智能推断"
        
        # 如果从文件名能推断出类型
        if filename:
            filename_lower = filename.lower()
            for table_type, keywords in name_keywords.items():
                if any(keyword in filename_lower for keyword in keywords):
                    print(f"[API管理器] 根据文件名推断为: {table_type}")
                    return True, table_type, "根据文件名智能推断"
        
        # 默认使用时间戳命名
        current_time = datetime.now()
        default_name = f"合并表_{current_time.strftime('%m%d_%H%M')}"
        print(f"[API管理器] 使用默认命名: {default_name}")
        return True, default_name, "使用默认命名规则"
    
    def test_connection(self):
        """测试API连接"""
        try:
            if self.provider == 'none':
                return True, "非LLM模式，无需连接测试"
            elif self.provider == 'ollama':
                base_url = self.config.get('url', 'http://localhost:11434')
                model = self.config.get('model', 'gemma2')
                client = OllamaAPIClient(base_url=base_url, model=model)
                return client.test_connection()
            elif self.provider == 'deepseek':
                api_key = self.config.get('key')
                client = DeepSeekAPIClient(api_key=api_key)
                return client.test_connection()
            else:
                return False, f"暂不支持 {self.provider} 提供商的连接测试"
                
        except Exception as e:
            return False, f"连接测试失败: {str(e)}"
    
    def get_available_models(self):
        """获取可用模型列表"""
        try:
            if self.provider == 'none':
                return True, [{'value': 'none', 'label': '不使用AI模型'}], "非LLM模式"
            elif self.provider == 'ollama':
                base_url = self.config.get('url', 'http://localhost:11434')
                model = self.config.get('model', 'gemma2')
                client = OllamaAPIClient(base_url=base_url, model=model)
                return client.get_available_models()
            elif self.provider == 'deepseek':
                api_key = self.config.get('key')
                client = DeepSeekAPIClient(api_key=api_key)
                return client.get_available_models()
            elif self.provider in ['openai', 'anthropic', 'gemini', 'zhipu', 'hunyuan', 'qwen', 'doubao', 'moonshot']:
                # 对于其他提供商，返回一些通用的模型选项
                return self._get_other_provider_models()
            else:
                return False, [], f"未知提供商: {self.provider}"
                
        except Exception as e:
            print(f"[API管理器] 获取模型列表失败: {str(e)}")
            return False, [], f"获取失败: {str(e)}"
    
    def _get_other_provider_models(self):
        """获取其他提供商的默认模型列表"""
        provider_models = {
            'openai': [
                {'value': 'gpt-3.5-turbo', 'label': 'GPT-3.5 Turbo'},
                {'value': 'gpt-4', 'label': 'GPT-4'},
                {'value': 'gpt-4-turbo', 'label': 'GPT-4 Turbo'},
                {'value': 'gpt-4o', 'label': 'GPT-4o'}
            ],
            'anthropic': [
                {'value': 'claude-3-haiku-20240307', 'label': 'Claude 3 Haiku'},
                {'value': 'claude-3-sonnet-20240229', 'label': 'Claude 3 Sonnet'},
                {'value': 'claude-3-opus-20240229', 'label': 'Claude 3 Opus'}
            ],
            'gemini': [
                {'value': 'gemini-pro', 'label': 'Gemini Pro'},
                {'value': 'gemini-pro-vision', 'label': 'Gemini Pro Vision'},
                {'value': 'gemini-1.5-pro', 'label': 'Gemini 1.5 Pro'}
            ],
            'zhipu': [
                {'value': 'glm-4', 'label': 'GLM-4'},
                {'value': 'glm-4v', 'label': 'GLM-4V'},
                {'value': 'glm-3-turbo', 'label': 'GLM-3 Turbo'}
            ],
            'hunyuan': [
                {'value': 'hunyuan-pro', 'label': '混元-Pro'},
                {'value': 'hunyuan-standard', 'label': '混元-Standard'},
                {'value': 'hunyuan-lite', 'label': '混元-Lite'}
            ],
            'qwen': [
                {'value': 'qwen-max', 'label': '通义千问-Max'},
                {'value': 'qwen-plus', 'label': '通义千问-Plus'},
                {'value': 'qwen-turbo', 'label': '通义千问-Turbo'}
            ],
            'doubao': [
                {'value': 'doubao-pro-4k', 'label': '豆包-Pro-4K'},
                {'value': 'doubao-lite-4k', 'label': '豆包-Lite-4K'},
                {'value': 'doubao-pro-32k', 'label': '豆包-Pro-32K'}
            ],
            'moonshot': [
                {'value': 'moonshot-v1-8k', 'label': 'Moonshot v1 8K'},
                {'value': 'moonshot-v1-32k', 'label': 'Moonshot v1 32K'},
                {'value': 'moonshot-v1-128k', 'label': 'Moonshot v1 128K'}
            ]
        }
        
        models = provider_models.get(self.provider, [])
        return True, models, f"获取 {self.provider} 默认模型列表"


class NonLLMNameGenerator:
    """非LLM模式的命名生成器（单例模式）"""
    
    _instance = None
    _counter = 1
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(NonLLMNameGenerator, cls).__new__(cls)
        return cls._instance
    
    def get_next_name(self):
        """获取下一个表格名称"""
        name = f"合并表{self._counter}"
        self._counter += 1
        return name
    
    def reset_counter(self):
        """重置计数器"""
        self._counter = 1


def create_api_manager_from_local_storage():
    """从localStorage配置创建API管理器（在后端模拟前端配置）"""
    try:
        # 这里应该从某种配置存储中读取配置
        # 暂时使用默认配置
        default_config = {
            'provider': 'none',
            'url': '',
            'key': '',
            'model': 'none'
        }
        return APIManager(default_config)
    except Exception as e:
        print(f"[API管理器] 创建管理器失败: {str(e)}")
        return APIManager()


if __name__ == "__main__":
    # 测试代码
    test_columns = ['姓名', '年龄', '部门', '职位', '工资']
    test_data = [
        ['张三', 28, '技术部', '工程师', 8000],
        ['李四', 32, '销售部', '销售经理', 12000]
    ]
    
    print("=== API管理器测试 ===")
    
    # 测试非LLM模式
    print("\n1. 测试非LLM模式:")
    config = {'provider': 'none'}
    manager = APIManager(config)
    for i in range(3):
        success, name, message = manager.generate_table_name(test_columns)
        print(f"生成结果 {i+1}: {name} ({message})")
    
    # 测试OLLAMA模式
    print("\n2. 测试OLLAMA模式:")
    ollama_config = {
        'provider': 'ollama',
        'url': 'http://localhost:11434',
        'model': 'gemma2'
    }
    ollama_manager = APIManager(ollama_config)
    success, name, message = ollama_manager.generate_table_name(test_columns, test_data)
    print(f"OLLAMA结果: {name} ({message})")
    
    # 测试连接
    success, message = ollama_manager.test_connection()
    print(f"OLLAMA连接测试: {success} - {message}")
    
    # 测试非LLM单例计数器
    print("\n3. 测试非LLM单例模式:")
    generator = NonLLMNameGenerator()
    for i in range(3):
        name = generator.get_next_name()
        print(f"单例生成结果 {i+1}: {name}")
