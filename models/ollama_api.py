"""
OLLAMA API集成模块
用于调用本地OLLAMA API生成合并表格的简短名称
"""

import requests
import json
from datetime import datetime


class OllamaAPIClient:
    """OLLAMA API客户端"""
    
    def __init__(self, base_url="http://localhost:11434", model="gemma2"):
        """
        初始化OLLAMA API客户端
        
        Args:
            base_url (str): OLLAMA服务器地址
            model (str): 使用的模型名称
        """
        self.base_url = base_url.rstrip('/')
        self.model = model
        print(f"[OLLAMA API] 初始化客户端，URL: {self.base_url}, 模型: {self.model}")

    def generate_table_name(self, column_names, sample_data=None, max_retries=3):
        """
        基于表格结构和样本数据生成简短的表格名称
        
        Args:
            column_names (list): 表格列名列表
            sample_data (list, optional): 样本数据（前几行数据）
            max_retries (int): 最大重试次数
            
        Returns:
            tuple: (成功标志, 表格名称, 消息)
        """
        print(f"[OLLAMA API] 开始生成表格名称，列名: {column_names}")
        
        try:
            # 构建提示词
            prompt = self._build_naming_prompt(column_names, sample_data)
            
            # 调用OLLAMA API进行多次重试
            for attempt in range(max_retries):
                try:
                    result = self._call_ollama_api(prompt)
                    if result and result.get('response'):
                        # 提取表格名称
                        table_name = self._extract_table_name(result['response'])
                        if table_name:
                            print(f"[OLLAMA API] 成功生成表格名称: {table_name}")
                            return True, table_name, "OLLAMA生成成功"
                        else:
                            print(f"[OLLAMA API] 第{attempt+1}次尝试：无法从响应中提取表格名称")
                    else:
                        print(f"[OLLAMA API] 第{attempt+1}次尝试失败，2秒后重试...")
                        if attempt < max_retries - 1:
                            import time
                            time.sleep(2)
                except Exception as e:
                    print(f"[OLLAMA API] 第{attempt+1}次尝试出错: {str(e)}")
                    if attempt < max_retries - 1:
                        import time
                        time.sleep(2)
            
            # 所有尝试都失败，使用默认命名
            fallback_name = self._generate_fallback_name(column_names)
            print(f"[OLLAMA API] API调用失败，使用默认命名: {fallback_name}")
            return True, fallback_name, "使用默认命名规则"
            
        except Exception as e:
            print(f"[OLLAMA API] 生成表格名称时出错: {str(e)}")
            # 生成默认名称作为后备方案
            fallback_name = self._generate_fallback_name(column_names)
            return False, fallback_name, f"API调用失败: {str(e)}"
    
    def _build_naming_prompt(self, column_names, sample_data):
        """构建AI提示词"""
        # 基础提示词
        prompt = f"""请根据以下表格的列名，为这个表格起一个简洁且准确的中文名称（4-8个字）。

表格列名: {', '.join(column_names)}
"""

        # 如果有样本数据，添加到提示词中
        if sample_data and len(sample_data) > 0:
            prompt += f"\n前几行样本数据:\n"
            for i, row in enumerate(sample_data[:3]):  # 只显示前3行
                row_text = ', '.join([str(cell)[:20] for cell in row])  # 每个单元格最多20字符
                prompt += f"第{i+1}行: {row_text}\n"
        
        prompt += """
命名要求：
1. 名称要简洁明了，4-8个汉字
2. 体现表格的主要内容或用途
3. 避免使用"表格"、"数据"等泛化词汇
4. 只返回表格名称，不要任何解释

表格名称："""
        
        return prompt

    def _call_ollama_api(self, prompt):
        """调用OLLAMA API"""
        try:
            url = f"{self.base_url}/api/generate"
            
            payload = {
                "model": self.model,
                "prompt": prompt,
                "stream": False,
                "options": {
                    "temperature": 0.7,
                    "top_p": 0.9,
                    "max_tokens": 50
                }
            }
            
            print(f"[OLLAMA API] 发送请求到: {url}")
            print(f"[OLLAMA API] 使用模型: {self.model}")
            
            headers = {'Content-Type': 'application/json'}
            response = requests.post(url, json=payload, headers=headers, timeout=30)
            
            print(f"[OLLAMA API] 响应状态码: {response.status_code}")
            
            if response.status_code == 200:
                result = response.json()
                if 'response' in result:
                    print(f"[OLLAMA API] API响应成功: {result['response'][:100]}")
                    return result
                else:
                    print(f"[OLLAMA API] 响应格式异常: {result}")
                    return None
            else:
                error_text = response.text
                print(f"[OLLAMA API] API请求失败: {response.status_code} - {error_text}")
                return None
                
        except requests.exceptions.Timeout:
            print("[OLLAMA API] 请求超时")
            return None
        except requests.exceptions.ConnectionError:
            print("[OLLAMA API] 连接错误，请确保OLLAMA服务正在运行")
            return None
        except Exception as e:
            print(f"[OLLAMA API] 请求异常: {str(e)}")
            return None

    def _extract_table_name(self, response_text):
        """从AI响应中提取表格名称"""
        if not response_text:
            return None
            
        # 清理响应文本
        text = response_text.strip()
        
        # 移除常见的前缀
        prefixes = ['表格名称：', '表格名称:', '名称：', '名称:', '表名：', '表名:']
        for prefix in prefixes:
            if text.startswith(prefix):
                text = text[len(prefix):].strip()
                break
        
        # 移除引号
        text = text.strip('"\'""''')
        
        # 取第一行作为表格名称（防止多行输出）
        lines = text.split('\n')
        if lines:
            table_name = lines[0].strip()
            
            # 确保名称长度合理
            if 2 <= len(table_name) <= 20:
                return table_name
        
        return None

    def _generate_fallback_name(self, column_names):
        """生成默认表格名称"""
        current_time = datetime.now()
        date_str = current_time.strftime("%m%d")
        time_str = current_time.strftime("%H%M")
        
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
                return f"{table_type}_{date_str}"
        
        # 如果没有匹配的关键词，使用默认命名
        return f"合并表_{date_str}_{time_str}"

    def test_connection(self):
        """测试与OLLAMA服务器的连接"""
        try:
            print(f"[OLLAMA API] 测试连接到: {self.base_url}")
            print(f"[OLLAMA API] 使用模型: {self.model}")
            
            url = f"{self.base_url}/api/generate"
            test_payload = {
                "model": self.model,
                "prompt": "你好",
                "stream": False,
                "options": {"max_tokens": 10}
            }
            
            response = requests.post(url, json=test_payload, timeout=10)
            
            if response.status_code == 200:
                return True, "连接成功"
            else:
                return False, f"服务器响应错误: {response.status_code}"
                
        except requests.exceptions.ConnectionError:
            error_msg = "连接失败，请确保OLLAMA服务正在运行"
            print(f"[OLLAMA API] 连接测试异常: {error_msg}")
            return False, error_msg
        except Exception as e:
            error_msg = f"连接测试失败: {str(e)}"
            print(f"[OLLAMA API] 连接测试异常: {error_msg}")
            return False, error_msg

    def get_available_models(self):
        """获取本地OLLAMA的可用模型列表"""
        try:
            url = f"{self.base_url}/api/tags"
            
            print(f"[OLLAMA API] 获取模型列表: {url}")
            
            response = requests.get(url, timeout=10)
            
            if response.status_code == 200:
                data = response.json()
                if 'models' in data:
                    models = []
                    for model in data['models']:
                        model_name = model.get('name', '')
                        model_size = model.get('size', 0)
                        size_gb = round(model_size / (1024**3), 1) if model_size > 0 else 0
                        
                        if model_name:
                            label = f"{model_name}"
                            if size_gb > 0:
                                label += f" ({size_gb}GB)"
                            
                            models.append({
                                'value': model_name,
                                'label': label
                            })
                    
                    if models:
                        print(f"[OLLAMA API] 找到 {len(models)} 个本地模型")
                        return True, models, f"找到 {len(models)} 个本地模型"
                    else:
                        # 没有找到模型，返回常用模型建议
                        suggested_models = [
                            {'value': 'gemma2', 'label': 'gemma2 (推荐)'},
                            {'value': 'llama3.1', 'label': 'llama3.1 (Meta)'},
                            {'value': 'qwen2.5', 'label': 'qwen2.5 (阿里)'},
                            {'value': 'mistral', 'label': 'mistral (轻量)'}
                        ]
                        return True, suggested_models, "未找到本地模型，显示推荐模型"
                else:
                    suggested_models = [
                        {'value': 'gemma2', 'label': 'gemma2 (推荐)'},
                        {'value': 'llama3.1', 'label': 'llama3.1 (Meta)'},
                        {'value': 'qwen2.5', 'label': 'qwen2.5 (阿里)'},
                        {'value': 'mistral', 'label': 'mistral (轻量)'}
                    ]
                    return True, suggested_models, "API响应格式异常，显示推荐模型"
            else:
                suggested_models = [
                    {'value': 'gemma2', 'label': 'gemma2 (推荐)'},
                    {'value': 'llama3.1', 'label': 'llama3.1 (Meta)'},
                    {'value': 'qwen2.5', 'label': 'qwen2.5 (阿里)'},
                    {'value': 'mistral', 'label': 'mistral (轻量)'}
                ]
                return False, suggested_models, f"API请求失败: {response.status_code}"
                
        except requests.exceptions.ConnectionError:
            error_msg = "连接失败，请确保OLLAMA服务正在运行"
            suggested_models = [
                {'value': 'gemma2', 'label': 'gemma2 (推荐)'},
                {'value': 'llama3.1', 'label': 'llama3.1 (Meta)'},
                {'value': 'qwen2.5', 'label': 'qwen2.5 (阿里)'},
                {'value': 'mistral', 'label': 'mistral (轻量)'}
            ]
            print(f"[OLLAMA API] 获取模型列表失败: {error_msg}")
            return False, suggested_models, error_msg
        except Exception as e:
            error_msg = f"获取模型列表失败: {str(e)}"
            suggested_models = [
                {'value': 'gemma2', 'label': 'gemma2 (推荐)'},
                {'value': 'llama3.1', 'label': 'llama3.1 (Meta)'},
                {'value': 'qwen2.5', 'label': 'qwen2.5 (阿里)'},
                {'value': 'mistral', 'label': 'mistral (轻量)'}
            ]
            print(f"[OLLAMA API] 获取模型列表异常: {error_msg}")
            return False, suggested_models, error_msg


def generate_smart_table_name(column_names, sample_data=None, base_url="http://localhost:11434", model="gemma2"):
    """
    便捷函数：为表格生成智能名称
    
    Args:
        column_names (list): 列名列表
        sample_data (list): 样本数据
        base_url (str): OLLAMA服务器地址
        model (str): 使用的模型名称
        
    Returns:
        str: 生成的表格名称
    """
    client = OllamaAPIClient(base_url, model)
    success, table_name, message = client.generate_table_name(column_names, sample_data)
    
    print(f"[智能命名] {message}: {table_name}")
    return table_name


if __name__ == "__main__":
    # 测试代码
    test_columns = ['姓名', '年龄', '部门', '职位', '工资']
    test_data = [
        ['张三', 28, '技术部', '工程师', 8000],
        ['李四', 32, '销售部', '销售经理', 12000],
        ['王五', 25, '人事部', '专员', 6000]
    ]
    
    print("=== OLLAMA API测试 ===")
    
    # 测试表格命名
    client = OllamaAPIClient()
    
    # 测试连接
    success, message = client.test_connection()
    print(f"连接测试: {success} - {message}")
    
    if success:
        # 测试表格命名功能
        success, table_name, message = client.generate_table_name(test_columns, test_data)
        print(f"命名结果: {success} - {table_name} ({message})")