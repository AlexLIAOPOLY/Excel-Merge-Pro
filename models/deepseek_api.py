"""
DeepSeek API集成模块
用于调用DeepSeek API生成合并表格的简短名称
"""

import requests
import time
import os
from datetime import datetime

class DeepSeekAPIClient:
    """DeepSeek API客户端"""
    
    def __init__(self, api_key=None, base_url=None, model=None):
        """
        初始化DeepSeek API客户端
        Args:
            api_key (str): DeepSeek API密钥
            base_url (str): API基础地址
            model (str): 使用的模型名称
        """
        self.api_key = api_key or os.getenv('DEEPSEEK_API_KEY')
        if not self.api_key:
            print("[DeepSeek API] 警告: API密钥未配置，某些功能可能不可用")
        
        self.base_url = base_url or "https://api.deepseek.com/v1/chat/completions"
        self.model = model or "deepseek-chat"  # 使用最便宜的deepseek-chat模型
        
    def generate_table_name(self, column_names, sample_data=None, max_retries=3):
        """
        基于表格结构和样本数据生成简短的表格名称
        
        Args:
            column_names (list): 表格列名列表
            sample_data (list): 样本数据，可选
            max_retries (int): 最大重试次数
            
        Returns:
            tuple: (success, table_name, message)
        """
        print(f"[DeepSeek API] 开始生成表格名称，列名: {column_names}")
        
        try:
            # 构建提示词
            prompt = self._build_naming_prompt(column_names, sample_data)
            
            # 调用API
            for attempt in range(max_retries):
                try:
                    response = self._call_api(prompt)
                    
                    if response:
                        table_name = self._extract_table_name(response)
                        if table_name:
                            print(f"[DeepSeek API] 成功生成表格名称: {table_name}")
                            return True, table_name, "表格命名成功"
                        else:
                            print(f"[DeepSeek API] 第{attempt+1}次尝试：无法从响应中提取表格名称")
                    
                    if attempt < max_retries - 1:
                        print(f"[DeepSeek API] 第{attempt+1}次尝试失败，2秒后重试...")
                        time.sleep(2)
                    
                except Exception as e:
                    print(f"[DeepSeek API] 第{attempt+1}次尝试出错: {str(e)}")
                    if attempt < max_retries - 1:
                        time.sleep(2)
                    else:
                        raise e
            
            # 所有重试都失败，使用默认命名
            fallback_name = self._generate_fallback_name(column_names)
            print(f"[DeepSeek API] API调用失败，使用默认命名: {fallback_name}")
            return True, fallback_name, "使用默认命名规则"
            
        except Exception as e:
            print(f"[DeepSeek API] 生成表格名称时出错: {str(e)}")
            # 出错时也提供一个默认名称
            fallback_name = self._generate_fallback_name(column_names)
            return False, fallback_name, f"API调用失败: {str(e)}"
    
    def _build_naming_prompt(self, column_names, sample_data):
        """构建AI提示词"""
        # 基础提示词
        prompt = """请根据以下Excel表格的列名信息，为这个表格生成一个简短、准确的中文名称。

要求：
1. 名称要简洁明了，最多8个字
2. 能准确反映表格的主要内容和用途
3. 使用常见的业务术语
4. 只返回表格名称，不要其他解释

表格列名："""

        # 添加列名信息
        column_str = "、".join(column_names[:10])  # 只取前10个列名避免过长
        prompt += f"\n{column_str}"
        
        # 如果有样本数据，添加样本信息
        if sample_data and len(sample_data) > 0:
            prompt += "\n\n样本数据（前3行）："
            for i, row in enumerate(sample_data[:3]):
                if i < 3:
                    row_str = "、".join(str(value)[:20] for value in row.values() if value)[:100]
                    prompt += f"\n第{i+1}行: {row_str}"
        
        prompt += "\n\n表格名称："
        
        return prompt
    
    def _call_api(self, prompt):
        """调用DeepSeek API"""
        if not self.api_key:
            raise ValueError("API密钥未配置")
        
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            "max_tokens": 50,  # 表格名称很短，不需要太多token
            "temperature": 0.3,  # 较低的温度确保结果稳定
            "stream": False
        }
        
        print(f"[DeepSeek API] 发送请求到: {self.base_url}")
        print(f"[DeepSeek API] 使用模型: {self.model}")
        
        try:
            response = requests.post(self.base_url, headers=headers, json=payload, timeout=30)
            
            print(f"[DeepSeek API] 响应状态码: {response.status_code}")
            
            if response.status_code == 200:
                result = response.json()
                if 'choices' in result and len(result['choices']) > 0:
                    content = result['choices'][0]['message']['content']
                    print(f"[DeepSeek API] API响应成功: {content}")
                    return content.strip()
                else:
                    print(f"[DeepSeek API] 响应格式异常: {result}")
                    return None
            else:
                error_text = response.text[:200] if response.text else "无错误信息"
                print(f"[DeepSeek API] API请求失败: {response.status_code} - {error_text}")
                raise requests.RequestException(f"HTTP {response.status_code}: {error_text}")
                
        except requests.exceptions.Timeout:
            print("[DeepSeek API] 请求超时")
            raise requests.RequestException("请求超时，请检查网络连接")
        except requests.exceptions.ConnectionError:
            print("[DeepSeek API] 连接错误")
            raise requests.RequestException("连接失败，请检查网络或API地址")
        except requests.exceptions.RequestException as e:
            print(f"[DeepSeek API] 请求异常: {str(e)}")
            raise e
    
    def _extract_table_name(self, response):
        """从API响应中提取表格名称"""
        if not response:
            return None
        
        # 清理响应内容
        name = response.strip()
        
        # 移除常见的前缀和后缀
        prefixes_to_remove = ['表格名称：', '名称：', '表格：', '建议：', '推荐：']
        for prefix in prefixes_to_remove:
            if name.startswith(prefix):
                name = name[len(prefix):].strip()
        
        # 移除引号
        if name.startswith('"') and name.endswith('"'):
            name = name[1:-1]
        if name.startswith("'") and name.endswith("'"):
            name = name[1:-1]
        
        # 移除换行符和多余空格
        name = ' '.join(name.split())
        
        # 长度检查
        if len(name) > 20:
            name = name[:20]
        
        # 确保名称不为空
        if not name or len(name.strip()) == 0:
            return None
        
        return name
    
    def _generate_fallback_name(self, column_names):
        """生成默认表格名称"""
        current_time = datetime.now()
        date_str = current_time.strftime("%m%d")
        time_str = current_time.strftime("%H%M")
        
        # 根据列名推断表格类型
        name_keywords = {
            '员工': ['员工', '姓名', '工号', '部门'],
            '客户': ['客户', '公司', '联系人'],
            '订单': ['订单', '商品', '数量', '金额'],
            '库存': ['库存', '商品', '数量', '仓库'],
            '财务': ['金额', '收入', '支出', '预算'],
            '项目': ['项目', '进度', '负责人'],
            '销售': ['销售', '业绩', '客户'],
            '产品': ['产品', '型号', '价格'],
            '合同': ['合同', '签订', '甲方', '乙方']
        }
        
        # 检查列名中是否包含关键词
        column_text = ''.join(column_names).lower()
        
        for table_type, keywords in name_keywords.items():
            if any(keyword in column_text for keyword in keywords):
                return f"{table_type}表_{date_str}_{time_str}"
        
        # 如果没有匹配的关键词，使用通用名称
        return f"合并表_{date_str}_{time_str}"
    
    def test_connection(self):
        """测试API连接"""
        try:
            # 检查基本配置
            if not self.api_key:
                return False, "API密钥未配置"
            
            if not self.base_url:
                return False, "API地址未配置"
            
            print(f"[DeepSeek API] 测试连接到: {self.base_url}")
            print(f"[DeepSeek API] 使用模型: {self.model}")
            
            test_prompt = "请回复'连接测试成功'"
            response = self._call_api(test_prompt)
            
            if response and ("成功" in response or "连接" in response or len(response.strip()) > 0):
                return True, f"API连接正常，响应: {response[:50]}..."
            elif response:
                return True, f"API连接正常，收到响应: {response[:30]}..."
            else:
                return False, "API无响应或响应为空"
                
        except Exception as e:
            error_msg = str(e)
            print(f"[DeepSeek API] 连接测试异常: {error_msg}")
            return False, f"连接测试失败: {error_msg}"


def generate_smart_table_name(column_names, sample_data=None, api_key=None):
    """
    便捷函数：为表格生成智能名称
    
    Args:
        column_names (list): 列名列表
        sample_data (list): 样本数据，可选
        api_key (str): API密钥，可选
    
    Returns:
        str: 生成的表格名称
    """
    client = DeepSeekAPIClient(api_key)
    success, table_name, message = client.generate_table_name(column_names, sample_data)
    
    print(f"[智能命名] {message}: {table_name}")
    return table_name


if __name__ == "__main__":
    # 测试代码
    print("=== DeepSeek API测试 ===")
    
    # 测试列名
    test_columns = ["员工编号", "姓名", "部门", "职位", "入职日期", "工资"]
    test_data = [
        {"员工编号": "E001", "姓名": "张三", "部门": "技术部", "职位": "工程师", "入职日期": "2023-01-01", "工资": "8000"},
        {"员工编号": "E002", "姓名": "李四", "部门": "销售部", "职位": "销售经理", "入职日期": "2023-02-01", "工资": "12000"},
    ]
    
    client = DeepSeekAPIClient()
    
    # 测试连接
    print("\n1. 测试API连接...")
    connection_success, message = client.test_connection()
    print(f"连接测试结果: {connection_success} - {message}")
    
    # 测试表格命名
    print("\n2. 测试表格命名...")
    success, table_name, message = client.generate_table_name(test_columns, test_data)
    print(f"命名结果: {success} - {table_name} ({message})")