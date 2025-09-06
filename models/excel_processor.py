try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False
    pd = None
import os
import time
import threading
from datetime import datetime
from difflib import SequenceMatcher
import hashlib
from functools import lru_cache
from contextlib import contextmanager
from models.database import db, TableData, TableSchema, UploadHistory, TableGroup, ColumnMapping
from models.deepseek_api import DeepSeekAPIClient
from models.api_manager import APIManager, NonLLMNameGenerator
from models.config_storage import get_api_config

class UniversalExcelProcessor:
    """通用Excel表格处理器，支持任意格式的表格合并和智能分组"""
    
    # 相似度阈值配置
    EXACT_MATCH_THRESHOLD = 1.0      # 完全匹配
    HIGH_SIMILARITY_THRESHOLD = 0.95 # 高相似度 - 提高阈值以便更好地合并相似表格
    MIN_SIMILARITY_THRESHOLD = 0.85  # 最低相似度 - 提高阈值避免错误合并
    
    # 并发安全控制
    _creation_lock = threading.Lock()
    _fingerprint_cache = {}
    _similarity_cache = {}
    
    # 进度跟踪
    _current_progress = {'stage': '', 'percent': 0, 'message': ''}
    
    @classmethod
    def clear_cache(cls):
        """清理缓存"""
        cls._fingerprint_cache.clear()
        cls._similarity_cache.clear()
    
    @classmethod
    def get_progress(cls):
        """获取当前处理进度"""
        return cls._current_progress.copy()
    
    @classmethod
    def _update_progress(cls, stage, percent, message):
        """更新处理进度"""
        cls._current_progress.update({
            'stage': stage,
            'percent': percent,
            'message': message
        })
    
    @staticmethod
    @contextmanager
    def database_transaction():
        """安全的数据库事务管理器"""
        try:
            yield db.session
            db.session.commit()
        except Exception as e:
            db.session.rollback()
            raise e
    
    @staticmethod
    def detect_header_row(df, max_check_rows=5):
        """智能检测表头行位置"""
        print("[系统] 开始检测表头位置...")
        
        for i in range(min(max_check_rows, len(df))):
            row = df.iloc[i]
            # 检查这一行是否像表头（非空值较多，且包含文字）
            non_null_count = row.count()
            text_count = sum(1 for val in row if isinstance(val, str) and len(str(val).strip()) > 0)
            
            # 如果非空值较多且大部分是文字，可能是表头
            if non_null_count >= len(row) * 0.5 and text_count >= non_null_count * 0.7:
                print(f"[系统] 检测到表头可能在第 {i+1} 行")
                return i
        
        print("[系统] 未检测到明显表头，使用第1行作为表头")
        return 0
    
    @staticmethod
    def clean_column_names(columns):
        """清理列名 - 增强版本，支持更智能的标准化和重复列名处理"""
        import re
        import unicodedata
        cleaned = []
        seen_names = {}  # 跟踪已使用的列名
        
        for i, col in enumerate(columns):
            if pd.isna(col):
                col_str = f"未命名列_{i+1}"
            else:
                col_str = str(col)
                
                # 1. Unicode标准化（处理各种空格字符）
                col_str = unicodedata.normalize('NFKC', col_str)
                
                # 2. 去除首尾空格
                col_str = col_str.strip()
                
                # 3. 将各种空白字符（包括中文全角空格、制表符等）替换为单个半角空格
                col_str = re.sub(r'[\s\u00A0\u2000-\u200B\u2028\u2029\u3000]+', ' ', col_str)
                
                # 4. 去除特殊的不可见字符
                col_str = re.sub(r'[\u200E\u200F\uFEFF]', '', col_str)
                
                # 5. 再次去除首尾空格
                col_str = col_str.strip()
                
                # 6. 如果清理后为空，使用默认名称
                if not col_str:
                    col_str = f"未命名列_{i+1}"
            
            # 7. 处理重复列名
            original_name = col_str
            counter = 1
            while col_str in seen_names:
                col_str = f"{original_name}_{counter}"
                counter += 1
                if counter > 100:  # 防止无限循环
                    col_str = f"{original_name}_{i+1}"
                    break
            
            seen_names[col_str] = True
            cleaned.append(col_str)
        
        return cleaned
    
    @staticmethod
    def update_table_schema(columns):
        """更新表格结构"""
        print(f"[系统] 更新表格结构，共 {len(columns)} 列")
        
        for i, col_name in enumerate(columns):
            # 检查列是否已存在
            existing = TableSchema.query.filter_by(column_name=col_name).first()
            if not existing:
                schema = TableSchema(
                    column_name=col_name,
                    column_type='text',
                    column_order=i,
                    is_active=True
                )
                db.session.add(schema)
                print(f"[系统] 添加新列: {col_name}")
        
        try:
            db.session.commit()
        except Exception as e:
            print(f"[错误] 更新表格结构时出错: {str(e)}")
            db.session.rollback()
    
    @staticmethod
    def get_current_schema():
        """获取当前表格结构"""
        schema = TableSchema.query.filter_by(is_active=True).order_by(TableSchema.column_order).all()
        return [s.column_name for s in schema]
    
    @staticmethod
    def process_excel_file(file_path, filename):
        """处理Excel文件，支持任意格式"""
        print(f"[系统] 开始处理Excel文件: {filename}")
        
        try:
            # 统一使用openpyxl处理所有Excel文件
            df_raw = pd.read_excel(file_path, engine='openpyxl', header=None)
            
            if df_raw.empty:
                return False, "文件为空", 0
            
            print(f"[系统] 文件读取成功，原始数据 {len(df_raw)} 行 x {len(df_raw.columns)} 列")
            
            # 智能检测表头位置
            header_row = UniversalExcelProcessor.detect_header_row(df_raw)
            
            # 重新读取文件，使用检测到的表头
            df = pd.read_excel(file_path, engine='openpyxl', header=header_row)
            
            # 清理数据
            df = df.dropna(how='all')  # 删除全空行
            df = df.dropna(axis=1, how='all')  # 删除全空列
            
            if df.empty:
                return False, "文件中没有有效数据", 0
            
            # 清理列名
            df.columns = UniversalExcelProcessor.clean_column_names(df.columns)
            
            print(f"[系统] 处理后数据: {len(df)} 行 x {len(df.columns)} 列")
            print(f"[系统] 检测到的列名: {list(df.columns)}")
            
            # 更新表格结构
            UniversalExcelProcessor.update_table_schema(df.columns)
            
            # 导入数据
            imported_count = 0
            for index, row in df.iterrows():
                # 跳过全空行
                if row.isna().all():
                    continue
                
                # 创建数据字典
                row_dict = {}
                for col_name in df.columns:
                    value = row[col_name]
                    if pd.notna(value):
                        row_dict[col_name] = str(value)
                    else:
                        row_dict[col_name] = ""
                
                # 如果整行都是空的，跳过
                if not any(v.strip() for v in row_dict.values() if v):
                    continue
                
                # 创建数据记录
                table_data = TableData()
                table_data.source_file = filename
                table_data.set_data(row_dict)
                
                try:
                    db.session.add(table_data)
                    imported_count += 1
                except Exception as e:
                    print(f"[错误] 导入第 {index+1} 行数据时出错: {str(e)}")
                    continue
            
            db.session.commit()
            print(f"[系统] 成功导入 {imported_count} 条数据")
            
            # 记录上传历史
            history = UploadHistory(
                filename=filename,
                rows_imported=imported_count,
                status='success'
            )
            history.set_columns(list(df.columns))
            db.session.add(history)
            db.session.commit()
            
            return True, "导入成功", imported_count
            
        except Exception as e:
            print(f"[错误] 处理文件时发生错误: {str(e)}")
            db.session.rollback()
            
            # 记录错误历史
            history = UploadHistory(
                filename=filename,
                rows_imported=0,
                status='failed',
                error_message=str(e)
            )
            db.session.add(history)
            db.session.commit()
            
            return False, str(e), 0
    
    @staticmethod
    def get_all_data():
        """获取所有数据，按创建时间排序"""
        data = TableData.query.order_by(TableData.created_at.asc()).all()
        return [item.to_dict() for item in data]
    
    @staticmethod
    def get_data_stats():
        """获取数据统计"""
        total_count = TableData.query.count()
        
        # 统计来源文件数 - 改进逻辑，排除空值和只有扩展名的情况
        source_files = db.session.query(TableData.source_file).distinct().all()
        valid_files = []
        for f in source_files:
            filename = f[0]
            if filename and filename.strip():
                # 排除只有扩展名的情况（如"xlsx", "xls"等）
                if '.' in filename and len(filename.split('.')[0]) > 0:
                    valid_files.append(filename)
                elif filename not in ['xlsx', 'xls', '手动添加']:
                    valid_files.append(filename)
        
        file_count = len(valid_files)
        
        # 统计列数 - 优先使用分组数据
        schema_count = 0
        groups = TableGroup.query.all()
        if groups:
            # 使用最新分组的列数
            latest_group = groups[-1]
            schema_count = latest_group.column_count
        else:
            # 兼容旧数据
            schema_count = TableSchema.query.filter_by(is_active=True).count()
        
        return {
            'total_records': total_count,
            'source_files': file_count,
            'total_columns': schema_count,
            'last_update': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
    
    @staticmethod
    def update_record(record_id, data):
        """更新记录"""
        record = TableData.query.get(record_id)
        if record:
            current_data = record.get_data()
            current_data.update(data)
            record.set_data(current_data)
            record.updated_at = datetime.utcnow()
            db.session.commit()
            return True
        return False
    
    @staticmethod
    def delete_record(record_id):
        """删除记录"""
        record = TableData.query.get(record_id)
        if record:
            db.session.delete(record)
            db.session.commit()
            return True
        return False
    
    @staticmethod
    def clear_all_data():
        """清空所有数据"""
        TableData.query.delete()
        TableSchema.query.delete()
        db.session.commit()
    
    @staticmethod
    def add_column(column_name, insert_position=None):
        """添加新列，支持指定位置插入"""
        existing = TableSchema.query.filter_by(column_name=column_name).first()
        if existing:
            return False, "列名已存在"
        
        try:
            if insert_position is None:
                # 添加到末尾
                max_order = db.session.query(db.func.max(TableSchema.column_order)).scalar() or 0
                new_order = max_order + 1
            else:
                # 在指定位置插入，需要调整后续列的顺序
                new_order = insert_position
                
                # 将指定位置及之后的列的order都加1
                columns_to_update = TableSchema.query.filter(
                    TableSchema.column_order >= insert_position,
                    TableSchema.is_active == True
                ).all()
                
                for col in columns_to_update:
                    col.column_order += 1
            
            # 创建新列
            new_column = TableSchema(
                column_name=column_name,
                column_type='text',
                column_order=new_order,
                is_active=True
            )
            
            db.session.add(new_column)
            db.session.commit()
            print(f"[系统] 成功添加列: {column_name} 在位置 {new_order}")
            return True, "列添加成功"
        except Exception as e:
            db.session.rollback()
            print(f"[错误] 添加列时出错: {str(e)}")
            return False, str(e)
    
    @staticmethod
    def delete_column(column_name):
        """删除列"""
        # 检查是否是最后一列
        active_columns = TableSchema.query.filter_by(is_active=True).count()
        if active_columns <= 1:
            return False, "不能删除最后一列"
        
        # 找到要删除的列
        column = TableSchema.query.filter_by(column_name=column_name, is_active=True).first()
        if not column:
            return False, "列不存在"
        
        try:
            # 软删除：设置为非活跃状态
            column.is_active = False
            
            # 从所有数据记录中删除该列的数据
            all_records = TableData.query.all()
            for record in all_records:
                data = record.get_data()
                if column_name in data:
                    del data[column_name]
                    record.set_data(data)
            
            db.session.commit()
            print(f"[系统] 成功删除列: {column_name}")
            return True, "列删除成功"
        except Exception as e:
            db.session.rollback()
            print(f"[错误] 删除列时出错: {str(e)}")
            return False, str(e)
    
    @staticmethod
    def rename_column(old_name, new_name):
        """重命名列"""
        # 检查旧列是否存在
        old_column = TableSchema.query.filter_by(column_name=old_name, is_active=True).first()
        if not old_column:
            return False, "原列名不存在"
        
        # 检查新列名是否已存在
        existing = TableSchema.query.filter_by(column_name=new_name, is_active=True).first()
        if existing:
            return False, "新列名已存在"
        
        try:
            # 更新列结构
            old_column.column_name = new_name
            
            # 更新所有数据记录中的字段名
            all_records = TableData.query.all()
            for record in all_records:
                data = record.get_data()
                if old_name in data:
                    data[new_name] = data.pop(old_name)
                    record.set_data(data)
            
            db.session.commit()
            print(f"[系统] 成功重命名列: {old_name} -> {new_name}")
            return True, "列重命名成功"
        except Exception as e:
            db.session.rollback()
            print(f"[错误] 重命名列时出错: {str(e)}")
            return False, str(e)
    
    @staticmethod
    def add_row(insert_after_id=None, row_data=None):
        """添加新行，支持指定位置插入"""
        try:
            # 获取当前表格结构
            schema = UniversalExcelProcessor.get_current_schema()
            
            if not schema:
                return False, "没有可用的表格结构", None
            
            # 创建行数据
            if row_data is None:
                row_data = {col: '' for col in schema}
            
            # 创建新的数据记录
            table_data = TableData()
            table_data.source_file = '手动添加'
            table_data.set_data(row_data)
            
            # 提交到数据库
            db.session.add(table_data)
            db.session.commit()
            
            # 如果指定了插入位置，我们通过设置创建时间来影响排序
            if insert_after_id:
                # 获取参考行的时间
                ref_row = TableData.query.get(insert_after_id)
                if ref_row:
                    # 设置新行的创建时间稍晚于参考行，但早于后续行
                    from datetime import datetime, timedelta
                    table_data.created_at = ref_row.created_at + timedelta(microseconds=1)
                    db.session.commit()
            
            print(f"[系统] 成功添加新行，ID: {table_data.id}, 插入在ID {insert_after_id} 之后")
            return True, "新行添加成功", table_data.id
            
        except Exception as e:
            db.session.rollback()
            print(f"[错误] 添加行时出错: {str(e)}")
            return False, str(e), None
    
    # ==================== 智能表格分组功能 ====================
    
    @classmethod
    def calculate_column_similarity(cls, columns1, columns2):
        """
        计算两个列名列表的相似度 - 智能版本，带缓存优化
        支持列顺序不同、空格差异、大小写差异的情况
        """
        if not columns1 or not columns2:
            return 0.0
        
        # 创建缓存键（确保顺序无关）
        key1 = '|'.join(sorted(str(col) for col in columns1))
        key2 = '|'.join(sorted(str(col) for col in columns2))
        cache_key = f"{key1}##{key2}" if key1 <= key2 else f"{key2}##{key1}"
        
        # 检查缓存
        if cache_key in cls._similarity_cache:
            return cls._similarity_cache[cache_key]
        
        # 首先清理和标准化列名
        cleaned_cols1 = [cls._normalize_column_name(col) for col in columns1]
        cleaned_cols2 = [cls._normalize_column_name(col) for col in columns2]
        
        # 计算相似度
        if len(cleaned_cols1) != len(cleaned_cols2):
            # 如果列数不同，使用集合相似度 + 长度惩罚
            set1 = set(cleaned_cols1)
            set2 = set(cleaned_cols2)
            
            # 计算Jaccard相似度
            intersection = len(set1.intersection(set2))
            union = len(set1.union(set2))
            jaccard_similarity = intersection / union if union > 0 else 0.0
            
            # 长度差异惩罚
            length_ratio = min(len(cleaned_cols1), len(cleaned_cols2)) / max(len(cleaned_cols1), len(cleaned_cols2))
            
            similarity = jaccard_similarity * length_ratio
        else:
            # 列数相同时，使用最优匹配算法
            similarity = cls._calculate_optimal_column_matching(cleaned_cols1, cleaned_cols2)
        
        # 缓存结果（限制缓存大小）
        if len(cls._similarity_cache) < 1000:
            cls._similarity_cache[cache_key] = similarity
        
        return similarity
    
    @staticmethod
    def _normalize_column_name(column_name):
        """标准化单个列名用于比较，包含智能别名识别"""
        if pd.isna(column_name):
            return ""
        
        # 转换为字符串并转小写
        normalized = str(column_name).lower()
        
        # 去除所有空白字符和特殊字符，只保留字母数字和中文字符
        import re
        normalized = re.sub(r'[^\w\u4e00-\u9fff]', '', normalized)
        
        # 智能别名识别 - 将常见的同义词标准化
        column_aliases = {
            # ID相关
            '编号': 'id', 'number': 'id', 'num': 'id', '序号': 'id', 'id': 'id',
            # 名称相关
            '名称': 'name', 'name': 'name', '姓名': 'name', '名字': 'name',
            # 日期相关
            '日期': 'date', 'date': 'date', '时间': 'date', 'time': 'date',
            '创建日期': 'createdate', '创建时间': 'createdate',
            '更新日期': 'updatedate', '更新时间': 'updatedate',
            # 状态相关
            '状态': 'status', 'status': 'status', '情况': 'status',
            # 备注相关
            '备注': 'remark', 'remark': 'remark', '说明': 'remark', 'description': 'remark',
            # 金额相关
            '金额': 'amount', 'amount': 'amount', '价格': 'amount', 'price': 'amount',
            # 数量相关
            '数量': 'quantity', 'quantity': 'quantity', '个数': 'quantity', '数目': 'quantity',
            # 部门相关
            '部门': 'department', 'department': 'department', 'dept': 'department',
            # 联系方式
            '电话': 'phone', 'phone': 'phone', '手机': 'phone', 'mobile': 'phone',
            '邮箱': 'email', 'email': 'email', '邮件': 'email', 'mail': 'email',
            # 地址相关
            '地址': 'address', 'address': 'address', '位置': 'address', 'location': 'address',
        }
        
        # 应用别名映射
        for alias, standard in column_aliases.items():
            if alias in normalized:
                normalized = normalized.replace(alias, standard)
        
        return normalized
    
    @staticmethod
    def _calculate_optimal_column_matching(cols1, cols2):
        """
        计算两个等长列表的最优匹配相似度
        使用匈牙利算法思想，找到最佳的列对应关系
        """
        n = len(cols1)
        if n == 0:
            return 1.0
        
        # 构建相似度矩阵
        similarity_matrix = []
        for col1 in cols1:
            row = []
            for col2 in cols2:
                # 计算单个列名的相似度
                if col1 == col2:
                    # 完全匹配
                    similarity = 1.0
                elif col1 and col2:
                    # 使用字符串相似度算法
                    similarity = SequenceMatcher(None, col1, col2).ratio()
                else:
                    similarity = 0.0
                row.append(similarity)
            similarity_matrix.append(row)
        
        # 使用简化的最优匹配算法
        if n <= 10:  # 对于小规模使用精确算法
            return UniversalExcelProcessor._find_best_matching(similarity_matrix)
        else:  # 对于大规模使用贪心算法
            return UniversalExcelProcessor._greedy_matching(similarity_matrix)
    
    @staticmethod
    def _find_best_matching(similarity_matrix):
        """找到最佳匹配（适用于小规模）"""
        import itertools
        n = len(similarity_matrix)
        best_score = 0.0
        
        # 尝试所有可能的排列
        for perm in itertools.permutations(range(n)):
            score = sum(similarity_matrix[i][perm[i]] for i in range(n)) / n
            best_score = max(best_score, score)
        
        return best_score
    
    @staticmethod
    def _greedy_matching(similarity_matrix):
        """贪心匹配算法（适用于大规模）"""
        n = len(similarity_matrix)
        used_cols = set()
        total_similarity = 0.0
        
        for i in range(n):
            best_match = -1
            best_sim = 0.0
            
            for j in range(n):
                if j not in used_cols and similarity_matrix[i][j] > best_sim:
                    best_sim = similarity_matrix[i][j]
                    best_match = j
            
            if best_match != -1:
                used_cols.add(best_match)
                total_similarity += best_sim
        
        return total_similarity / n
    
    @classmethod
    def generate_schema_fingerprint(cls, columns):
        """
        生成表头指纹用于快速匹配 - 增强版本，带缓存优化
        使用更智能的标准化方法，确保列顺序不同、空格差异等情况下指纹一致
        """
        if not columns:
            return hashlib.md5(b'').hexdigest()
        
        # 创建缓存键
        cache_key = '|'.join(str(col) for col in columns)
        
        # 检查缓存
        if cache_key in cls._fingerprint_cache:
            return cls._fingerprint_cache[cache_key]
        
        # 使用新的标准化方法，确保空格、大小写、特殊字符处理一致
        normalized_columns = []
        for col in columns:
            normalized = cls._normalize_column_name(col)
            if normalized:  # 只添加非空的标准化列名
                normalized_columns.append(normalized)
        
        # 排序确保顺序无关
        normalized_columns.sort()
        
        # 生成指纹
        fingerprint_text = '|'.join(normalized_columns)
        fingerprint = hashlib.md5(fingerprint_text.encode('utf-8')).hexdigest()
        
        # 缓存结果（限制缓存大小）
        if len(cls._fingerprint_cache) < 1000:
            cls._fingerprint_cache[cache_key] = fingerprint
        
        return fingerprint
    
    @classmethod
    def find_matching_table_group(cls, columns):
        """查找匹配的表格分组 - 强化版本，绝对避免重复分组"""
        print(f"[系统] 查找匹配的表格分组，列数: {len(columns)}")
        print(f"[系统] 输入列结构: {columns}")
        
        # 首先清理输入的列名
        cleaned_columns = cls.clean_column_names(columns)
        print(f"[系统] 清理后列结构: {cleaned_columns}")
        
        # 生成当前表头的指纹
        current_fingerprint = cls.generate_schema_fingerprint(cleaned_columns)
        print(f"[系统] 生成的指纹: {current_fingerprint}")
        
        # 首先查找完全匹配的分组 - 使用指纹和列数双重验证
        exact_matches = TableGroup.query.filter_by(
            schema_fingerprint=current_fingerprint,
            column_count=len(cleaned_columns)
        ).all()
        
        if exact_matches:
            if len(exact_matches) > 1:
                print(f"[警告] 发现 {len(exact_matches)} 个指纹重复的分组！需要合并")
                # 如果有多个完全匹配的分组，合并它们
                main_group = exact_matches[0]
                for i in range(1, len(exact_matches)):
                    duplicate_group = exact_matches[i]
                    print(f"[系统] 合并重复分组 {duplicate_group.id} 到 {main_group.id}")
                    
                    # 迁移数据
                    data_records = TableData.query.filter_by(table_group_id=duplicate_group.id).all()
                    for record in data_records:
                        record.table_group_id = main_group.id
                    
                    # 删除重复分组
                    db.session.delete(duplicate_group)
                
                db.session.commit()
                print(f"[系统] 重复分组合并完成，使用分组: {main_group.group_name}")
                return main_group, cls.EXACT_MATCH_THRESHOLD
            else:
                exact_match = exact_matches[0]
                print(f"[系统] 找到完全匹配的分组: {exact_match.group_name}")
                return exact_match, cls.EXACT_MATCH_THRESHOLD
        
        # 查找相似的分组 - 优化版本：使用join减少数据库查询次数
        groups_with_schemas = db.session.query(TableGroup, TableSchema).join(
            TableSchema, TableGroup.id == TableSchema.table_group_id
        ).filter(
            TableGroup.column_count == len(cleaned_columns),
            TableSchema.is_active == True
        ).order_by(TableGroup.id, TableSchema.column_order).all()
        
        # 按分组ID聚合列名
        group_columns_map = {}
        for group, schema in groups_with_schemas:
            if group.id not in group_columns_map:
                group_columns_map[group.id] = {'group': group, 'columns': []}
            group_columns_map[group.id]['columns'].append(schema.column_name)
        
        best_match = None
        best_similarity = 0.0
        
        for group_id, group_data in group_columns_map.items():
            group = group_data['group']
            group_columns = group_data['columns']
            
            # 计算相似度
            similarity = cls.calculate_column_similarity(cleaned_columns, group_columns)
            
            if similarity > best_similarity and similarity >= cls.MIN_SIMILARITY_THRESHOLD:
                best_match = group
                best_similarity = similarity
        
        if best_match:
            print(f"[系统] 找到相似分组: {best_match.group_name}, 相似度: {best_similarity:.2f}")
            return best_match, best_similarity
        
        print("[系统] 未找到匹配的分组，将创建新分组")
        return None, 0.0
    
    @classmethod
    def create_table_group(cls, columns, filename):
        """创建新的表格分组 - 并发安全版本，绝对避免重复创建"""
        
        # 使用锁确保并发安全
        with cls._creation_lock:
            # 首先清理列名
            cleaned_columns = cls.clean_column_names(columns)
            
            # 再次检查是否已存在匹配的分组（双重保险）
            existing_group, similarity = cls.find_matching_table_group(cleaned_columns)
            if existing_group and similarity >= cls.HIGH_SIMILARITY_THRESHOLD:
                print(f"[系统] 在创建分组前发现匹配分组: {existing_group.group_name}，直接使用")
                return existing_group
            
            # 生成指纹，再次检查是否存在相同指纹的分组
            new_fingerprint = cls.generate_schema_fingerprint(cleaned_columns)
            
            # 使用事务确保原子性
            with cls.database_transaction() as session:
                # 再次检查是否已存在相同指纹的分组（防止并发创建）
                fingerprint_group = TableGroup.query.filter_by(
                    schema_fingerprint=new_fingerprint,
                    column_count=len(cleaned_columns)
                ).first()
                
                if fingerprint_group:
                    print(f"[系统] 发现相同指纹的分组: {fingerprint_group.group_name}，直接使用")
                    return fingerprint_group
            
                # 使用API管理器生成智能表格名称
                group_name = cls._generate_smart_table_name(cleaned_columns, filename)

                # 确保名称唯一
                original_name = group_name
                counter = 1
                import re
                merge_table_pattern = re.compile(r'^合并表(\d+)$')
                while True:
                    existing_group = TableGroup.query.filter_by(group_name=group_name).first()
                    if not existing_group:
                        break

                    # 若是“合并表N”样式，则改为寻找下一个未使用的“合并表<number>”
                    m = merge_table_pattern.match(original_name)
                    if m:
                        try:
                            # 收集所有已存在的“合并表<number>”序号
                            all_names = [g.group_name for g in TableGroup.query.all() if g.group_name]
                            used = set()
                            for nm in all_names:
                                mm = merge_table_pattern.match(nm.strip())
                                if mm:
                                    used.add(int(mm.group(1)))
                            next_num = 1
                            while next_num in used:
                                next_num += 1
                            group_name = f"合并表{next_num}"
                            # 循环继续校验唯一性
                            continue
                        except Exception:
                            # 回退到下划线方式
                            pass

                    # 默认回退：原名加 _序号
                    counter += 1
                    group_name = f"{original_name}_{counter}"
                    
                    # 避免无限循环
                    if counter > 100:
                        group_name = f"{original_name}_{int(time.time())}"
                        break
                
                # 创建分组
                group = TableGroup(
                    group_name=group_name,
                    description=f"基于文件 {filename} 创建的表格分组",
                    schema_fingerprint=new_fingerprint,
                    column_count=len(cleaned_columns),
                    confidence_score=1.0  # 新创建的分组置信度为100%
                )
                
                session.add(group)
                session.flush()  # 获取group.id
                
                # 创建schema
                for i, col_name in enumerate(cleaned_columns):
                    schema = TableSchema(
                        column_name=col_name,
                        column_type='text',
                        column_order=i,
                        is_active=True,
                        table_group_id=group.id
                    )
                    session.add(schema)
                
                print(f"[系统] 创建新表格分组: {group_name}")
                return group
    
    @staticmethod
    def create_column_mappings(group, original_columns, target_columns, filename, similarity_score):
        """创建列名映射关系"""
        print(f"[系统] 创建列映射关系，相似度: {similarity_score:.2f}")
        
        for orig_col, target_col in zip(original_columns, target_columns):
            # 计算单列相似度
            col_similarity = SequenceMatcher(None, orig_col.lower().strip(), target_col.lower().strip()).ratio()
            
            mapping = ColumnMapping(
                table_group_id=group.id,
                original_column=orig_col,
                mapped_column=target_col,
                source_file=filename,
                similarity_score=col_similarity,
                is_confirmed=col_similarity >= UniversalExcelProcessor.HIGH_SIMILARITY_THRESHOLD
            )
            db.session.add(mapping)
        
        db.session.commit()
    
    @classmethod
    def process_excel_file_with_grouping(cls, file_path, filename):
        """处理Excel文件并进行智能分组 - 增强版本，包含文件大小检查和错误处理"""
        print(f"[系统] 开始处理Excel文件进行智能分组: {filename}")
        
        try:
            # 1. 文件大小检查
            import os
            file_size = os.path.getsize(file_path)
            max_size = 100 * 1024 * 1024  # 100MB 限制
            
            if file_size > max_size:
                return False, f"文件过大（{file_size/1024/1024:.1f}MB），最大支持100MB", 0, None
            
            print(f"[系统] 文件大小: {file_size/1024/1024:.2f}MB")
            
            # 2. 尝试读取文件（支持多种格式）
            try:
                df_raw = pd.read_excel(file_path, engine='openpyxl', header=None)
            except Exception as e1:
                try:
                    # 尝试使用xlrd引擎（适用于.xls文件）
                    df_raw = pd.read_excel(file_path, engine='xlrd', header=None)
                    print("[系统] 使用xlrd引擎读取文件")
                except Exception as e2:
                    return False, f"无法读取Excel文件: {str(e1)}", 0, None
            
            if df_raw.empty:
                return False, "文件为空", 0, None
                
            # 3. 基本数据验证
            total_rows, total_cols = df_raw.shape
            print(f"[系统] 文件维度: {total_rows} 行 x {total_cols} 列")
            
            if total_rows > 50000:
                return False, f"数据行数过多（{total_rows}行），最大支持50000行", 0, None
            
            if total_cols > 200:
                return False, f"列数过多（{total_cols}列），最大支持200列", 0, None
            
            # 检测表头
            header_row = cls.detect_header_row(df_raw)
            df = pd.read_excel(file_path, engine='openpyxl', header=header_row)
            
            # 清理数据
            df = df.dropna(how='all').dropna(axis=1, how='all')
            if df.empty:
                return False, "文件中没有有效数据", 0, None
            
            # 清理列名
            original_columns = cls.clean_column_names(df.columns)
            df.columns = original_columns
            
            print(f"[系统] 检测到的列名: {original_columns}")
            
            # 更新进度
            cls._update_progress('分组处理', 30, '正在查找匹配的表格分组...')
            
            # 强化版分组处理逻辑
            print("[系统] ==================== 开始分组处理 ====================")
            
            # 查找匹配的表格分组
            matching_group, similarity = cls.find_matching_table_group(original_columns)
            
            if matching_group is None:
                print("[系统] 未找到匹配分组，创建新分组")
                # 创建新分组（内部已有多重检查）
                group = cls.create_table_group(original_columns, filename)
                target_columns = cls.clean_column_names(original_columns)  # 确保使用清理后的列名
                print(f"[系统] ✅ 成功创建/获取分组: {group.group_name} (ID: {group.id})")
            else:
                print(f"[系统] 找到匹配分组: {matching_group.group_name} (ID: {matching_group.id}), 相似度: {similarity:.3f}")
                group = matching_group
                
                # 获取目标列结构
                target_schemas = TableSchema.query.filter_by(
                    table_group_id=group.id, 
                    is_active=True
                ).order_by(TableSchema.column_order).all()
                target_columns = [schema.column_name for schema in target_schemas]
                
                print(f"[系统] 目标列结构: {target_columns}")
                
                # 更新分组的置信度（基于历史平均相似度）
                current_confidence = group.confidence_score or 1.0
                current_file_count = len(group.data_records) + 1  # 包括即将添加的文件
                
                # 计算新的置信度：加权平均
                new_confidence = ((current_confidence * (current_file_count - 1)) + similarity) / current_file_count
                group.confidence_score = new_confidence
                
                print(f"[系统] 更新置信度: {current_confidence:.3f} -> {new_confidence:.3f}")
                
                # 如果不是完全匹配，创建映射关系
                if similarity < cls.EXACT_MATCH_THRESHOLD:
                    print(f"[系统] 相似度 {similarity:.3f} < 1.0，创建列映射关系")
                    cls.create_column_mappings(
                        group, original_columns, target_columns, filename, similarity
                    )
                else:
                    print("[系统] 完全匹配，无需列映射")
                
                print(f"[系统] ✅ 使用现有分组: {group.group_name}")
            
            print("[系统] ==================== 分组处理完成 ====================")
            
            # 验证分组状态
            if not group or not group.id:
                raise Exception("分组创建失败，group为空或无效")
                
            print(f"[系统] 最终使用分组: {group.group_name} (ID: {group.id})")
            
            # 更新进度
            cls._update_progress('数据导入', 60, '正在导入数据...')
            
            # 导入数据 - 分批处理优化内存使用
            imported_count = 0
            batch_size = 1000  # 每批处理1000条记录
            batch_data = []
            
            total_rows = len(df)
            print(f"[系统] 开始分批导入数据，总行数: {total_rows}，批次大小: {batch_size}")
            
            for index, row in df.iterrows():
                if row.isna().all():
                    continue
                
                # 创建数据字典，使用目标列名
                row_dict = {}
                for orig_col, target_col in zip(original_columns, target_columns):
                    value = row[orig_col] if orig_col in row.index else ""
                    if pd.notna(value):
                        row_dict[target_col] = str(value)
                    else:
                        row_dict[target_col] = ""
                
                # 跳过空行
                if not any(v.strip() for v in row_dict.values() if v):
                    continue
                
                # 创建数据记录
                table_data = TableData()
                table_data.source_file = filename
                table_data.table_group_id = group.id
                table_data.set_data(row_dict)
                
                batch_data.append(table_data)
                
                # 达到批次大小或最后一行时提交
                if len(batch_data) >= batch_size or index == total_rows - 1:
                    try:
                        db.session.add_all(batch_data)
                        db.session.commit()
                        imported_count += len(batch_data)
                        print(f"[系统] 已导入 {imported_count} 条数据")
                        batch_data.clear()  # 清空批次数据
                    except Exception as e:
                        db.session.rollback()
                        print(f"[错误] 批次导入时出错: {str(e)}")
                        # 尝试逐条导入这个批次
                        for data in batch_data:
                            try:
                                db.session.add(data)
                                db.session.commit()
                                imported_count += 1
                            except Exception as e2:
                                db.session.rollback()
                                print(f"[错误] 导入第 {imported_count+1} 行数据时出错: {str(e2)}")
                        batch_data.clear()
            
            print(f"[系统] 成功导入 {imported_count} 条数据到分组: {group.group_name}")
            
            # 记录上传历史
            history = UploadHistory(
                filename=filename,
                rows_imported=imported_count,
                status='success'
            )
            history.set_columns(original_columns)
            db.session.add(history)
            db.session.commit()
            
            return True, f"导入成功，分组: {group.group_name}", imported_count, group.id
            
        except Exception as e:
            db.session.rollback()
            print(f"[错误] 处理文件时出错: {str(e)}")
            return False, str(e), 0, None
    
    @classmethod
    def _generate_smart_table_name(cls, columns, filename):
        """使用配置的API提供商生成智能表格名称"""
        print(f"[智能命名] 开始为表格生成智能名称，列数: {len(columns)}")
        
        try:
            # 获取API配置（从某种配置存储中读取）
            api_config = cls._get_api_config()
            
            # 创建API管理器
            api_manager = APIManager(api_config)
            
            # 调用API生成名称
            success, table_name, message = api_manager.generate_table_name(columns, filename=filename)
            
            if success and table_name:
                print(f"[智能命名] API生成名称成功: {table_name}")
                return table_name
            else:
                print(f"[智能命名] API生成失败: {message}")
                return cls._generate_fallback_name(columns, filename)
                
        except Exception as e:
            print(f"[智能命名] API调用异常: {str(e)}")
            return cls._generate_fallback_name(columns, filename)
    
    @classmethod
    def _get_api_config(cls):
        """获取API配置"""
        return get_api_config()
    
    @staticmethod
    def _generate_fallback_name(columns, filename):
        """生成后备表格名称

        优先根据列/文件关键词推断，否则采用“合并表<number>”连续编号，
        避免出现下划线序号形式。
        """
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
        column_text = ''.join(columns).lower()
        
        for table_type, keywords in name_keywords.items():
            if sum(1 for keyword in keywords if keyword in column_text) >= 2:
                print(f"[智能命名] 根据列名推断为: {table_type}")
                return table_type
        
        # 如果从文件名能推断出类型
        if filename:
            filename_lower = filename.lower()
            for table_type, keywords in name_keywords.items():
                if any(keyword in filename_lower for keyword in keywords):
                    print(f"[智能命名] 根据文件名推断为: {table_type}")
                    return table_type
        
        # 默认使用“合并表<number>”编号
        try:
            from models.database import TableGroup
            import re
            existing_names = [g.group_name for g in TableGroup.query.all() if g.group_name]
            pattern = re.compile(r'^合并表(\d+)$')
            used_numbers = set()
            for name in existing_names:
                m = pattern.match(name.strip())
                if m:
                    try:
                        used_numbers.add(int(m.group(1)))
                    except ValueError:
                        continue
            next_num = 1
            while next_num in used_numbers:
                next_num += 1
            default_name = f"合并表{next_num}"
        except Exception:
            # 兜底到时间戳命名，确保不失败
            current_time = datetime.now()
            default_name = f"合并表_{current_time.strftime('%m%d_%H%M')}"

        print(f"[智能命名] 使用默认命名: {default_name}")
        return default_name
    
    @staticmethod
    def _number_to_chinese(num):
        """将数字转换为中文数字"""
        chinese_numbers = ['', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十']
        
        if num <= 0:
            return '一'
        elif num <= 10:
            return chinese_numbers[num]
        elif num <= 19:
            return '十' + chinese_numbers[num - 10]
        elif num <= 99:
            tens = num // 10
            ones = num % 10
            if ones == 0:
                return chinese_numbers[tens] + '十'
            else:
                return chinese_numbers[tens] + '十' + chinese_numbers[ones]
        else:
            # 对于更大的数字，简化处理
            return str(num)
    
    @staticmethod
    def cleanup_duplicate_groups():
        """清理重复的表格分组"""
        print("[系统] 开始清理重复分组...")
        
        try:
            # 获取所有分组，按指纹分组
            all_groups = TableGroup.query.all()
            fingerprint_groups = {}
            
            for group in all_groups:
                fp = group.schema_fingerprint
                if fp not in fingerprint_groups:
                    fingerprint_groups[fp] = []
                fingerprint_groups[fp].append(group)
            
            cleaned_count = 0
            for fingerprint, groups in fingerprint_groups.items():
                if len(groups) > 1:
                    print(f"[系统] 发现指纹 {fingerprint[:10]}... 有 {len(groups)} 个重复分组")
                    
                    # 选择记录数最多的作为主分组
                    main_group = max(groups, key=lambda g: len(g.data_records))
                    
                    for group in groups:
                        if group.id != main_group.id:
                            print(f"[系统] 合并分组 {group.group_name} (ID: {group.id}) 到 {main_group.group_name} (ID: {main_group.id})")
                            
                            # 迁移数据
                            data_records = TableData.query.filter_by(table_group_id=group.id).all()
                            for record in data_records:
                                record.table_group_id = main_group.id
                            
                            # 删除重复分组
                            db.session.delete(group)
                            cleaned_count += 1
            
            if cleaned_count > 0:
                db.session.commit()
                print(f"[系统] ✅ 清理完成，合并了 {cleaned_count} 个重复分组")
            else:
                print("[系统] ✅ 没有发现重复分组")
                
            return cleaned_count
            
        except Exception as e:
            db.session.rollback()
            print(f"[错误] 清理重复分组时出错: {str(e)}")
            return 0
    
    @classmethod
    def validate_system_health(cls):
        """验证系统健康状态和功能完整性"""
        print("🔍 [系统] 开始系统健康检查...")
        
        issues = []
        
        try:
            # 1. 检查数据库连接
            total_groups = TableGroup.query.count()
            total_data = TableData.query.count()
            print(f"✅ [数据库] 连接正常，共 {total_groups} 个分组，{total_data} 条数据")
            
            # 2. 检查重复分组
            duplicate_count = cls.cleanup_duplicate_groups()
            if duplicate_count > 0:
                issues.append(f"发现并清理了 {duplicate_count} 个重复分组")
            
            # 3. 检查数据完整性
            orphaned_data = db.session.query(TableData).outerjoin(
                TableGroup, TableData.table_group_id == TableGroup.id
            ).filter(TableGroup.id.is_(None)).count()
            
            if orphaned_data > 0:
                issues.append(f"发现 {orphaned_data} 条孤儿数据（无对应分组）")
            
            # 4. 检查列名别名功能
            test_columns = ['编号', 'ID', '名称', 'name']
            normalized = [UniversalExcelProcessor._normalize_column_name(col) for col in test_columns]
            if normalized[0] == normalized[1] and normalized[2] == normalized[3]:
                print("✅ [智能识别] 列名别名功能正常")
            else:
                issues.append("列名别名识别功能异常")
            
            # 5. 检查相似度计算
            similarity = UniversalExcelProcessor.calculate_column_similarity(
                ['项目编号', '项目名称', '申请部门'],
                ['编号', '名称', '部门']
            )
            if similarity > 0.5:
                print(f"✅ [相似度计算] 功能正常，测试相似度: {similarity:.3f}")
            else:
                issues.append(f"相似度计算可能异常，测试结果: {similarity:.3f}")
            
            # 6. 性能检查
            import time
            start_time = time.time()
            
            # 模拟计算较大列表的相似度
            large_cols1 = [f"列{i}" for i in range(50)]
            large_cols2 = [f"col{i}" for i in range(50)]
            
            similarity = UniversalExcelProcessor.calculate_column_similarity(large_cols1, large_cols2)
            
            elapsed = time.time() - start_time
            if elapsed < 1.0:
                print(f"✅ [性能] 50列相似度计算耗时 {elapsed:.3f}s")
            else:
                issues.append(f"性能问题：50列相似度计算耗时 {elapsed:.3f}s")
            
            # 总结
            if not issues:
                print("🎉 [系统] 健康检查通过，所有功能正常！")
                return True, "系统状态良好"
            else:
                warning_msg = "存在以下问题：" + "；".join(issues)
                print(f"⚠️ [系统] {warning_msg}")
                return False, warning_msg
                
        except Exception as e:
            error_msg = f"健康检查失败: {str(e)}"
            print(f"❌ [系统] {error_msg}")
            return False, error_msg
