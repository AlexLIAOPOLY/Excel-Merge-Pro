try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False
    pd = None
import os
from datetime import datetime
from difflib import SequenceMatcher
import hashlib
from models.database import db, TableData, TableSchema, UploadHistory, TableGroup, ColumnMapping

class UniversalExcelProcessor:
    """通用Excel表格处理器，支持任意格式的表格合并和智能分组"""
    
    # 相似度阈值配置
    EXACT_MATCH_THRESHOLD = 1.0      # 完全匹配
    HIGH_SIMILARITY_THRESHOLD = 0.8  # 高相似度
    MIN_SIMILARITY_THRESHOLD = 0.6   # 最低相似度
    
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
        """清理列名"""
        cleaned = []
        for col in columns:
            # Handle case when pandas is not available
            is_na = pd.isna(col) if HAS_PANDAS else (col is None or col == '')
            if is_na:
                cleaned.append(f"未命名列_{len(cleaned)+1}")
            else:
                col_str = str(col).strip()
                if not col_str:
                    col_str = f"未命名列_{len(cleaned)+1}"
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
        
        if not HAS_PANDAS:
            return False, "pandas未安装，无法处理Excel文件", 0
        
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
                    # Handle case when pandas is not available
                    is_not_na = pd.notna(value) if HAS_PANDAS else (value is not None and value != '')
                    if is_not_na:
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
    
    @staticmethod
    def calculate_column_similarity(columns1, columns2):
        """计算两个列名列表的相似度"""
        if len(columns1) != len(columns2):
            # 长度不同时，基于最长公共子序列计算相似度
            matcher = SequenceMatcher(None, columns1, columns2)
            return matcher.ratio()
        
        # 长度相同时，逐列计算相似度
        similarities = []
        for col1, col2 in zip(columns1, columns2):
            # 计算单个列名的相似度
            col_similarity = SequenceMatcher(None, col1.lower().strip(), col2.lower().strip()).ratio()
            similarities.append(col_similarity)
        
        # 返回平均相似度
        return sum(similarities) / len(similarities) if similarities else 0.0
    
    @staticmethod
    def generate_schema_fingerprint(columns):
        """生成表头指纹用于快速匹配"""
        # 标准化列名：转小写、去空格、排序
        normalized_columns = sorted([col.lower().strip() for col in columns])
        fingerprint_text = '|'.join(normalized_columns)
        return hashlib.md5(fingerprint_text.encode()).hexdigest()
    
    @staticmethod
    def find_matching_table_group(columns):
        """查找匹配的表格分组"""
        print(f"[系统] 查找匹配的表格分组，列数: {len(columns)}")
        
        # 生成当前表头的指纹
        current_fingerprint = UniversalExcelProcessor.generate_schema_fingerprint(columns)
        
        # 首先查找完全匹配的分组
        exact_match = TableGroup.query.filter_by(
            schema_fingerprint=current_fingerprint,
            column_count=len(columns)
        ).first()
        
        if exact_match:
            print(f"[系统] 找到完全匹配的分组: {exact_match.group_name}")
            return exact_match, UniversalExcelProcessor.EXACT_MATCH_THRESHOLD
        
        # 查找相似的分组
        all_groups = TableGroup.query.filter_by(column_count=len(columns)).all()
        best_match = None
        best_similarity = 0.0
        
        for group in all_groups:
            # 获取分组的列结构
            group_schemas = TableSchema.query.filter_by(
                table_group_id=group.id, 
                is_active=True
            ).order_by(TableSchema.column_order).all()
            
            group_columns = [schema.column_name for schema in group_schemas]
            
            # 计算相似度
            similarity = UniversalExcelProcessor.calculate_column_similarity(columns, group_columns)
            
            if similarity > best_similarity and similarity >= UniversalExcelProcessor.MIN_SIMILARITY_THRESHOLD:
                best_match = group
                best_similarity = similarity
        
        if best_match:
            print(f"[系统] 找到相似分组: {best_match.group_name}, 相似度: {best_similarity:.2f}")
            return best_match, best_similarity
        
        print("[系统] 未找到匹配的分组，将创建新分组")
        return None, 0.0
    
    @staticmethod
    def create_table_group(columns, filename):
        """创建新的表格分组"""
        group_name = f"表格组_{len(TableGroup.query.all()) + 1}_{filename.split('.')[0] if '.' in filename else filename}"
        
        # 创建分组
        group = TableGroup(
            group_name=group_name,
            description=f"基于文件 {filename} 创建的表格分组",
            schema_fingerprint=UniversalExcelProcessor.generate_schema_fingerprint(columns),
            column_count=len(columns)
        )
        
        db.session.add(group)
        db.session.flush()  # 获取group.id
        
        # 创建schema
        for i, col_name in enumerate(columns):
            schema = TableSchema(
                column_name=col_name,
                column_type='text',
                column_order=i,
                is_active=True,
                table_group_id=group.id
            )
            db.session.add(schema)
        
        db.session.commit()
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
    
    @staticmethod
    def process_excel_file_with_grouping(file_path, filename):
        """处理Excel文件并进行智能分组"""
        print(f"[系统] 开始处理Excel文件进行智能分组: {filename}")
        
        if not HAS_PANDAS:
            return False, "pandas未安装，无法处理Excel文件", 0, None
        
        try:
            # 读取文件
            df_raw = pd.read_excel(file_path, engine='openpyxl', header=None)
            
            if df_raw.empty:
                return False, "文件为空", 0, None
            
            # 检测表头
            header_row = UniversalExcelProcessor.detect_header_row(df_raw)
            df = pd.read_excel(file_path, engine='openpyxl', header=header_row)
            
            # 清理数据
            df = df.dropna(how='all').dropna(axis=1, how='all')
            if df.empty:
                return False, "文件中没有有效数据", 0, None
            
            # 清理列名
            original_columns = UniversalExcelProcessor.clean_column_names(df.columns)
            df.columns = original_columns
            
            print(f"[系统] 检测到的列名: {original_columns}")
            
            # 查找匹配的表格分组
            matching_group, similarity = UniversalExcelProcessor.find_matching_table_group(original_columns)
            
            if matching_group is None:
                # 创建新分组
                group = UniversalExcelProcessor.create_table_group(original_columns, filename)
                target_columns = original_columns
                print(f"[系统] 创建新分组: {group.group_name}")
            else:
                group = matching_group
                # 获取目标列结构
                target_schemas = TableSchema.query.filter_by(
                    table_group_id=group.id, 
                    is_active=True
                ).order_by(TableSchema.column_order).all()
                target_columns = [schema.column_name for schema in target_schemas]
                
                # 如果不是完全匹配，创建映射关系
                if similarity < UniversalExcelProcessor.EXACT_MATCH_THRESHOLD:
                    UniversalExcelProcessor.create_column_mappings(
                        group, original_columns, target_columns, filename, similarity
                    )
                    print(f"[系统] 映射到现有分组: {group.group_name}")
            
            # 导入数据
            imported_count = 0
            for index, row in df.iterrows():
                if row.isna().all():
                    continue
                
                # 创建数据字典，使用目标列名
                row_dict = {}
                for orig_col, target_col in zip(original_columns, target_columns):
                    value = row[orig_col] if orig_col in row.index else ""
                    # Handle case when pandas is not available
                    is_not_na = pd.notna(value) if HAS_PANDAS else (value is not None and value != '')
                    if is_not_na:
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
                
                try:
                    db.session.add(table_data)
                    imported_count += 1
                except Exception as e:
                    print(f"[错误] 导入第 {index+1} 行数据时出错: {str(e)}")
                    continue
            
            db.session.commit()
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