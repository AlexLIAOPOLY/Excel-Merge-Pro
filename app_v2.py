from flask import Flask, render_template, request, jsonify, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False
    pd = None
    print("[警告] pandas未安装，高级数据处理功能将不可用")
from datetime import datetime
from models.database import db, TableData, TableSchema, UploadHistory
from models.excel_processor import UniversalExcelProcessor
from config import config

app = Flask(__name__)

# 根据环境变量选择配置
config_name = os.environ.get('FLASK_ENV', 'default')
app.config.from_object(config[config_name])

CORS(app)
db.init_app(app)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    print("[系统] 收到文件上传请求")
    
    if 'files[]' not in request.files:
        print("[错误] 没有找到上传的文件")
        return jsonify({'success': False, 'message': '没有选择文件'})
    
    files = request.files.getlist('files[]')
    results = []
    
    for file in files:
        if file.filename == '':
            continue
            
        if not allowed_file(file.filename):
            results.append({
                'filename': file.filename,
                'success': False,
                'message': '不支持的文件格式，请使用.xlsx或.xls文件'
            })
            continue
        
        original_filename = file.filename  # 保存原始文件名用于数据库
        secure_filename_value = secure_filename(file.filename)  # 安全文件名用于文件系统
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename_value)
        
        try:
            file.save(file_path)
            print(f"[系统] 文件保存成功: {file_path}")
            
            success, message, count, group_id = UniversalExcelProcessor.process_excel_file_with_grouping(file_path, original_filename)
            
            results.append({
                'filename': original_filename,
                'success': success,
                'message': message,
                'count': count,
                'group_id': group_id
            })
            
            # 清理临时文件
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"[系统] 临时文件已删除: {file_path}")
                
        except Exception as e:
            print(f"[错误] 处理文件 {original_filename} 时出错: {str(e)}")
            results.append({
                'filename': original_filename,
                'success': False,
                'message': f'处理失败: {str(e)}',
                'count': 0,
                'group_id': None
            })
    
    return jsonify({'results': results})

@app.route('/data')
def get_data():
    print("[系统] 获取表格数据请求")
    try:
        from models.database import TableGroup
        
        # 获取所有分组的合并数据
        all_data = []
        all_schema = set()
        
        groups = TableGroup.query.order_by(TableGroup.updated_at.desc()).all()
        if groups:
            # 如果有分组，获取最新更新的分组数据
            latest_group = groups[0]  # 最近更新的分组
            data_records = TableData.query.filter_by(table_group_id=latest_group.id).order_by(TableData.created_at.asc()).all()
            all_data = [record.to_dict() for record in data_records]
            
            # 获取分组的表结构
            schemas = TableSchema.query.filter_by(table_group_id=latest_group.id, is_active=True).order_by(TableSchema.column_order).all()
            schema = [s.column_name for s in schemas]
        else:
            # 兼容旧数据：如果没有分组，使用原来的方式
            data = UniversalExcelProcessor.get_all_data()
            schema = UniversalExcelProcessor.get_current_schema()
            all_data = data
        
        # 计算统计信息
        stats = UniversalExcelProcessor.get_data_stats()
        
        print(f"[系统] 返回 {len(all_data)} 条记录，{len(schema)} 个列")
        return jsonify({
            'success': True,
            'data': all_data,
            'stats': stats,
            'schema': schema
        })
    except Exception as e:
        print(f"[错误] 获取数据时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/schema')
def get_schema():
    """获取当前表格结构"""
    try:
        schema = UniversalExcelProcessor.get_current_schema()
        return jsonify({
            'success': True,
            'schema': schema
        })
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/update', methods=['POST'])
def update_record():
    try:
        data = request.get_json()
        record_id = data.get('id')
        update_data = {k: v for k, v in data.items() if k not in ['id', 'source_file', 'created_at', 'updated_at']}
        
        print(f"[系统] 更新记录 ID: {record_id}")
        
        success = UniversalExcelProcessor.update_record(record_id, update_data)
        
        if success:
            print(f"[系统] 记录更新成功")
            return jsonify({'success': True, 'message': '更新成功'})
        else:
            print(f"[错误] 记录更新失败")
            return jsonify({'success': False, 'message': '记录不存在'})
            
    except Exception as e:
        print(f"[错误] 更新记录时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/delete/<int:record_id>', methods=['DELETE'])
def delete_record(record_id):
    try:
        print(f"[系统] 删除记录 ID: {record_id}")
        
        success = UniversalExcelProcessor.delete_record(record_id)
        
        if success:
            print(f"[系统] 记录删除成功")
            return jsonify({'success': True, 'message': '删除成功'})
        else:
            print(f"[错误] 记录删除失败")
            return jsonify({'success': False, 'message': '记录不存在'})
            
    except Exception as e:
        print(f"[错误] 删除记录时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/export')
def export_data():
    try:
        print("[系统] 开始导出数据")
        
        data = UniversalExcelProcessor.get_all_data()
        
        if not data:
            return jsonify({'success': False, 'message': '没有数据可导出'})
        
        # 获取当前表格结构
        schema = UniversalExcelProcessor.get_current_schema()
        
        if not schema:
            return jsonify({'success': False, 'message': '没有有效的表格结构'})
        
        # 系统字段列表（不应该出现在导出中）
        system_fields = {'id', 'source_file', 'created_at', 'updated_at'}
        
        # 过滤掉系统字段，只保留业务数据列
        business_columns = [col for col in schema if col not in system_fields and col.strip()]
        
        if not business_columns:
            return jsonify({'success': False, 'message': '没有可导出的业务数据列'})
        
        print(f"[系统] 业务数据列: {', '.join(business_columns)}")
        
        # 创建DataFrame，只包含业务数据列
        rows = []
        valid_data_count = 0
        
        for item in data:
            row = {}
            has_data = False
            
            # 只包含业务数据列，确保不包含任何系统字段
            for col in business_columns:
                value = item.get(col, '')
                # 清理数据：去除首尾空格，转换None为空字符串
                if value is not None:
                    value = str(value).strip()
                else:
                    value = ''
                row[col] = value
                
                # 检查是否有有效数据（非空且不是手动添加的空行）
                if value and value != '':
                    has_data = True
            
            # 只添加有实际数据的行，过滤掉完全空的行
            if has_data:
                rows.append(row)
                valid_data_count += 1
        
        if not rows:
            return jsonify({'success': False, 'message': '没有有效数据可导出'})
        
        # 创建DataFrame，确保列顺序与schema一致
        if HAS_PANDAS:
            df = pd.DataFrame(rows, columns=schema)
        else:
            # 简单的数据结构替代pandas
            df = {"rows": rows, "columns": schema}
        
        # 清理DataFrame：移除完全空的列
        df = df.dropna(axis=1, how='all')
        
        # 生成导出文件
        export_filename = f'数据表格_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        export_path = os.path.join('static/uploads', export_filename)
        
        # 确保导出目录存在
        os.makedirs(os.path.dirname(export_path), exist_ok=True)
        
        # 使用openpyxl写入Excel，优化格式
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, PatternFill
        from openpyxl.utils.dataframe import dataframe_to_rows
        
        wb = Workbook()
        ws = wb.active
        ws.title = "数据表格"
        
        # 写入数据
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # 设置表头样式
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center")
        
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
        
        # 自动调整列宽
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)  # 最大宽度50
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # 保存文件
        wb.save(export_path)
        
        print(f"[系统] 数据导出成功: {export_filename}, 共导出 {valid_data_count} 条有效记录")
        
        return send_file(export_path, as_attachment=True, download_name=export_filename)
        
    except Exception as e:
        print(f"[错误] 导出数据时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})

def apply_enhanced_excel_styling(ws, num_columns, num_rows):
    """应用增强的Excel样式，使其与网页格式保持一致"""
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, NamedStyle
    from openpyxl.utils import get_column_letter
    
    # 定义专业的样式（突出的标题栏高亮效果）
    header_font = Font(bold=True, color="ffffff", name="Microsoft YaHei", size=13)  # 白色字体更醒目
    header_fill = PatternFill(start_color="4f46e5", end_color="3730a3", fill_type="solid")  # 深蓝色渐变背景
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    data_font = Font(color="374151", name="Microsoft YaHei", size=11)
    data_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    # 数值类型数据的对齐方式
    number_alignment = Alignment(horizontal="right", vertical="center", wrap_text=False)
    
    # 精美的边框样式
    thin_border = Border(
        left=Side(border_style="thin", color="d1d5db"),
        right=Side(border_style="thin", color="d1d5db"),
        top=Side(border_style="thin", color="d1d5db"),
        bottom=Side(border_style="thin", color="d1d5db")
    )
    
    header_border = Border(
        left=Side(border_style="medium", color="1e293b"),
        right=Side(border_style="medium", color="1e293b"),
        top=Side(border_style="thick", color="1e293b"),
        bottom=Side(border_style="thick", color="1e293b")
    )
    
    # 应用表头样式
    for col in range(1, num_columns + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = header_border
    
    # 应用数据行样式
    for row in range(2, num_rows + 2):  # +2 because we have header + data rows
        for col in range(1, num_columns + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = data_font
            
            # 智能判断数据类型并应用对应的对齐方式
            cell_value = str(cell.value) if cell.value else ""
            if cell_value and (cell_value.replace('.','').replace('-','').replace('+','').isdigit() or 
                             any(char in cell_value for char in ['¥', '$', '元', '%'])):
                cell.alignment = number_alignment
            else:
                cell.alignment = data_alignment
            
            cell.border = thin_border
            
            # 更精致的交替行背景色
            if row % 2 == 0:
                cell.fill = PatternFill(start_color="f8fafc", end_color="f8fafc", fill_type="solid")
            else:
                cell.fill = PatternFill(start_color="ffffff", end_color="ffffff", fill_type="solid")
    
    # 智能调整列宽（优化算法）
    for col in range(1, num_columns + 1):
        column_letter = get_column_letter(col)
        max_length = 0
        
        # 计算列的最大内容长度，优化中文字符处理
        for row in range(1, num_rows + 2):
            cell = ws.cell(row=row, column=col)
            if cell.value:
                cell_str = str(cell.value)
                # 更精确的中文字符宽度计算
                ascii_count = sum(1 for c in cell_str if ord(c) < 128)
                chinese_count = sum(1 for c in cell_str if ord(c) >= 128)
                adjusted_length = ascii_count + chinese_count * 1.8  # 中文字符宽度系数
                max_length = max(max_length, adjusted_length)
        
        # 设置列宽，考虑表头加粗效果和内容类型
        header_cell = ws.cell(row=1, column=col)
        header_length = len(str(header_cell.value)) * 1.2 if header_cell.value else 0  # 加粗字体稍宽
        
        final_width = max(10, min(max(max_length + 4, header_length + 2), 60))  # 改进的宽度计算
        ws.column_dimensions[column_letter].width = final_width
    
    # 设置更合适的行高（突出标题栏）
    ws.row_dimensions[1].height = 45  # 表头行高（更高更醒目）
    for row in range(2, num_rows + 2):
        ws.row_dimensions[row].height = 28  # 数据行高（稍高一些）
    
    # 冻结首行（便于浏览大量数据）
    ws.freeze_panes = "A2"
    
    # 优化打印和显示设置
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = False
    ws.page_setup.orientation = 'landscape'  # 横向打印
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.75
    ws.page_margins.bottom = 0.75

@app.route('/export-all-groups')
def export_all_groups():
    """导出所有表格分组，每个分组一个工作表"""
    try:
        print("[系统] 开始导出所有表格分组")
        
        from models.database import TableGroup
        
        # 获取所有表格分组
        groups = TableGroup.query.order_by(TableGroup.created_at.asc()).all()
        
        if not groups:
            return jsonify({'success': False, 'message': '没有表格分组可导出'})
        
        # 创建工作簿
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        from openpyxl.utils.dataframe import dataframe_to_rows
        
        wb = Workbook()
        # 删除默认工作表
        wb.remove(wb.active)
        
        total_exported_records = 0
        exported_groups = 0
        
        # 系统字段列表（不应该出现在导出中）
        system_fields = {'id', 'source_file', 'created_at', 'updated_at', 'table_group_id'}
        
        for group in groups:
            # 获取该分组的数据
            data_records = TableData.query.filter_by(table_group_id=group.id).order_by(TableData.created_at.asc()).all()
            
            if not data_records:
                continue
                
            # 获取该分组的表格结构
            schemas = TableSchema.query.filter_by(table_group_id=group.id, is_active=True).order_by(TableSchema.column_order).all()
            schema_columns = [s.column_name for s in schemas]
            
            # 过滤掉系统字段
            business_columns = [col for col in schema_columns if col not in system_fields and col.strip()]
            
            if not business_columns:
                continue
            
            # 准备数据
            rows = []
            for record in data_records:
                row = {}
                has_data = False
                
                for col in business_columns:
                    value = getattr(record, col, '') if hasattr(record, col) else record.to_dict().get(col, '')
                    
                    # 清理数据
                    if value is not None:
                        value = str(value).strip()
                    else:
                        value = ''
                    row[col] = value
                    
                    # 检查是否有有效数据
                    if value and value != '':
                        has_data = True
                
                if has_data:
                    rows.append(row)
            
            if not rows:
                continue
            
            # 创建DataFrame
            if HAS_PANDAS:
                df = pd.DataFrame(rows, columns=business_columns)
            else:
                # 简单的数据结构替代pandas
                df = {"rows": rows, "columns": business_columns}
            
            # 创建工作表
            ws_name = group.group_name.replace('表格组_', '').replace('_', '-')[:31]  # Excel工作表名称长度限制
            ws = wb.create_sheet(title=ws_name)
            
            # 写入数据
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)
            
            # 增强的样式设置
            apply_enhanced_excel_styling(ws, len(business_columns), len(rows))
            
            total_exported_records += len(rows)
            exported_groups += 1
            print(f"[系统] 导出分组 '{group.group_name}': {len(rows)} 条记录")
        
        if exported_groups == 0:
            return jsonify({'success': False, 'message': '没有有效数据可导出'})
        
        # 生成有规律的导出文件名
        current_time = datetime.now()
        date_str = current_time.strftime("%Y%m%d")
        time_str = current_time.strftime("%H%M%S")
        
        # 构建更有意义的文件名：分组数量+记录数量+时间戳
        export_filename = f'表格数据汇总_{exported_groups}组_{total_exported_records}条_{date_str}_{time_str}.xlsx'
        export_path = os.path.join('static/uploads', export_filename)
        
        # 确保导出目录存在
        os.makedirs(os.path.dirname(export_path), exist_ok=True)
        
        # 保存文件
        wb.save(export_path)
        
        print(f"[系统] 所有表格分组导出成功: {export_filename}, 共导出 {exported_groups} 个分组, {total_exported_records} 条记录")
        
        return send_file(export_path, as_attachment=True, download_name=export_filename)
        
    except Exception as e:
        print(f"[错误] 导出所有表格分组时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})

@app.route('/history')
def get_upload_history():
    try:
        print("[系统] 获取上传历史")
        history = UploadHistory.query.order_by(UploadHistory.upload_time.desc()).limit(20).all()
        return jsonify({
            'success': True,
            'history': [h.to_dict() for h in history]
        })
    except Exception as e:
        print(f"[错误] 获取上传历史时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/uploaded-files')
def get_uploaded_files():
    """获取所有上传的文件列表"""
    try:
        print("[系统] 获取已上传文件列表")
        
        # 获取成功上传的文件历史
        files = UploadHistory.query.filter_by(status='success').order_by(UploadHistory.upload_time.desc()).all()
        
        # 为每个文件添加其数据统计
        files_with_stats = []
        for file_record in files:
            file_data = file_record.to_dict()
            
            # 统计该文件的数据记录数
            data_count = TableData.query.filter_by(source_file=file_record.filename).count()
            file_data['current_records'] = data_count
            
            # 检查文件是否还有关联的数据（可能被部分删除了）
            file_data['has_data'] = data_count > 0
            
            files_with_stats.append(file_data)
        
        print(f"[系统] 返回 {len(files_with_stats)} 个文件记录")
        return jsonify({
            'success': True,
            'files': files_with_stats,
            'total': len(files_with_stats)
        })
    except Exception as e:
        print(f"[错误] 获取文件列表时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/uploaded-files/<filename>/data')
def get_file_data(filename):
    """获取特定文件的数据"""
    try:
        print(f"[系统] 获取文件 {filename} 的数据")
        
        # 获取该文件的所有数据记录
        data_records = TableData.query.filter_by(source_file=filename).order_by(TableData.created_at.asc()).all()
        
        if not data_records:
            return jsonify({'success': False, 'message': '文件数据不存在或已被删除'})
        
        # 获取数据
        data = [record.to_dict() for record in data_records]
        
        # 从第一条记录获取schema（所有记录应该有相同的结构）
        if data:
            first_record = data[0]
            # 排除系统字段，获取业务字段作为schema
            system_fields = {'id', 'source_file', 'created_at', 'updated_at', 'table_group_id'}
            schema = [key for key in first_record.keys() if key not in system_fields]
        else:
            schema = []
        
        # 统计信息
        stats = {
            'total_records': len(data),
            'source_file': filename,
            'total_columns': len(schema),
            'upload_time': data_records[0].created_at.strftime('%Y-%m-%d %H:%M:%S') if data_records else None
        }
        
        print(f"[系统] 返回文件 {filename} 的 {len(data)} 条记录")
        return jsonify({
            'success': True,
            'data': data,
            'schema': schema,
            'stats': stats
        })
    except Exception as e:
        print(f"[错误] 获取文件数据时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/uploaded-files/<filename>/export')
def export_file_data(filename):
    """导出特定文件的数据"""
    try:
        print(f"[系统] 开始导出文件 {filename}")
        
        # 获取该文件的所有数据记录
        data_records = TableData.query.filter_by(source_file=filename).order_by(TableData.created_at.asc()).all()
        
        if not data_records:
            return jsonify({'success': False, 'message': '文件数据不存在或已被删除'})
        
        # 准备数据
        data = [record.to_dict() for record in data_records]
        
        if not data:
            return jsonify({'success': False, 'message': '没有数据可导出'})
        
        # 系统字段列表
        system_fields = {'id', 'source_file', 'created_at', 'updated_at', 'table_group_id'}
        
        # 获取业务数据列
        first_record = data[0]
        business_columns = [col for col in first_record.keys() if col not in system_fields and col.strip()]
        
        if not business_columns:
            return jsonify({'success': False, 'message': '没有可导出的数据列'})
        
        # 准备导出数据
        rows = []
        for item in data:
            row = {}
            has_data = False
            
            for col in business_columns:
                value = item.get(col, '')
                if value is not None:
                    value = str(value).strip()
                else:
                    value = ''
                row[col] = value
                
                if value and value != '':
                    has_data = True
            
            if has_data:
                rows.append(row)
        
        if not rows:
            return jsonify({'success': False, 'message': '没有有效数据'})
        
        # 创建DataFrame
        if HAS_PANDAS:
            df = pd.DataFrame(rows, columns=business_columns)
        else:
            df = {"rows": rows, "columns": business_columns}
        
        # 生成导出文件名
        safe_filename = filename.replace('.xlsx', '').replace('.xls', '').replace(' ', '_')
        current_time = datetime.now()
        date_str = current_time.strftime("%Y%m%d_%H%M%S")
        
        export_filename = f'{safe_filename}_{len(rows)}条_{date_str}.xlsx'
        export_path = os.path.join('static/uploads', export_filename)
        
        # 确保导出目录存在
        os.makedirs(os.path.dirname(export_path), exist_ok=True)
        
        # 使用openpyxl写入Excel
        from openpyxl import Workbook
        from openpyxl.utils.dataframe import dataframe_to_rows
        
        wb = Workbook()
        ws = wb.active
        ws.title = safe_filename[:31]
        
        # 写入数据
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # 应用样式
        apply_enhanced_excel_styling(ws, len(business_columns), len(rows))
        
        # 保存文件
        wb.save(export_path)
        
        print(f"[系统] 文件导出成功: {export_filename}, 共导出 {len(rows)} 条记录")
        
        return send_file(export_path, as_attachment=True, download_name=export_filename)
        
    except Exception as e:
        print(f"[错误] 导出文件数据时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})

@app.route('/uploaded-files/<filename>/delete', methods=['DELETE'])
def delete_file_data(filename):
    """删除特定文件的所有数据"""
    try:
        print(f"[系统] 开始删除文件 {filename} 的数据")
        
        # 删除该文件的所有数据记录
        deleted_count = TableData.query.filter_by(source_file=filename).delete()
        
        # 删除上传历史记录
        UploadHistory.query.filter_by(filename=filename).delete()
        
        db.session.commit()
        
        print(f"[系统] 文件 {filename} 删除成功，删除了 {deleted_count} 条数据记录")
        return jsonify({
            'success': True,
            'message': f'文件删除成功，共删除 {deleted_count} 条记录'
        })
        
    except Exception as e:
        db.session.rollback()
        print(f"[错误] 删除文件数据时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

# 旧的清空数据接口已移除，使用 /clear-all 代替

@app.route('/column/add', methods=['POST'])
def add_column():
    try:
        data = request.get_json()
        column_name = data.get('column_name', '').strip()
        insert_position = data.get('insert_position')  # 新增位置参数
        
        if not column_name:
            return jsonify({'success': False, 'message': '列名不能为空'})
        
        print(f"[系统] 添加新列: {column_name}, 位置: {insert_position}")
        
        success, message = UniversalExcelProcessor.add_column(column_name, insert_position)
        
        return jsonify({'success': success, 'message': message})
        
    except Exception as e:
        print(f"[错误] 添加列时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/column/delete', methods=['POST'])
def delete_column():
    try:
        data = request.get_json()
        column_name = data.get('column_name', '').strip()
        
        if not column_name:
            return jsonify({'success': False, 'message': '列名不能为空'})
        
        print(f"[系统] 删除列: {column_name}")
        
        success, message = UniversalExcelProcessor.delete_column(column_name)
        
        return jsonify({'success': success, 'message': message})
        
    except Exception as e:
        print(f"[错误] 删除列时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/column/rename', methods=['POST'])
def rename_column():
    try:
        data = request.get_json()
        old_name = data.get('old_name', '').strip()
        new_name = data.get('new_name', '').strip()
        
        if not old_name or not new_name:
            return jsonify({'success': False, 'message': '列名不能为空'})
        
        if old_name == new_name:
            return jsonify({'success': False, 'message': '新列名与原列名相同'})
        
        print(f"[系统] 重命名列: {old_name} -> {new_name}")
        
        success, message = UniversalExcelProcessor.rename_column(old_name, new_name)
        
        return jsonify({'success': success, 'message': message})
        
    except Exception as e:
        print(f"[错误] 重命名列时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/add_row', methods=['POST'])
def add_row():
    try:
        data = request.get_json()
        insert_after_id = data.get('insert_after_id') if data else None
        insert_position = data.get('insert_position', -1) if data else -1
        
        print(f"[系统] 添加新行，插入在ID {insert_after_id} 之后，位置: {insert_position}")
        
        # 使用更新后的add_row方法
        success, message, row_id = UniversalExcelProcessor.add_row(insert_after_id)
        
        if success:
            return jsonify({'success': True, 'message': message, 'id': row_id})
        else:
            return jsonify({'success': False, 'message': message})
        
    except Exception as e:
        print(f"[错误] 添加行时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

# ==================== 智能表格分组 API ====================

@app.route('/table-groups')
def get_table_groups():
    """获取所有表格分组"""
    try:
        from models.database import TableGroup
        groups = TableGroup.query.order_by(TableGroup.updated_at.desc()).all()
        
        groups_data = []
        for group in groups:
            group_dict = group.to_dict()
            
            # 获取列映射信息
            from models.database import ColumnMapping
            mappings = ColumnMapping.query.filter_by(table_group_id=group.id).all()
            group_dict['has_mappings'] = len(mappings) > 0
            group_dict['mapping_count'] = len(mappings)
            
            groups_data.append(group_dict)
        
        return jsonify({
            'success': True,
            'groups': groups_data,
            'total_groups': len(groups_data)
        })
        
    except Exception as e:
        print(f"[错误] 获取表格分组时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/table-groups/<int:group_id>/data')
def get_group_data(group_id):
    """获取指定分组的数据"""
    try:
        from models.database import TableGroup, TableData, TableSchema
        
        # 验证分组是否存在
        group = TableGroup.query.get(group_id)
        if not group:
            return jsonify({'success': False, 'message': '分组不存在'})
        
        # 获取分组的表结构
        schemas = TableSchema.query.filter_by(
            table_group_id=group_id, 
            is_active=True
        ).order_by(TableSchema.column_order).all()
        schema_columns = [s.column_name for s in schemas]
        
        # 获取分组的数据
        data_records = TableData.query.filter_by(
            table_group_id=group_id
        ).order_by(TableData.created_at.asc()).all()
        
        data = [record.to_dict() for record in data_records]
        
        # 统计信息
        stats = {
            'total_records': len(data),
            'source_files': len(set(record.source_file for record in data_records if record.source_file)),
            'total_columns': len(schema_columns),
            'last_update': group.updated_at.strftime('%Y-%m-%d %H:%M:%S')
        }
        
        return jsonify({
            'success': True,
            'data': data,
            'schema': schema_columns,
            'stats': stats,
            'group_info': group.to_dict()
        })
        
    except Exception as e:
        print(f"[错误] 获取分组数据时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/table-groups/<int:group_id>/mappings')
def get_group_mappings(group_id):
    """获取指定分组的列映射关系"""
    try:
        from models.database_v2 import ColumnMapping
        
        mappings = ColumnMapping.query.filter_by(table_group_id=group_id).all()
        mappings_data = [mapping.to_dict() for mapping in mappings]
        
        # 按来源文件分组
        by_file = {}
        for mapping in mappings_data:
            file = mapping['source_file']
            if file not in by_file:
                by_file[file] = []
            by_file[file].append(mapping)
        
        return jsonify({
            'success': True,
            'mappings': mappings_data,
            'by_file': by_file,
            'total_mappings': len(mappings_data)
        })
        
    except Exception as e:
        print(f"[错误] 获取列映射时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/confirm-mapping', methods=['POST'])
def confirm_column_mapping():
    """确认列映射关系"""
    try:
        from models.database_v2 import ColumnMapping
        
        data = request.get_json()
        mapping_id = data.get('mapping_id')
        
        mapping = ColumnMapping.query.get(mapping_id)
        if not mapping:
            return jsonify({'success': False, 'message': '映射不存在'})
        
        mapping.is_confirmed = True
        db.session.commit()
        
        return jsonify({'success': True, 'message': '映射确认成功'})
        
    except Exception as e:
        print(f"[错误] 确认映射时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/table-groups/rename', methods=['POST'])
def rename_table_group():
    try:
        from models.database import TableGroup
        
        data = request.get_json()
        group_id = data.get('group_id')
        new_name = data.get('new_name', '').strip()
        
        if not group_id or not new_name:
            return jsonify({'success': False, 'message': '参数不完整'})
        
        # 查找表格分组
        table_group = TableGroup.query.get(group_id)
        if not table_group:
            return jsonify({'success': False, 'message': '表格分组不存在'})
        
        # 检查名称是否已存在（同一分组名称应该是唯一的）
        existing_group = TableGroup.query.filter(
            TableGroup.group_name == new_name,
            TableGroup.id != group_id
        ).first()
        
        if existing_group:
            return jsonify({'success': False, 'message': f'表格名称 "{new_name}" 已存在'})
        
        # 更新表格名称
        old_name = table_group.group_name
        table_group.group_name = new_name
        
        db.session.commit()
        
        print(f"[系统] 表格重命名成功: {old_name} -> {new_name}")
        return jsonify({'success': True, 'message': '表格重命名成功'})
        
    except Exception as e:
        db.session.rollback()
        print(f"[错误] 重命名表格时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})

@app.route('/table-groups/<int:group_id>/delete', methods=['DELETE'])
def delete_table_group(group_id):
    """删除指定的表格分组"""
    try:
        from models.database import TableGroup, ColumnMapping
        
        # 查找表格分组
        table_group = TableGroup.query.get(group_id)
        if not table_group:
            return jsonify({'success': False, 'message': '表格不存在'})
        
        print(f"[系统] 开始删除表格分组: {table_group.group_name}")
        
        # 删除关联的数据记录
        TableData.query.filter_by(table_group_id=group_id).delete()
        
        # 删除关联的表格结构
        TableSchema.query.filter_by(table_group_id=group_id).delete()
        
        # 删除关联的列映射
        ColumnMapping.query.filter_by(table_group_id=group_id).delete()
        
        # 删除表格分组本身
        db.session.delete(table_group)
        
        db.session.commit()
        
        print(f"[系统] 表格分组删除成功: {table_group.group_name}")
        return jsonify({'success': True, 'message': '表格删除成功'})
        
    except Exception as e:
        db.session.rollback()
        print(f"[错误] 删除表格分组时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/table-groups/<int:group_id>/export')
def export_single_group(group_id):
    """导出指定表格分组"""
    try:
        from models.database import TableGroup
        
        # 验证分组是否存在
        group = TableGroup.query.get(group_id)
        if not group:
            return jsonify({'success': False, 'message': '表格不存在'})
        
        print(f"[系统] 开始导出表格分组: {group.group_name}")
        
        # 获取该分组的数据
        data_records = TableData.query.filter_by(table_group_id=group_id).order_by(TableData.created_at.asc()).all()
        
        if not data_records:
            return jsonify({'success': False, 'message': '表格中没有数据'})
            
        # 获取该分组的表格结构
        schemas = TableSchema.query.filter_by(table_group_id=group_id, is_active=True).order_by(TableSchema.column_order).all()
        schema_columns = [s.column_name for s in schemas]
        
        # 系统字段列表（不应该出现在导出中）
        system_fields = {'id', 'source_file', 'created_at', 'updated_at', 'table_group_id'}
        
        # 过滤掉系统字段
        business_columns = [col for col in schema_columns if col not in system_fields and col.strip()]
        
        if not business_columns:
            return jsonify({'success': False, 'message': '没有可导出的数据列'})
        
        # 准备数据
        rows = []
        for record in data_records:
            row = {}
            has_data = False
            
            for col in business_columns:
                value = getattr(record, col, '') if hasattr(record, col) else record.to_dict().get(col, '')
                
                # 清理数据
                if value is not None:
                    value = str(value).strip()
                else:
                    value = ''
                row[col] = value
                
                # 检查是否有有效数据
                if value and value != '':
                    has_data = True
            
            if has_data:
                rows.append(row)
        
        if not rows:
            return jsonify({'success': False, 'message': '没有有效数据'})
        
        # 创建DataFrame
        if HAS_PANDAS:
            df = pd.DataFrame(rows, columns=business_columns)
        else:
            # 简单的数据结构替代pandas
            df = {"rows": rows, "columns": business_columns}
        
        # 生成导出文件
        safe_group_name = group.group_name.replace('表格组_', '').replace('_', '-')
        current_time = datetime.now()
        date_str = current_time.strftime("%Y%m%d_%H%M%S")
        
        export_filename = f'{safe_group_name}_{len(rows)}条_{date_str}.xlsx'
        export_path = os.path.join('static/uploads', export_filename)
        
        # 确保导出目录存在
        os.makedirs(os.path.dirname(export_path), exist_ok=True)
        
        # 使用openpyxl写入Excel
        from openpyxl import Workbook
        from openpyxl.utils.dataframe import dataframe_to_rows
        
        wb = Workbook()
        ws = wb.active
        ws.title = safe_group_name[:31]  # Excel工作表名称长度限制
        
        # 写入数据
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # 应用样式
        apply_enhanced_excel_styling(ws, len(business_columns), len(rows))
        
        # 保存文件
        wb.save(export_path)
        
        print(f"[系统] 表格导出成功: {export_filename}, 共导出 {len(rows)} 条记录")
        
        return send_file(export_path, as_attachment=True, download_name=export_filename)
        
    except Exception as e:
        print(f"[错误] 导出表格分组时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})

@app.route('/clear-all', methods=['POST'])
def clear_all_data():
    """清空所有数据和表格分组"""
    try:
        from models.database import TableGroup, ColumnMapping
        
        print("[系统] 开始清空所有数据...")
        
        # 删除所有数据记录
        TableData.query.delete()
        
        # 删除所有表格结构
        TableSchema.query.delete()
        
        # 删除所有表格分组
        TableGroup.query.delete()
        
        # 删除所有列映射
        ColumnMapping.query.delete()
        
        # 删除所有上传历史
        UploadHistory.query.delete()
        
        db.session.commit()
        
        print("[系统] 所有数据已清空")
        return jsonify({'success': True, 'message': '所有数据已清空'})
        
    except Exception as e:
        db.session.rollback()
        print(f"[错误] 清空数据时出错: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})

def create_app():
    """应用程序工厂函数"""
    with app.app_context():
        # 创建数据库表
        db.create_all()
        print("[系统] 数据库初始化完成")
        
        # 确保上传目录存在
        if not os.path.exists(app.config['UPLOAD_FOLDER']):
            os.makedirs(app.config['UPLOAD_FOLDER'])
            print("[系统] 上传目录创建完成")
    
    return app

if __name__ == '__main__':
    app = create_app()
    # 本地开发环境
    port = int(os.environ.get('PORT', 5002))
    print(f"[系统] 启动通用表格合并系统... 端口: {port}")
    app.run(debug=os.environ.get('FLASK_ENV') != 'production', host='0.0.0.0', port=port)
else:
    # 生产环境 (Render/Gunicorn)
    app = create_app()