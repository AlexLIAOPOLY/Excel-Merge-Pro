from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import json
from difflib import SequenceMatcher

db = SQLAlchemy()

class TableGroup(db.Model):
    """表格分组模型 - 管理相同/相似结构的表格"""
    __tablename__ = 'table_groups'
    
    id = db.Column(db.Integer, primary_key=True)
    group_name = db.Column(db.String(200))         # 分组名称
    description = db.Column(db.Text)               # 分组描述
    schema_fingerprint = db.Column(db.String(500)) # 表头指纹用于匹配
    column_count = db.Column(db.Integer)           # 列数
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # 关联关系
    data_records = db.relationship('TableData', backref='table_group', lazy=True)
    schemas = db.relationship('TableSchema', backref='table_group', lazy=True)
    
    def to_dict(self):
        return {
            'id': self.id,
            'group_name': self.group_name,
            'description': self.description,
            'column_count': self.column_count,
            'record_count': len(self.data_records),
            'created_at': self.created_at.strftime('%Y-%m-%d %H:%M:%S'),
            'updated_at': self.updated_at.strftime('%Y-%m-%d %H:%M:%S')
        }

class TableData(db.Model):
    """通用表格数据模型，支持任意列结构"""
    __tablename__ = 'table_data_v2'
    
    id = db.Column(db.Integer, primary_key=True)
    source_file = db.Column(db.String(200))        # 来源文件名
    row_data = db.Column(db.Text)                  # JSON格式存储行数据
    table_group_id = db.Column(db.Integer, db.ForeignKey('table_groups.id')) # 关联表格分组
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def get_data(self):
        """获取行数据"""
        return json.loads(self.row_data) if self.row_data else {}
    
    def set_data(self, data):
        """设置行数据"""
        self.row_data = json.dumps(data, ensure_ascii=False)
    
    def to_dict(self):
        data = self.get_data()
        data['id'] = self.id
        data['source_file'] = self.source_file
        data['created_at'] = self.created_at.strftime('%Y-%m-%d %H:%M:%S')
        data['updated_at'] = self.updated_at.strftime('%Y-%m-%d %H:%M:%S')
        return data

class TableSchema(db.Model):
    """表格结构信息"""
    __tablename__ = 'table_schema'
    
    id = db.Column(db.Integer, primary_key=True)
    column_name = db.Column(db.String(200))        # 列名
    column_type = db.Column(db.String(50))         # 列类型
    column_order = db.Column(db.Integer)           # 列顺序
    is_active = db.Column(db.Boolean, default=True) # 是否激活
    table_group_id = db.Column(db.Integer, db.ForeignKey('table_groups.id')) # 关联表格分组
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    def to_dict(self):
        return {
            'id': self.id,
            'column_name': self.column_name,
            'column_type': self.column_type,
            'column_order': self.column_order,
            'is_active': self.is_active,
            'table_group_id': self.table_group_id
        }

class ColumnMapping(db.Model):
    """列名映射表 - 处理相似列名的映射关系"""
    __tablename__ = 'column_mappings'
    
    id = db.Column(db.Integer, primary_key=True)
    table_group_id = db.Column(db.Integer, db.ForeignKey('table_groups.id'))
    original_column = db.Column(db.String(200))    # 原始列名
    mapped_column = db.Column(db.String(200))      # 映射后的列名
    source_file = db.Column(db.String(200))        # 来源文件
    similarity_score = db.Column(db.Float)         # 相似度分数
    is_confirmed = db.Column(db.Boolean, default=False) # 是否已确认映射
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    def to_dict(self):
        return {
            'id': self.id,
            'table_group_id': self.table_group_id,
            'original_column': self.original_column,
            'mapped_column': self.mapped_column,
            'source_file': self.source_file,
            'similarity_score': self.similarity_score,
            'is_confirmed': self.is_confirmed,
            'created_at': self.created_at.strftime('%Y-%m-%d %H:%M:%S')
        }

class UploadHistory(db.Model):
    """上传历史记录"""
    __tablename__ = 'upload_history_v2'
    
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(200))
    upload_time = db.Column(db.DateTime, default=datetime.utcnow)
    rows_imported = db.Column(db.Integer)
    columns_detected = db.Column(db.Text)          # JSON格式存储检测到的列
    status = db.Column(db.String(20))              # success, failed
    error_message = db.Column(db.Text)
    
    def get_columns(self):
        """获取检测到的列"""
        return json.loads(self.columns_detected) if self.columns_detected else []
    
    def set_columns(self, columns):
        """设置检测到的列"""
        self.columns_detected = json.dumps(columns, ensure_ascii=False)
    
    def to_dict(self):
        return {
            'id': self.id,
            'filename': self.filename,
            'upload_time': self.upload_time.strftime('%Y-%m-%d %H:%M:%S'),
            'rows_imported': self.rows_imported,
            'columns_detected': self.get_columns(),
            'status': self.status,
            'error_message': self.error_message
        }