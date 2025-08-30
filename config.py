import os
from urllib.parse import urlparse

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'universal-table-merger-2025-fallback'
    
    # 数据库配置
    DATABASE_URL = os.environ.get('DATABASE_URL')
    if DATABASE_URL:
        # Render PostgreSQL 数据库
        url = urlparse(DATABASE_URL)
        SQLALCHEMY_DATABASE_URI = DATABASE_URL.replace('postgres://', 'postgresql://', 1) if DATABASE_URL.startswith('postgres://') else DATABASE_URL
    else:
        # 本地开发使用 SQLite
        SQLALCHEMY_DATABASE_URI = 'sqlite:///universal_table_processor.db'
    
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    
    # 文件上传配置
    UPLOAD_FOLDER = os.path.join(os.getcwd(), 'static', 'uploads')
    MAX_CONTENT_LENGTH = 32 * 1024 * 1024  # 32MB
    
    # 确保上传目录存在
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

class DevelopmentConfig(Config):
    DEBUG = True
    SQLALCHEMY_DATABASE_URI = 'sqlite:///universal_table_processor.db'

class ProductionConfig(Config):
    DEBUG = False
    
config = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'default': DevelopmentConfig
}
