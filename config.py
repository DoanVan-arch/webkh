import os
from datetime import timedelta

class Config:
    """Cấu hình cơ bản cho ứng dụng"""
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'khoa-hoc-quan-su-secret-key-2025'
    SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL') or 'sqlite:///khoa_hoc_quan_su.db'
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    PERMANENT_SESSION_LIFETIME = timedelta(hours=2)
    
    # Upload configuration
    UPLOAD_FOLDER = 'static/uploads'
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size
    ALLOWED_EXTENSIONS = {'pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt'}
