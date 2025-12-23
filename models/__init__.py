from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()

# Danh sách đơn vị
DON_VI_LIST = {
    'Khối cơ quan': [
        'Phòng Tham mưu - Hành chính (B1)',
        'Phòng Đào tạo (B2)',
        'Phòng Chính trị (B3)',
        'Phòng Khoa học quân sự (B4)',
        'Phòng Hậu cần - Kỹ thuật (B5)',
        'Ban Tài Chính (B6)',
        'Ban Khảo thí và đảm bảo chất lượng giáo dục đào tạo (B7)',
        'Ban Sau Đại học (B8)',
        'Ban thông tin khoa học quân sự (B9)',
        'Uỷ Ban kiểm tra Đảng (B11)',
        'Thanh tra Nhà trường (B12)',
    ],
    'Các khoa': [
        'Khoa Triết học Mác - Lênin (K1)',
        'Khoa Lịch sử Đảng Cộng sản Việt Nam (K2)',
        'Khoa Công tác Đảng,Công tác Chính trị (K3)',
        'Khoa Chiến thuật (K4)',
        'Khoa Văn hóa - Ngoại ngữ (K5)',
        'Khoa Kinh tế chính trị Mác - Lênin (K6)',
        'Khoa Chủ nghĩa xã hội khoa học (K7)',
        'Khoa Tâm lý học quân sự (K8)',
        'Khoa Bắn súng (K9)',
        'Khoa Quân sự chung (K10)',
        'Khoa Giáo dục thể chất (K11)',
        'Khoa Sư phạm quân sự (K12)',
        'Khoa Tư tưởng Hồ Chí Minh (K13)',
        'Khoa Nhà nước & Pháp luật (K14)',
    ],
    'Đơn vị trực thuộc': [
        'Hệ 1 (Hệ Chuyển loại cán bộ chính trị, hoàn thiện đại học)',
        'Hệ 2 (Hệ sau đại học)',
        'Hệ 3 (Hệ quốc tế)',
        'Tiểu đoàn 1',
        'Tiểu đoàn 2',
        'Tiểu đoàn 3',
        'Tiểu đoàn 4',
        'Tiểu đoàn 5',
        'Tiểu đoàn 6',
        'Tiểu đoàn 7',
        'Tiểu đoàn 8',
        'Tiểu đoàn 9',
        'Tiểu đoàn 10',
        'Tiểu đoàn 11',
        'Tiểu đoàn 12',
    ]
}

class User(UserMixin, db.Model):
    """Model người dùng"""
    __tablename__ = 'users'
    
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(120), nullable=False)
    role = db.Column(db.String(20), default='user')  # admin, manager, user
    don_vi = db.Column(db.String(120))  # Đơn vị của người dùng
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    def set_password(self, password):
        self.password_hash = generate_password_hash(password)
    
    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

class DeTai(db.Model):
    """Model đề tài nghiên cứu"""
    __tablename__ = 'de_tai'
    
    id = db.Column(db.Integer, primary_key=True)
    ma_de_tai = db.Column(db.String(50), unique=True, nullable=False)
    ten_de_tai = db.Column(db.String(255), nullable=False)
    mo_ta = db.Column(db.Text)
    linh_vuc = db.Column(db.String(100))
    chu_nhiem = db.Column(db.String(120), nullable=False)
    don_vi = db.Column(db.String(120))
    ngay_bat_dau = db.Column(db.Date)
    ngay_ket_thuc = db.Column(db.Date)
    kinh_phi = db.Column(db.Float, default=0)
    trang_thai = db.Column(db.String(50), default='Đang thực hiện')  # Đang thực hiện, Hoàn thành, Tạm dừng
    ket_qua = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    creator = db.relationship('User', backref='de_tai_created')

class GiaoTrinh(db.Model):
    """Model giáo trình và tài liệu"""
    __tablename__ = 'giao_trinh'
    
    id = db.Column(db.Integer, primary_key=True)
    ma_tai_lieu = db.Column(db.String(50), unique=True, nullable=False)
    ten_tai_lieu = db.Column(db.String(255), nullable=False)
    loai = db.Column(db.String(50))  # Giáo trình, Tài liệu, Sách tham khảo
    tac_gia = db.Column(db.String(255))
    nha_xuat_ban = db.Column(db.String(120))
    nam_xuat_ban = db.Column(db.Integer)
    mon_hoc = db.Column(db.String(120))
    mo_ta = db.Column(db.Text)
    don_vi = db.Column(db.String(120))  # Đơn vị sở hữu tài liệu
    file_path = db.Column(db.String(255))
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    creator = db.relationship('User', backref='tai_lieu_created')

class TiemLuc(db.Model):
    """Model tiềm lực khoa học (giảng viên, nhà khoa học)"""
    __tablename__ = 'tiem_luc'
    
    id = db.Column(db.Integer, primary_key=True)
    ma_can_bo = db.Column(db.String(50), unique=True, nullable=False)
    ho_ten = db.Column(db.String(120), nullable=False)
    ngay_sinh = db.Column(db.Date)
    gioi_tinh = db.Column(db.String(10))
    don_vi = db.Column(db.String(120))
    chuc_vu = db.Column(db.String(100))
    hoc_vi = db.Column(db.String(50))  # Tiến sĩ, Thạc sĩ, Cử nhân
    hoc_ham = db.Column(db.String(50))  # Giáo sư, Phó Giáo sư
    chuyen_mon = db.Column(db.String(255))
    email = db.Column(db.String(120))
    dien_thoai = db.Column(db.String(20))
    kinh_nghiem = db.Column(db.Text)
    thanh_tich = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    creator = db.relationship('User', backref='tiem_luc_created')
