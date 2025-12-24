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

# Danh sách chức danh SQCM-KT-NV
CHUC_DANH_LIST = [
    'Giảng viên',
    'Giảng viên chính',
    'Giảng viên cao cấp',
    'Trợ giảng',
    'Nghiên cứu viên',
    'Nghiên cứu viên chính',
    'Nghiên cứu viên cao cấp',
    'Biên tập viên',
    'Biên tập viên chính',
    'Biên tập viên cao cấp',
    'Kỹ sư',
    'Kỹ sư chính',
    'Kỹ sư cao cấp',
    'Bác sĩ',
    'Bác sĩ chính',
    'Bác sĩ cao cấp',
    'Thư viện viên',
    'Thư viện viên chính',
    'Thư viện viên cao cấp',
]

# Danh sách danh hiệu nhà giáo
NHA_GIAO_LIST = [
    'Nhà giáo nhân dân',
    'Nhà giáo ưu tú',
    'Nhà giáo dạy giỏi cấp trường',
    'Nhà giáo dạy giỏi cấp nhà nước',
]

# Danh sách vai trò trong Ban/Hội đồng
VAI_TRO_BAN = ['Chủ nhiệm', 'Thư ký', 'Thành viên']
VAI_TRO_HOI_DONG = ['Chủ tịch', 'Thư ký', 'Thành viên']

# Danh sách loại tài liệu
LOAI_TAI_LIEU = [
    'Giáo trình',
    'Tài liệu dạy học',
    'Sách chuyên khảo',
    'Sách tham khảo',
]

# Danh sách cấp độ
CAP_DO_LIST = [
    'Cấp nhà trường',
    'Cấp bộ',
    'Cấp nhà nước',
]

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

# ==================== ĐỀ TÀI ====================

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
    cap_do = db.Column(db.String(50))  # Cấp nhà trường, Cấp bộ, Cấp nhà nước
    ngay_bat_dau = db.Column(db.Date)
    ngay_ket_thuc = db.Column(db.Date)
    trang_thai = db.Column(db.String(50), default='Đang thực hiện')
    ket_qua = db.Column(db.Text)
    file_word = db.Column(db.String(255))
    file_word_content = db.Column(db.Text)
    
    # File biên bản và tài liệu
    bien_ban_ban = db.Column(db.String(255))  # Biên bản của Ban đề tài
    bien_ban_de_cuong = db.Column(db.String(255))  # Biên bản duyệt đề cương
    bien_ban_nghiem_thu = db.Column(db.String(255))  # Biên bản nghiệm thu
    file_sau_sua = db.Column(db.String(255))  # File đề tài sau khi sửa
    
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    creator = db.relationship('User', backref='de_tai_created')
    ban_thanh_vien = db.relationship('BanDeTai', backref='de_tai', lazy='dynamic', cascade='all, delete-orphan')
    hoi_dong_de_cuong = db.relationship('HoiDongDeCuongDeTai', backref='de_tai', lazy='dynamic', cascade='all, delete-orphan')
    hoi_dong_nghiem_thu = db.relationship('HoiDongNghiemThuDeTai', backref='de_tai', lazy='dynamic', cascade='all, delete-orphan')

class BanDeTai(db.Model):
    """Model Ban đề tài"""
    __tablename__ = 'ban_de_tai'
    
    id = db.Column(db.Integer, primary_key=True)
    de_tai_id = db.Column(db.Integer, db.ForeignKey('de_tai.id'), nullable=False)
    ho_ten = db.Column(db.String(120), nullable=False)
    vai_tro = db.Column(db.String(50), nullable=False)  # Chủ nhiệm, Thư ký, Thành viên
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class HoiDongDeCuongDeTai(db.Model):
    """Model Hội đồng duyệt đề cương của đề tài"""
    __tablename__ = 'hoi_dong_de_cuong_de_tai'
    
    id = db.Column(db.Integer, primary_key=True)
    de_tai_id = db.Column(db.Integer, db.ForeignKey('de_tai.id'), nullable=False)
    ho_ten = db.Column(db.String(120), nullable=False)
    vai_tro = db.Column(db.String(50), nullable=False)  # Chủ tịch, Thư ký, Thành viên
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class HoiDongNghiemThuDeTai(db.Model):
    """Model Hội đồng nghiệm thu của đề tài"""
    __tablename__ = 'hoi_dong_nghiem_thu_de_tai'
    
    id = db.Column(db.Integer, primary_key=True)
    de_tai_id = db.Column(db.Integer, db.ForeignKey('de_tai.id'), nullable=False)
    ho_ten = db.Column(db.String(120), nullable=False)
    vai_tro = db.Column(db.String(50), nullable=False)  # Chủ tịch, Thư ký, Thành viên
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

# ==================== SÁNG KIẾN ====================

class SangKien(db.Model):
    """Model sáng kiến (nội dung như đề tài)"""
    __tablename__ = 'sang_kien'
    
    id = db.Column(db.Integer, primary_key=True)
    ma_sang_kien = db.Column(db.String(50), unique=True, nullable=False)
    ten_sang_kien = db.Column(db.String(255), nullable=False)
    mo_ta = db.Column(db.Text)
    linh_vuc = db.Column(db.String(100))
    chu_nhiem = db.Column(db.String(120), nullable=False)
    don_vi = db.Column(db.String(120))
    cap_do = db.Column(db.String(50))  # Cấp nhà trường, Cấp bộ, Cấp nhà nước
    ngay_bat_dau = db.Column(db.Date)
    ngay_ket_thuc = db.Column(db.Date)
    trang_thai = db.Column(db.String(50), default='Đang thực hiện')
    ket_qua = db.Column(db.Text)
    file_word = db.Column(db.String(255))
    file_word_content = db.Column(db.Text)
    
    # File biên bản và tài liệu
    bien_ban_ban = db.Column(db.String(255))
    bien_ban_de_cuong = db.Column(db.String(255))
    bien_ban_nghiem_thu = db.Column(db.String(255))
    file_sau_sua = db.Column(db.String(255))
    
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    creator = db.relationship('User', backref='sang_kien_created')
    ban_thanh_vien = db.relationship('BanSangKien', backref='sang_kien', lazy=True, cascade='all, delete-orphan')
    hoi_dong_de_cuong = db.relationship('HoiDongDeCuongSangKien', backref='sang_kien', lazy=True, cascade='all, delete-orphan')
    hoi_dong_nghiem_thu = db.relationship('HoiDongNghiemThuSangKien', backref='sang_kien', lazy=True, cascade='all, delete-orphan')

class BanSangKien(db.Model):
    """Model Ban sáng kiến"""
    __tablename__ = 'ban_sang_kien'
    
    id = db.Column(db.Integer, primary_key=True)
    sang_kien_id = db.Column(db.Integer, db.ForeignKey('sang_kien.id'), nullable=False)
    ho_ten = db.Column(db.String(120), nullable=False)
    vai_tro = db.Column(db.String(50), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class HoiDongDeCuongSangKien(db.Model):
    """Model Hội đồng duyệt đề cương của sáng kiến"""
    __tablename__ = 'hoi_dong_de_cuong_sang_kien'
    
    id = db.Column(db.Integer, primary_key=True)
    sang_kien_id = db.Column(db.Integer, db.ForeignKey('sang_kien.id'), nullable=False)
    ho_ten = db.Column(db.String(120), nullable=False)
    vai_tro = db.Column(db.String(50), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class HoiDongNghiemThuSangKien(db.Model):
    """Model Hội đồng nghiệm thu của sáng kiến"""
    __tablename__ = 'hoi_dong_nghiem_thu_sang_kien'
    
    id = db.Column(db.Integer, primary_key=True)
    sang_kien_id = db.Column(db.Integer, db.ForeignKey('sang_kien.id'), nullable=False)
    ho_ten = db.Column(db.String(120), nullable=False)
    vai_tro = db.Column(db.String(50), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

# ==================== TÀI LIỆU ====================

class TaiLieu(db.Model):
    """Model tài liệu (giáo trình, tài liệu dạy học, sách chuyên khảo, sách tham khảo)"""
    __tablename__ = 'tai_lieu'
    
    id = db.Column(db.Integer, primary_key=True)
    ma_tai_lieu = db.Column(db.String(50), unique=True, nullable=False)
    ten_tai_lieu = db.Column(db.String(255), nullable=False)
    loai = db.Column(db.String(50))  # Giáo trình, Tài liệu dạy học, Sách chuyên khảo, Sách tham khảo
    mo_ta = db.Column(db.Text)
    linh_vuc = db.Column(db.String(100))
    tac_gia = db.Column(db.Text)  # Danh sách tác giả (mỗi tác giả 1 dòng)
    don_vi = db.Column(db.String(120))
    ngay_bat_dau = db.Column(db.Date)
    ngay_ket_thuc = db.Column(db.Date)
    trang_thai = db.Column(db.String(50), default='Đang thực hiện')
    ket_qua = db.Column(db.Text)
    file_word = db.Column(db.String(255))
    file_word_content = db.Column(db.Text)
    
    # Thông tin xuất bản
    nha_xuat_ban = db.Column(db.String(120))
    nam_xuat_ban = db.Column(db.Integer)
    mon_hoc = db.Column(db.String(120))
    
    # File tài liệu (PDF hoặc Word)
    file_tai_lieu = db.Column(db.String(255))
    
    # File biên bản và tài liệu
    bien_ban_ban = db.Column(db.String(255))
    bien_ban_de_cuong = db.Column(db.String(255))
    bien_ban_nghiem_thu = db.Column(db.String(255))
    file_sau_sua = db.Column(db.String(255))
    
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    creator = db.relationship('User', backref='tai_lieu_created')
    ban_thanh_vien = db.relationship('BanTaiLieu', backref='tai_lieu', lazy='dynamic', cascade='all, delete-orphan')
    hoi_dong_de_cuong = db.relationship('HoiDongDeCuongTaiLieu', backref='tai_lieu', lazy='dynamic', cascade='all, delete-orphan')
    hoi_dong_nghiem_thu = db.relationship('HoiDongNghiemThuTaiLieu', backref='tai_lieu', lazy='dynamic', cascade='all, delete-orphan')

class BanTaiLieu(db.Model):
    """Model Ban tài liệu"""
    __tablename__ = 'ban_tai_lieu'
    
    id = db.Column(db.Integer, primary_key=True)
    tai_lieu_id = db.Column(db.Integer, db.ForeignKey('tai_lieu.id'), nullable=False)
    ho_ten = db.Column(db.String(120), nullable=False)
    vai_tro = db.Column(db.String(50), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class HoiDongDeCuongTaiLieu(db.Model):
    """Model Hội đồng duyệt đề cương của tài liệu"""
    __tablename__ = 'hoi_dong_de_cuong_tai_lieu'
    
    id = db.Column(db.Integer, primary_key=True)
    tai_lieu_id = db.Column(db.Integer, db.ForeignKey('tai_lieu.id'), nullable=False)
    ho_ten = db.Column(db.String(120), nullable=False)
    vai_tro = db.Column(db.String(50), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class HoiDongNghiemThuTaiLieu(db.Model):
    """Model Hội đồng nghiệm thu của tài liệu"""
    __tablename__ = 'hoi_dong_nghiem_thu_tai_lieu'
    
    id = db.Column(db.Integer, primary_key=True)
    tai_lieu_id = db.Column(db.Integer, db.ForeignKey('tai_lieu.id'), nullable=False)
    ho_ten = db.Column(db.String(120), nullable=False)
    vai_tro = db.Column(db.String(50), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

# ==================== TIỀM LỰC ====================

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
    
    # Chức danh SQCM-KT-NV
    chuc_danh = db.Column(db.String(100))
    
    # Danh hiệu nhà giáo
    nha_giao = db.Column(db.String(100))
    
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

# Model cũ để tương thích (sẽ bị bỏ)
class GiaoTrinh(db.Model):
    """Model giáo trình (deprecated - sử dụng TaiLieu thay thế)"""
    __tablename__ = 'giao_trinh'
    
    id = db.Column(db.Integer, primary_key=True)
    ma_tai_lieu = db.Column(db.String(50), unique=True, nullable=False)
    ten_tai_lieu = db.Column(db.String(255), nullable=False)
    loai = db.Column(db.String(50))
    tac_gia = db.Column(db.String(255))
    nha_xuat_ban = db.Column(db.String(120))
    nam_xuat_ban = db.Column(db.Integer)
    mon_hoc = db.Column(db.String(120))
    mo_ta = db.Column(db.Text)
    don_vi = db.Column(db.String(120))
    file_path = db.Column(db.String(255))
    file_word = db.Column(db.String(255))
    file_word_content = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    creator = db.relationship('User', backref='giao_trinh_created')
