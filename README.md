# HỆ THỐNG QUẢN LÝ KHOA HỌC QUÂN SỰ

Hệ thống quản lý khoa học quân sự trên nền tảng web sử dụng Flask framework.

## Tính năng

✅ **Đăng nhập/Đăng ký** - Hệ thống xác thực người dùng an toàn  
✅ **Quản lý Đề tài** - Quản lý đầy đủ các đề tài nghiên cứu khoa học  
✅ **Quản lý Giáo trình & Tài liệu** - Quản lý tài liệu, giáo trình, sách tham khảo  
✅ **Quản lý Tiềm lực Khoa học** - Quản lý cán bộ, giảng viên, nhà khoa học  
✅ **Dashboard** - Thống kê tổng quan hệ thống  
✅ **Tìm kiếm** - Tìm kiếm nhanh trong từng module  
✅ **Giao diện hiện đại** - Responsive, thân thiện với người dùng  

## Công nghệ sử dụng

- **Backend**: Flask 2.3.3
- **Database**: SQLite (SQLAlchemy ORM)
- **Authentication**: Flask-Login
- **Frontend**: HTML5, CSS3, JavaScript (jQuery)
- **Icons**: Font Awesome 6.4.0

## Cài đặt

### 1. Clone hoặc tải project

```bash
cd d:\Codes\webkh
```

### 2. Cài đặt các thư viện

```bash
pip install -r requirements.txt
```

Hoặc cài đặt thủ công:

```bash
pip install Flask Flask-SQLAlchemy Flask-Login Flask-WTF WTForms Werkzeug email-validator python-dotenv
```

### 3. Chạy ứng dụng

```bash
python app.py
```

Ứng dụng sẽ chạy tại: **http://localhost:5000**

## Tài khoản mặc định

Khi chạy lần đầu, hệ thống tự động tạo tài khoản admin:

- **Username**: `admin`
- **Password**: `admin123`
- **Email**: `admin@military.edu.vn`

## Cấu trúc thư mục

```
webkh/
├── app.py                  # File chính chạy ứng dụng
├── config.py               # Cấu hình ứng dụng
├── requirements.txt        # Danh sách thư viện
├── models/
│   └── __init__.py        # Models database
├── routes/
│   ├── __init__.py
│   ├── auth.py            # Routes xác thực
│   └── main.py            # Routes chính
├── static/
│   ├── css/
│   │   └── style.css      # CSS chính
│   ├── js/
│   ├── images/
│   └── uploads/           # Thư mục upload file
├── templates/
│   ├── base.html          # Template cơ sở
│   ├── login.html         # Trang đăng nhập
│   ├── register.html      # Trang đăng ký
│   ├── dashboard.html     # Trang dashboard
│   ├── detai/             # Templates đề tài
│   ├── giaotrinh/         # Templates giáo trình
│   └── tiemluc/           # Templates tiềm lực
└── utils/                 # Các tiện ích
```

## Database Models

### User (Người dùng)
- username, email, password_hash
- full_name, role (admin/manager/user)
- created_at

### DeTai (Đề tài nghiên cứu)
- ma_de_tai, ten_de_tai, mo_ta
- linh_vuc, chu_nhiem, don_vi
- ngay_bat_dau, ngay_ket_thuc, kinh_phi
- trang_thai, ket_qua

### GiaoTrinh (Giáo trình & Tài liệu)
- ma_tai_lieu, ten_tai_lieu, loai
- tac_gia, nha_xuat_ban, nam_xuat_ban
- mon_hoc, mo_ta

### TiemLuc (Tiềm lực khoa học)
- ma_can_bo, ho_ten, ngay_sinh, gioi_tinh
- don_vi, chuc_vu, hoc_vi, hoc_ham
- chuyen_mon, email, dien_thoai
- kinh_nghiem, thanh_tich

## Tính năng chính

### 1. Dashboard
- Thống kê tổng số đề tài, giáo trình, tiềm lực
- Hiển thị đề tài gần đây
- Quick actions để thêm mới nhanh

### 2. Quản lý Đề tài
- Thêm/Sửa/Xóa đề tài
- Tìm kiếm theo mã, tên, chủ nhiệm
- Quản lý trạng thái: Đang thực hiện, Hoàn thành, Tạm dừng
- Theo dõi kinh phí, thời gian thực hiện

### 3. Quản lý Giáo trình & Tài liệu
- Thêm/Sửa/Xóa tài liệu
- Phân loại: Giáo trình, Tài liệu, Sách tham khảo, Bài báo
- Quản lý thông tin xuất bản

### 4. Quản lý Tiềm lực Khoa học
- Thêm/Sửa/Xóa cán bộ
- Quản lý học vị, học hàm
- Theo dõi kinh nghiệm và thành tích

## Bảo mật

- Mật khẩu được mã hóa bằng Werkzeug
- Session timeout sau 2 giờ
- Login required cho tất cả trang quản lý
- CSRF protection với Flask-WTF

## Customization

### Thay đổi màu sắc
Chỉnh sửa biến CSS trong `static/css/style.css`:

```css
:root {
    --primary-color: #2c5f2d;
    --secondary-color: #97b885;
    --accent-color: #ff6b35;
}
```

### Thay đổi cấu hình
Chỉnh sửa file `config.py`:

```python
SECRET_KEY = 'your-secret-key'
SQLALCHEMY_DATABASE_URI = 'sqlite:///your_database.db'
```

## Hỗ trợ

Nếu gặp vấn đề, vui lòng kiểm tra:

1. Python version >= 3.7
2. Tất cả thư viện đã được cài đặt
3. Port 5000 chưa được sử dụng
4. Quyền ghi file trong thư mục dự án

## License

MIT License - Tự do sử dụng cho mục đích học tập và nghiên cứu.

## Phát triển bởi

Hệ thống Quản lý Khoa học Quân sự - 2025
