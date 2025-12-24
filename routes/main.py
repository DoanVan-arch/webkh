from flask import Blueprint, render_template, redirect, url_for, flash, request, send_file
from flask_login import login_required, current_user
from models import (db, DeTai, TaiLieu, TiemLuc, SangKien, User, DON_VI_LIST, LOAI_TAI_LIEU, 
                    CHUC_DANH_LIST, NHA_GIAO_LIST, VAI_TRO_BAN, VAI_TRO_HOI_DONG,
                    BanDeTai, HoiDongDeCuongDeTai, HoiDongNghiemThuDeTai,
                    BanSangKien, HoiDongDeCuongSangKien, HoiDongNghiemThuSangKien,
                    BanTaiLieu, HoiDongDeCuongTaiLieu, HoiDongNghiemThuTaiLieu, CAP_DO_LIST)
from datetime import datetime
from sqlalchemy import func
from utils.file_handler import save_uploaded_file
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

main_bp = Blueprint('main', 'main')

@main_bp.route('/')
def index():
    """Trang chủ"""
    if current_user.is_authenticated:
        return redirect(url_for('main.dashboard'))
    return redirect(url_for('auth.login'))

@main_bp.route('/dashboard')
@login_required
def dashboard():
    """Trang dashboard"""
    # Thống kê
    total_detai = DeTai.query.count()
    total_tailieu = TaiLieu.query.count()
    total_tiemluc = TiemLuc.query.count()
    total_sangkien = SangKien.query.count()
    total_users = User.query.count()
    
    # Đề tài gần đây
    recent_detai = DeTai.query.order_by(DeTai.created_at.desc()).limit(5).all()
    
    return render_template('dashboard.html',
                         total_detai=total_detai,
                         total_giaotrinh=total_tailieu,  # Giữ tên biến cũ để tương thích template
                         total_tiemluc=total_tiemluc,
                         total_sangkien=total_sangkien,
                         total_users=total_users,
                         recent_detai=recent_detai)

# ==================== QUẢN LÝ ĐỀ TÀI ====================
@main_bp.route('/detai')
@login_required
def detai_list():
    """Danh sách đề tài"""
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '', type=str)
    
    query = DeTai.query
    if search:
        query = query.filter(
            db.or_(
                DeTai.ten_de_tai.contains(search),
                DeTai.ma_de_tai.contains(search),
                DeTai.chu_nhiem.contains(search)
            )
        )
    
    pagination = query.order_by(DeTai.created_at.desc()).paginate(
        page=page, per_page=10, error_out=False
    )
    
    return render_template('detai/list.html', pagination=pagination, search=search)

@main_bp.route('/detai/add', methods=['GET', 'POST'])
@login_required
def detai_add():
    """Thêm đề tài mới - Tự động gán đơn vị của user"""
    if request.method == 'POST':
        # Chỉ cho phép thêm với đơn vị của mình (nếu không phải admin)
        don_vi = current_user.don_vi if current_user.role != 'admin' else request.form.get('don_vi')
        
        detai = DeTai(
            ma_de_tai=request.form.get('ma_de_tai'),
            ten_de_tai=request.form.get('ten_de_tai'),
            mo_ta=request.form.get('mo_ta'),
            linh_vuc=request.form.get('linh_vuc'),
            chu_nhiem=request.form.get('chu_nhiem'),
            don_vi=don_vi,
            trang_thai=request.form.get('trang_thai'),
            created_by=current_user.id,
            cap_do=request.form.get('cap_do')
        )
        
        # Xử lý ngày tháng
        ngay_bat_dau = request.form.get('ngay_bat_dau')
        ngay_ket_thuc = request.form.get('ngay_ket_thuc')
        if ngay_bat_dau:
            detai.ngay_bat_dau = datetime.strptime(ngay_bat_dau, '%Y-%m-%d').date()
        if ngay_ket_thuc:
            detai.ngay_ket_thuc = datetime.strptime(ngay_ket_thuc, '%Y-%m-%d').date()
        
        db.session.add(detai)
        db.session.flush()  # Để lấy ID
        
        # Tạo thư mục riêng cho đề tài: detai_<id>_<ma_de_tai>
        folder_name = f"detai_{detai.id}_{detai.ma_de_tai}"
        
        # Xử lý upload các file
        if 'bien_ban_ban' in request.files:
            file = request.files['bien_ban_ban']
            if file and file.filename:
                detai.bien_ban_ban = save_uploaded_file(file, folder_name)
        
        if 'bien_ban_de_cuong' in request.files:
            file = request.files['bien_ban_de_cuong']
            if file and file.filename:
                detai.bien_ban_de_cuong = save_uploaded_file(file, folder_name)
        
        if 'bien_ban_nghiem_thu' in request.files:
            file = request.files['bien_ban_nghiem_thu']
            if file and file.filename:
                detai.bien_ban_nghiem_thu = save_uploaded_file(file, folder_name)
        
        # Xử lý thêm thành viên Ban đề tài
        ban_de_tai_text = request.form.get('ban_de_tai', '').strip()
        if ban_de_tai_text:
            for line in ban_de_tai_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    ban_member = BanDeTai(
                        de_tai_id=detai.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(ban_member)
        
        # Xử lý thêm thành viên Hội đồng duyệt đề cương
        hoi_dong_de_cuong_text = request.form.get('hoi_dong_de_cuong', '').strip()
        if hoi_dong_de_cuong_text:
            for line in hoi_dong_de_cuong_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    hd_member = HoiDongDeCuongDeTai(
                        de_tai_id=detai.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(hd_member)
        
        # Xử lý thêm thành viên Hội đồng nghiệm thu
        hoi_dong_nghiem_thu_text = request.form.get('hoi_dong_nghiem_thu', '').strip()
        if hoi_dong_nghiem_thu_text:
            for line in hoi_dong_nghiem_thu_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    hd_member = HoiDongNghiemThuDeTai(
                        de_tai_id=detai.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(hd_member)
        
        db.session.commit()
        
        flash('Thêm đề tài thành công!', 'success')
        return redirect(url_for('main.detai_list'))
    
    return render_template('detai/add.html', don_vi_list=DON_VI_LIST, cap_do_list=CAP_DO_LIST)

@main_bp.route('/detai/edit/<int:id>', methods=['GET', 'POST'])
@login_required
def detai_edit(id):
    """Sửa đề tài - Chỉ sửa được đề tài của đơn vị mình hoặc admin"""
    detai = DeTai.query.get_or_404(id)
    
    # Kiểm tra quyền: chỉ admin hoặc user cùng đơn vị mới sửa được
    if current_user.role != 'admin' and detai.don_vi != current_user.don_vi:
        flash('Bạn không có quyền chỉnh sửa đề tài của đơn vị khác!', 'danger')
        return redirect(url_for('main.detai_list'))
    
    if request.method == 'POST':
        detai.ma_de_tai = request.form.get('ma_de_tai')
        detai.ten_de_tai = request.form.get('ten_de_tai')
        detai.mo_ta = request.form.get('mo_ta')
        detai.linh_vuc = request.form.get('linh_vuc')
        detai.chu_nhiem = request.form.get('chu_nhiem')
        detai.cap_do = request.form.get('cap_do')
        
        # Admin có thể đổi đơn vị
        if current_user.role == 'admin':
            detai.don_vi = request.form.get('don_vi')
        
        detai.trang_thai = request.form.get('trang_thai')
        detai.ket_qua = request.form.get('ket_qua')
        
        # Xử lý ngày tháng
        ngay_bat_dau = request.form.get('ngay_bat_dau')
        ngay_ket_thuc = request.form.get('ngay_ket_thuc')
        if ngay_bat_dau:
            detai.ngay_bat_dau = datetime.strptime(ngay_bat_dau, '%Y-%m-%d').date()
        if ngay_ket_thuc:
            detai.ngay_ket_thuc = datetime.strptime(ngay_ket_thuc, '%Y-%m-%d').date()
        
        # Tạo thư mục riêng cho đề tài
        folder_name = f"detai_{detai.id}_{detai.ma_de_tai}"
        
        # Xử lý upload các file (chỉ upload nếu có file mới)
        if 'bien_ban_ban' in request.files:
            file = request.files['bien_ban_ban']
            if file and file.filename:
                detai.bien_ban_ban = save_uploaded_file(file, folder_name)
        
        if 'bien_ban_de_cuong' in request.files:
            file = request.files['bien_ban_de_cuong']
            if file and file.filename:
                detai.bien_ban_de_cuong = save_uploaded_file(file, folder_name)
        
        if 'bien_ban_nghiem_thu' in request.files:
            file = request.files['bien_ban_nghiem_thu']
            if file and file.filename:
                detai.bien_ban_nghiem_thu = save_uploaded_file(file, folder_name)
        
        if 'file_sau_sua' in request.files:
            file = request.files['file_sau_sua']
            if file and file.filename:
                detai.file_sau_sua = save_uploaded_file(file, folder_name)
        
        # ======== CẬP NHẬT THÀNH VIÊN BAN ĐỀ TÀI ========
        # Xóa toàn bộ thành viên cũ
        BanDeTai.query.filter_by(de_tai_id=detai.id).delete()
        
        # Thêm lại từ form
        ban_de_tai_text = request.form.get('ban_de_tai', '').strip()
        if ban_de_tai_text:
            for line in ban_de_tai_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    ban_member = BanDeTai(
                        de_tai_id=detai.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(ban_member)
        
        # ======== CẬP NHẬT HỘI ĐỒNG DUYỆT ĐỀ CƯƠNG ========
        # Xóa toàn bộ thành viên cũ
        HoiDongDeCuongDeTai.query.filter_by(de_tai_id=detai.id).delete()
        
        # Thêm lại từ form
        hoi_dong_de_cuong_text = request.form.get('hoi_dong_de_cuong', '').strip()
        if hoi_dong_de_cuong_text:
            for line in hoi_dong_de_cuong_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    hd_member = HoiDongDeCuongDeTai(
                        de_tai_id=detai.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(hd_member)
        
        # ======== CẬP NHẬT HỘI ĐỒNG NGHIỆM THU ========
        # Xóa toàn bộ thành viên cũ
        HoiDongNghiemThuDeTai.query.filter_by(de_tai_id=detai.id).delete()
        
        # Thêm lại từ form
        hoi_dong_nghiem_thu_text = request.form.get('hoi_dong_nghiem_thu', '').strip()
        if hoi_dong_nghiem_thu_text:
            for line in hoi_dong_nghiem_thu_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    hd_member = HoiDongNghiemThuDeTai(
                        de_tai_id=detai.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(hd_member)
        
        db.session.commit()
        
        flash('Cập nhật đề tài thành công!', 'success')
        return redirect(url_for('main.detai_list'))
    
    return render_template('detai/edit.html', detai=detai, don_vi_list=DON_VI_LIST, cap_do_list=CAP_DO_LIST)

@main_bp.route('/detai/delete/<int:id>')
@login_required
def detai_delete(id):
    """Xóa đề tài - Chỉ xóa được đề tài của đơn vị mình hoặc admin"""
    detai = DeTai.query.get_or_404(id)
    
    # Kiểm tra quyền
    if current_user.role != 'admin' and detai.don_vi != current_user.don_vi:
        flash('Bạn không có quyền xóa đề tài của đơn vị khác!', 'danger')
        return redirect(url_for('main.detai_list'))
    
    db.session.delete(detai)
    db.session.commit()
    
    flash('Xóa đề tài thành công!', 'success')
    return redirect(url_for('main.detai_list'))

# ==================== QUẢN LÝ TÀI LIỆU ====================
@main_bp.route('/tailieu')
@login_required
def tailieu_list():
    """Danh sách tài liệu"""
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '', type=str)
    loai = request.args.get('loai', '', type=str)
    
    query = TaiLieu.query
    if search:
        query = query.filter(
            db.or_(
                TaiLieu.ten_tai_lieu.contains(search),
                TaiLieu.ma_tai_lieu.contains(search),
                TaiLieu.tac_gia.contains(search),  # Đổi từ chu_nhiem sang tac_gia
                TaiLieu.don_vi.contains(search)
            )
        )
    
    if loai:
        query = query.filter(TaiLieu.loai == loai)
    
    pagination = query.order_by(TaiLieu.created_at.desc()).paginate(
        page=page, per_page=10, error_out=False
    )
    
    return render_template('tailieu/list.html', pagination=pagination, search=search, loai=loai, loai_list=LOAI_TAI_LIEU)

# Route tương thích với tên cũ
@main_bp.route('/giaotrinh')
@login_required
def giaotrinh_list():
    """Chuyển hướng sang tài liệu"""
    return redirect(url_for('main.tailieu_list'))

@main_bp.route('/tailieu/add', methods=['GET', 'POST'])
@login_required
def tailieu_add():
    """Thêm tài liệu mới"""
    if request.method == 'POST':
        don_vi = current_user.don_vi if current_user.role != 'admin' else request.form.get('don_vi')
        
        tailieu = TaiLieu(
            ma_tai_lieu=request.form.get('ma_tai_lieu'),
            ten_tai_lieu=request.form.get('ten_tai_lieu'),
            loai=request.form.get('loai'),
            tac_gia=request.form.get('tac_gia'),  # Textarea nhiều dòng
            nha_xuat_ban=request.form.get('nha_xuat_ban'),
            nam_xuat_ban=int(request.form.get('nam_xuat_ban', 0)) if request.form.get('nam_xuat_ban') else None,
            mon_hoc=request.form.get('mon_hoc'),
            mo_ta=request.form.get('mo_ta'),
            linh_vuc=request.form.get('linh_vuc'),
            don_vi=don_vi,
            trang_thai=request.form.get('trang_thai', 'Đang thực hiện'),
            created_by=current_user.id
        )
        
        # Xử lý ngày tháng
        ngay_bat_dau = request.form.get('ngay_bat_dau')
        ngay_ket_thuc = request.form.get('ngay_ket_thuc')
        if ngay_bat_dau:
            tailieu.ngay_bat_dau = datetime.strptime(ngay_bat_dau, '%Y-%m-%d').date()
        if ngay_ket_thuc:
            tailieu.ngay_ket_thuc = datetime.strptime(ngay_ket_thuc, '%Y-%m-%d').date()
        
        db.session.add(tailieu)
        db.session.flush()  # Để lấy ID
        
        # Tạo thư mục riêng cho tài liệu: tailieu_<id>_<ma_tai_lieu>
        folder_name = f"tailieu_{tailieu.id}_{tailieu.ma_tai_lieu}"
        
        # Xử lý upload file tài liệu (PDF hoặc Word)
        if 'file_tai_lieu' in request.files:
            file = request.files['file_tai_lieu']
            if file and file.filename:
                tailieu.file_tai_lieu = save_uploaded_file(file, folder_name)
        
        # Xử lý upload các file
        if 'bien_ban_ban' in request.files:
            file = request.files['bien_ban_ban']
            if file and file.filename:
                tailieu.bien_ban_ban = save_uploaded_file(file, folder_name)
        
        if 'bien_ban_de_cuong' in request.files:
            file = request.files['bien_ban_de_cuong']
            if file and file.filename:
                tailieu.bien_ban_de_cuong = save_uploaded_file(file, folder_name)
        
        if 'bien_ban_nghiem_thu' in request.files:
            file = request.files['bien_ban_nghiem_thu']
            if file and file.filename:
                tailieu.bien_ban_nghiem_thu = save_uploaded_file(file, folder_name)
        
        if 'file_sau_sua' in request.files:
            file = request.files['file_sau_sua']
            if file and file.filename:
                tailieu.file_sau_sua = save_uploaded_file(file, folder_name)
        
        # Xử lý thêm thành viên Ban tài liệu
        ban_tai_lieu_text = request.form.get('ban_tai_lieu', '').strip()
        if ban_tai_lieu_text:
            for line in ban_tai_lieu_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    ban_member = BanTaiLieu(
                        tai_lieu_id=tailieu.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(ban_member)
        
        # Xử lý thêm thành viên Hội đồng duyệt đề cương
        hoi_dong_de_cuong_text = request.form.get('hoi_dong_de_cuong', '').strip()
        if hoi_dong_de_cuong_text:
            for line in hoi_dong_de_cuong_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    hd_member = HoiDongDeCuongTaiLieu(
                        tai_lieu_id=tailieu.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(hd_member)
        
        # Xử lý thêm thành viên Hội đồng nghiệm thu
        hoi_dong_nghiem_thu_text = request.form.get('hoi_dong_nghiem_thu', '').strip()
        if hoi_dong_nghiem_thu_text:
            for line in hoi_dong_nghiem_thu_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    hd_member = HoiDongNghiemThuTaiLieu(
                        tai_lieu_id=tailieu.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(hd_member)
        
        db.session.commit()
        
        flash('Thêm tài liệu thành công!', 'success')
        return redirect(url_for('main.tailieu_list'))
    
    return render_template('tailieu/add.html', don_vi_list=DON_VI_LIST, loai_list=LOAI_TAI_LIEU)

# Route tương thích
@main_bp.route('/giaotrinh/add', methods=['GET', 'POST'])
@login_required
def giaotrinh_add():
    """Chuyển hướng sang thêm tài liệu"""
    if request.method == 'POST':
        return tailieu_add()
    return redirect(url_for('main.tailieu_add'))

@main_bp.route('/tailieu/edit/<int:id>', methods=['GET', 'POST'])
@login_required
def tailieu_edit(id):
    """Sửa tài liệu"""
    tailieu = TaiLieu.query.get_or_404(id)
    
    if current_user.role != 'admin' and tailieu.don_vi != current_user.don_vi:
        flash('Bạn không có quyền chỉnh sửa tài liệu của đơn vị khác!', 'danger')
        return redirect(url_for('main.tailieu_list'))
    
    if request.method == 'POST':
        tailieu.ma_tai_lieu = request.form.get('ma_tai_lieu')
        tailieu.ten_tai_lieu = request.form.get('ten_tai_lieu')
        tailieu.loai = request.form.get('loai')
        tailieu.tac_gia = request.form.get('tac_gia')  # Đổi từ chu_nhiem sang tac_gia
        tailieu.nha_xuat_ban = request.form.get('nha_xuat_ban')
        tailieu.nam_xuat_ban = int(request.form.get('nam_xuat_ban', 0)) if request.form.get('nam_xuat_ban') else None
        tailieu.mon_hoc = request.form.get('mon_hoc')
        tailieu.mo_ta = request.form.get('mo_ta')
        tailieu.linh_vuc = request.form.get('linh_vuc')
        tailieu.trang_thai = request.form.get('trang_thai')
        tailieu.ket_qua = request.form.get('ket_qua')
        
        if current_user.role == 'admin':
            tailieu.don_vi = request.form.get('don_vi')
        
        ngay_bat_dau = request.form.get('ngay_bat_dau')
        ngay_ket_thuc = request.form.get('ngay_ket_thuc')
        if ngay_bat_dau:
            tailieu.ngay_bat_dau = datetime.strptime(ngay_bat_dau, '%Y-%m-%d').date()
        if ngay_ket_thuc:
            tailieu.ngay_ket_thuc = datetime.strptime(ngay_ket_thuc, '%Y-%m-%d').date()
        
        # Tạo thư mục riêng cho tài liệu
        folder_name = f"tailieu_{tailieu.id}_{tailieu.ma_tai_lieu}"
        
        # Xử lý upload file tài liệu (PDF hoặc Word)
        if 'file_tai_lieu' in request.files:
            file = request.files['file_tai_lieu']
            if file and file.filename:
                tailieu.file_tai_lieu = save_uploaded_file(file, folder_name)
        
        # Xử lý upload các file (chỉ upload nếu có file mới)
        if 'bien_ban_ban' in request.files:
            file = request.files['bien_ban_ban']
            if file and file.filename:
                tailieu.bien_ban_ban = save_uploaded_file(file, folder_name)
        
        if 'bien_ban_de_cuong' in request.files:
            file = request.files['bien_ban_de_cuong']
            if file and file.filename:
                tailieu.bien_ban_de_cuong = save_uploaded_file(file, folder_name)
        
        if 'bien_ban_nghiem_thu' in request.files:
            file = request.files['bien_ban_nghiem_thu']
            if file and file.filename:
                tailieu.bien_ban_nghiem_thu = save_uploaded_file(file, folder_name)
        
        if 'file_sau_sua' in request.files:
            file = request.files['file_sau_sua']
            if file and file.filename:
                tailieu.file_sau_sua = save_uploaded_file(file, folder_name)
        
        # ======== CẬP NHẬT THÀNH VIÊN BAN TÀI LIỆU ========
        # Xóa toàn bộ thành viên cũ
        BanTaiLieu.query.filter_by(tai_lieu_id=tailieu.id).delete()
        
        # Thêm lại từ form
        ban_tai_lieu_text = request.form.get('ban_tai_lieu', '').strip()
        if ban_tai_lieu_text:
            for line in ban_tai_lieu_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    ban_member = BanTaiLieu(
                        tai_lieu_id=tailieu.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(ban_member)
        
        # ======== CẬP NHẬT HỘI ĐỒNG DUYỆT ĐỀ CƯƠNG ========
        # Xóa toàn bộ thành viên cũ
        HoiDongDeCuongTaiLieu.query.filter_by(tai_lieu_id=tailieu.id).delete()
        
        # Thêm lại từ form
        hoi_dong_de_cuong_text = request.form.get('hoi_dong_de_cuong', '').strip()
        if hoi_dong_de_cuong_text:
            for line in hoi_dong_de_cuong_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    hd_member = HoiDongDeCuongTaiLieu(
                        tai_lieu_id=tailieu.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(hd_member)
        
        # ======== CẬP NHẬT HỘI ĐỒNG NGHIỆM THU ========
        # Xóa toàn bộ thành viên cũ
        HoiDongNghiemThuTaiLieu.query.filter_by(tai_lieu_id=tailieu.id).delete()
        
        # Thêm lại từ form
        hoi_dong_nghiem_thu_text = request.form.get('hoi_dong_nghiem_thu', '').strip()
        if hoi_dong_nghiem_thu_text:
            for line in hoi_dong_nghiem_thu_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    hd_member = HoiDongNghiemThuTaiLieu(
                        tai_lieu_id=tailieu.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(hd_member)
        
        db.session.commit()
        
        flash('Cập nhật tài liệu thành công!', 'success')
        return redirect(url_for('main.tailieu_list'))
    
    return render_template('tailieu/edit.html', tailieu=tailieu, don_vi_list=DON_VI_LIST, loai_list=LOAI_TAI_LIEU)

# Route tương thích
@main_bp.route('/giaotrinh/edit/<int:id>', methods=['GET', 'POST'])
@login_required
def giaotrinh_edit(id):
    """Chuyển hướng sang sửa tài liệu"""
    if request.method == 'POST':
        return tailieu_edit(id)
    return redirect(url_for('main.tailieu_edit', id=id))

@main_bp.route('/tailieu/delete/<int:id>')
@login_required
def tailieu_delete(id):
    """Xóa tài liệu"""
    tailieu = TaiLieu.query.get_or_404(id)
    
    if current_user.role != 'admin' and tailieu.don_vi != current_user.don_vi:
        flash('Bạn không có quyền xóa tài liệu của đơn vị khác!', 'danger')
        return redirect(url_for('main.tailieu_list'))
    
    db.session.delete(tailieu)
    db.session.commit()
    
    flash('Xóa tài liệu thành công!', 'success')
    return redirect(url_for('main.tailieu_list'))

# Route tương thích
@main_bp.route('/giaotrinh/delete/<int:id>')
@login_required
def giaotrinh_delete(id):
    """Chuyển hướng sang xóa tài liệu"""
    return tailieu_delete(id)

# ==================== QUẢN LÝ TIỀM LỰC ====================
@main_bp.route('/tiemluc')
@login_required
def tiemluc_list():
    """Danh sách tiềm lực khoa học"""
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '', type=str)
    
    query = TiemLuc.query
    if search:
        query = query.filter(
            db.or_(
                TiemLuc.ho_ten.contains(search),
                TiemLuc.ma_can_bo.contains(search),
                TiemLuc.don_vi.contains(search)
            )
        )
    
    pagination = query.order_by(TiemLuc.created_at.desc()).paginate(
        page=page, per_page=10, error_out=False
    )
    
    return render_template('tiemluc/list.html', pagination=pagination, search=search)

@main_bp.route('/tiemluc/add', methods=['GET', 'POST'])
@login_required
def tiemluc_add():
    """Thêm tiềm lực mới - Tự động gán đơn vị của user"""
    if request.method == 'POST':
        # Chỉ cho phép thêm với đơn vị của mình (nếu không phải admin)
        don_vi = current_user.don_vi if current_user.role != 'admin' else request.form.get('don_vi')
        
        tiemluc = TiemLuc(
            ma_can_bo=request.form.get('ma_can_bo'),
            ho_ten=request.form.get('ho_ten'),
            gioi_tinh=request.form.get('gioi_tinh'),
            don_vi=don_vi,
            chuc_vu=request.form.get('chuc_vu'),
            hoc_vi=request.form.get('hoc_vi'),
            hoc_ham=request.form.get('hoc_ham'),
            chuyen_mon=request.form.get('chuyen_mon'),
            email=request.form.get('email'),
            dien_thoai=request.form.get('dien_thoai'),
            kinh_nghiem=request.form.get('kinh_nghiem'),
            thanh_tich=request.form.get('thanh_tich'),
            created_by=current_user.id
        )
        
        ngay_sinh = request.form.get('ngay_sinh')
        if ngay_sinh:
            tiemluc.ngay_sinh = datetime.strptime(ngay_sinh, '%Y-%m-%d').date()
        
        db.session.add(tiemluc)
        db.session.commit()
        
        flash('Thêm tiềm lực khoa học thành công!', 'success')
        return redirect(url_for('main.tiemluc_list'))
    
    return render_template('tiemluc/add.html', don_vi_list=DON_VI_LIST)

@main_bp.route('/tiemluc/edit/<int:id>', methods=['GET', 'POST'])
@login_required
def tiemluc_edit(id):
    """Sửa tiềm lực - Chỉ sửa được cán bộ của đơn vị mình hoặc admin"""
    tiemluc = TiemLuc.query.get_or_404(id)
    
    # Kiểm tra quyền: chỉ admin hoặc user cùng đơn vị mới sửa được
    if current_user.role != 'admin' and tiemluc.don_vi != current_user.don_vi:
        flash('Bạn không có quyền chỉnh sửa cán bộ của đơn vị khác!', 'danger')
        return redirect(url_for('main.tiemluc_list'))
    
    if request.method == 'POST':
        tiemluc.ma_can_bo = request.form.get('ma_can_bo')
        tiemluc.ho_ten = request.form.get('ho_ten')
        tiemluc.gioi_tinh = request.form.get('gioi_tinh')
        
        # Admin có thể đổi đơn vị
        if current_user.role == 'admin':
            tiemluc.don_vi = request.form.get('don_vi')
        
        tiemluc.chuc_vu = request.form.get('chuc_vu')
        tiemluc.hoc_vi = request.form.get('hoc_vi')
        tiemluc.hoc_ham = request.form.get('hoc_ham')
        tiemluc.chuyen_mon = request.form.get('chuyen_mon')
        tiemluc.email = request.form.get('email')
        tiemluc.dien_thoai = request.form.get('dien_thoai')
        tiemluc.kinh_nghiem = request.form.get('kinh_nghiem')
        tiemluc.thanh_tich = request.form.get('thanh_tich')
        
        ngay_sinh = request.form.get('ngay_sinh')
        if ngay_sinh:
            tiemluc.ngay_sinh = datetime.strptime(ngay_sinh, '%Y-%m-%d').date()
        
        db.session.commit()
        
        flash('Cập nhật tiềm lực khoa học thành công!', 'success')
        return redirect(url_for('main.tiemluc_list'))
    
    return render_template('tiemluc/edit.html', tiemluc=tiemluc, don_vi_list=DON_VI_LIST)

@main_bp.route('/tiemluc/delete/<int:id>')
@login_required
def tiemluc_delete(id):
    """Xóa tiềm lực - Chỉ xóa được cán bộ của đơn vị mình hoặc admin"""
    tiemluc = TiemLuc.query.get_or_404(id)
    
    # Kiểm tra quyền
    if current_user.role != 'admin' and tiemluc.don_vi != current_user.don_vi:
        flash('Bạn không có quyền xóa cán bộ của đơn vị khác!', 'danger')
        return redirect(url_for('main.tiemluc_list'))
    
    db.session.delete(tiemluc)
    db.session.commit()
    
    flash('Xóa tiềm lực khoa học thành công!', 'success')
    return redirect(url_for('main.tiemluc_list'))

# ==================== QUẢN LÝ SÁNG KIẾN ====================
@main_bp.route('/sangkien')
@login_required
def sangkien_list():
    """Danh sách sáng kiến"""
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '', type=str)
    
    query = SangKien.query
    if search:
        query = query.filter(
            db.or_(
                SangKien.ten_sang_kien.contains(search),
                SangKien.ma_sang_kien.contains(search),
                SangKien.chu_nhiem.contains(search)
            )
        )
    
    pagination = query.order_by(SangKien.created_at.desc()).paginate(
        page=page, per_page=10, error_out=False
    )
    
    return render_template('sangkien/list.html', pagination=pagination, search=search)

@main_bp.route('/sangkien/add', methods=['GET', 'POST'])
@login_required
def sangkien_add():
    """Thêm sáng kiến mới - Tự động gán đơn vị của user"""
    if request.method == 'POST':
        # Kiểm tra mã sáng kiến đã tồn tại chưa
        ma_sang_kien = request.form.get('ma_sang_kien')
        existing = SangKien.query.filter_by(ma_sang_kien=ma_sang_kien).first()
        if existing:
            flash(f'Mã sáng kiến "{ma_sang_kien}" đã tồn tại! Vui lòng sử dụng mã khác.', 'danger')
            return render_template('sangkien/add.html', don_vi_list=DON_VI_LIST, form_data=request.form)
        
        # Chỉ cho phép thêm với đơn vị của mình (nếu không phải admin)
        don_vi = current_user.don_vi if current_user.role != 'admin' else request.form.get('don_vi')
        
        sangkien = SangKien(
            ma_sang_kien=ma_sang_kien,
            ten_sang_kien=request.form.get('ten_sang_kien'),
            mo_ta=request.form.get('mo_ta'),
            linh_vuc=request.form.get('linh_vuc'),
            chu_nhiem=request.form.get('chu_nhiem'),
            don_vi=don_vi,
            trang_thai=request.form.get('trang_thai', 'Đang thực hiện'),
            ket_qua=request.form.get('ket_qua'),
            created_by=current_user.id
        )
        
        # Xử lý ngày tháng
        ngay_bat_dau = request.form.get('ngay_bat_dau')
        ngay_ket_thuc = request.form.get('ngay_ket_thuc')
        if ngay_bat_dau:
            sangkien.ngay_bat_dau = datetime.strptime(ngay_bat_dau, '%Y-%m-%d').date()
        if ngay_ket_thuc:
            sangkien.ngay_ket_thuc = datetime.strptime(ngay_ket_thuc, '%Y-%m-%d').date()
        
        db.session.add(sangkien)
        db.session.flush()  # Để lấy ID
        
        # Tạo thư mục riêng cho sáng kiến: sangkien_<id>_<ma_sang_kien>
        folder_name = f"sangkien_{sangkien.id}_{sangkien.ma_sang_kien}"
        
        # Xử lý upload các file
        if 'bien_ban_ban' in request.files:
            file = request.files['bien_ban_ban']
            if file and file.filename:
                sangkien.bien_ban_ban = save_uploaded_file(file, folder_name)
        
        if 'bien_ban_de_cuong' in request.files:
            file = request.files['bien_ban_de_cuong']
            if file and file.filename:
                sangkien.bien_ban_de_cuong = save_uploaded_file(file, folder_name)
        
        if 'bien_ban_nghiem_thu' in request.files:
            file = request.files['bien_ban_nghiem_thu']
            if file and file.filename:
                sangkien.bien_ban_nghiem_thu = save_uploaded_file(file, folder_name)
        
        if 'file_sau_sua' in request.files:
            file = request.files['file_sau_sua']
            if file and file.filename:
                sangkien.file_sau_sua = save_uploaded_file(file, folder_name)
        
        # Xử lý thêm thành viên Ban sáng kiến
        ban_sang_kien_text = request.form.get('ban_sang_kien', '').strip()
        if ban_sang_kien_text:
            for line in ban_sang_kien_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    ban_member = BanSangKien(
                        sang_kien_id=sangkien.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(ban_member)
        
        # Xử lý thêm thành viên Hội đồng duyệt đề cương
        hoi_dong_de_cuong_text = request.form.get('hoi_dong_de_cuong', '').strip()
        if hoi_dong_de_cuong_text:
            for line in hoi_dong_de_cuong_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    hd_member = HoiDongDeCuongSangKien(
                        sang_kien_id=sangkien.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(hd_member)
        
        # Xử lý thêm thành viên Hội đồng nghiệm thu
        hoi_dong_nghiem_thu_text = request.form.get('hoi_dong_nghiem_thu', '').strip()
        if hoi_dong_nghiem_thu_text:
            for line in hoi_dong_nghiem_thu_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    hd_member = HoiDongNghiemThuSangKien(
                        sang_kien_id=sangkien.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(hd_member)
        
        db.session.commit()
        
        flash('Thêm sáng kiến thành công!', 'success')
        return redirect(url_for('main.sangkien_list'))
    
    return render_template('sangkien/add.html', don_vi_list=DON_VI_LIST, cap_do_list=CAP_DO_LIST)

@main_bp.route('/sangkien/edit/<int:id>', methods=['GET', 'POST'])
@login_required
def sangkien_edit(id):
    """Sửa sáng kiến - Chỉ sửa được sáng kiến của đơn vị mình hoặc admin"""
    sangkien = SangKien.query.get_or_404(id)
    
    # Kiểm tra quyền: chỉ admin hoặc user cùng đơn vị mới sửa được
    if current_user.role != 'admin' and sangkien.don_vi != current_user.don_vi:
        flash('Bạn không có quyền chỉnh sửa sáng kiến của đơn vị khác!', 'danger')
        return redirect(url_for('main.sangkien_list'))
    
    if request.method == 'POST':
        sangkien.ma_sang_kien = request.form.get('ma_sang_kien')
        sangkien.ten_sang_kien = request.form.get('ten_sang_kien')
        sangkien.mo_ta = request.form.get('mo_ta')
        sangkien.linh_vuc = request.form.get('linh_vuc')
        sangkien.chu_nhiem = request.form.get('chu_nhiem')
        sangkien.cap_do = request.form.get('cap_do')
        
        # Admin có thể đổi đơn vị
        if current_user.role == 'admin':
            sangkien.don_vi = request.form.get('don_vi')
        
        sangkien.trang_thai = request.form.get('trang_thai')
        sangkien.ket_qua = request.form.get('ket_qua')
        
        # Xử lý ngày tháng
        ngay_bat_dau = request.form.get('ngay_bat_dau')
        ngay_ket_thuc = request.form.get('ngay_ket_thuc')
        if ngay_bat_dau:
            sangkien.ngay_bat_dau = datetime.strptime(ngay_bat_dau, '%Y-%m-%d').date()
        if ngay_ket_thuc:
            sangkien.ngay_ket_thuc = datetime.strptime(ngay_ket_thuc, '%Y-%m-%d').date()
        
        # Tạo thư mục riêng cho sáng kiến
        folder_name = f"sangkien_{sangkien.id}_{sangkien.ma_sang_kien}"
        
        # Xử lý upload các file (chỉ upload nếu có file mới)
        if 'bien_ban_ban' in request.files:
            file = request.files['bien_ban_ban']
            if file and file.filename:
                sangkien.bien_ban_ban = save_uploaded_file(file, folder_name)
        
        if 'bien_ban_de_cuong' in request.files:
            file = request.files['bien_ban_de_cuong']
            if file and file.filename:
                sangkien.bien_ban_de_cuong = save_uploaded_file(file, folder_name)
        
        if 'bien_ban_nghiem_thu' in request.files:
            file = request.files['bien_ban_nghiem_thu']
            if file and file.filename:
                sangkien.bien_ban_nghiem_thu = save_uploaded_file(file, folder_name)
        
        if 'file_sau_sua' in request.files:
            file = request.files['file_sau_sua']
            if file and file.filename:
                sangkien.file_sau_sua = save_uploaded_file(file, folder_name)
        
        # ======== CẬP NHẬT THÀNH VIÊN BAN SÁNG KIẾN ========
        # Xóa toàn bộ thành viên cũ
        BanSangKien.query.filter_by(sang_kien_id=sangkien.id).delete()
        
        # Thêm lại từ form
        ban_sang_kien_text = request.form.get('ban_sang_kien', '').strip()
        if ban_sang_kien_text:
            for line in ban_sang_kien_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    ban_member = BanSangKien(
                        sang_kien_id=sangkien.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(ban_member)
        
        # ======== CẬP NHẬT HỘI ĐỒNG DUYỆT ĐỀ CƯƠNG ========
        # Xóa toàn bộ thành viên cũ
        HoiDongDeCuongSangKien.query.filter_by(sang_kien_id=sangkien.id).delete()
        
        # Thêm lại từ form
        hoi_dong_de_cuong_text = request.form.get('hoi_dong_de_cuong', '').strip()
        if hoi_dong_de_cuong_text:
            for line in hoi_dong_de_cuong_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    hd_member = HoiDongDeCuongSangKien(
                        sang_kien_id=sangkien.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(hd_member)
        
        # ======== CẬP NHẬT HỘI ĐỒNG NGHIỆM THU ========
        # Xóa toàn bộ thành viên cũ
        HoiDongNghiemThuSangKien.query.filter_by(sang_kien_id=sangkien.id).delete()
        
        # Thêm lại từ form
        hoi_dong_nghiem_thu_text = request.form.get('hoi_dong_nghiem_thu', '').strip()
        if hoi_dong_nghiem_thu_text:
            for line in hoi_dong_nghiem_thu_text.split('\n'):
                line = line.strip()
                if line and '-' in line:
                    parts = line.split('-', 1)
                    vai_tro = parts[0].strip()
                    ho_ten = parts[1].strip()
                    hd_member = HoiDongNghiemThuSangKien(
                        sang_kien_id=sangkien.id,
                        vai_tro=vai_tro,
                        ho_ten=ho_ten
                    )
                    db.session.add(hd_member)
        
        db.session.commit()
        
        flash('Cập nhật sáng kiến thành công!', 'success')
        return redirect(url_for('main.sangkien_list'))
    
    return render_template('sangkien/edit.html', sangkien=sangkien, don_vi_list=DON_VI_LIST, cap_do_list=CAP_DO_LIST)

@main_bp.route('/sangkien/delete/<int:id>')
@login_required
def sangkien_delete(id):
    """Xóa sáng kiến - Chỉ xóa được sáng kiến của đơn vị mình hoặc admin"""
    sangkien = SangKien.query.get_or_404(id)
    
    # Kiểm tra quyền
    if current_user.role != 'admin' and sangkien.don_vi != current_user.don_vi:
        flash('Bạn không có quyền xóa sáng kiến của đơn vị khác!', 'danger')
        return redirect(url_for('main.sangkien_list'))
    
    db.session.delete(sangkien)
    db.session.commit()
    
    flash('Xóa sáng kiến thành công!', 'success')
    return redirect(url_for('main.sangkien_list'))

# ==================== THỐNG KÊ ====================
@main_bp.route('/thongke')
@login_required
def thongke():
    """Trang thống kê tổng hợp"""
    # Lấy tham số lọc
    don_vi_filter = request.args.get('don_vi', '')
    nam_filter = request.args.get('nam', '')
    loai_filter = request.args.get('loai', 'donvi')  # donvi hoặc canhan
    
    # Base query cho đề tài
    detai_query = DeTai.query
    sangkien_query = SangKien.query
    tailieu_query = TaiLieu.query
    tiemluc_query = TiemLuc.query
    
    # Áp dụng bộ lọc đơn vị
    if don_vi_filter:
        detai_query = detai_query.filter(DeTai.don_vi == don_vi_filter)
        sangkien_query = sangkien_query.filter(SangKien.don_vi == don_vi_filter)
        tailieu_query = tailieu_query.filter(TaiLieu.don_vi == don_vi_filter)
        tiemluc_query = tiemluc_query.filter(TiemLuc.don_vi == don_vi_filter)
    
    # Áp dụng bộ lọc năm
    if nam_filter:
        nam = int(nam_filter)
        detai_query = detai_query.filter(func.extract('year', DeTai.ngay_bat_dau) == nam)
        sangkien_query = sangkien_query.filter(func.extract('year', SangKien.ngay_bat_dau) == nam)
        tailieu_query = tailieu_query.filter(func.extract('year', TaiLieu.ngay_bat_dau) == nam)
    
    # Thống kê đề tài theo đơn vị
    detai_donvi = db.session.query(
        DeTai.don_vi,
        func.count(DeTai.id).label('total')
    )
    if don_vi_filter:
        detai_donvi = detai_donvi.filter(DeTai.don_vi == don_vi_filter)
    if nam_filter:
        detai_donvi = detai_donvi.filter(func.extract('year', DeTai.ngay_bat_dau) == int(nam_filter))
    detai_donvi = detai_donvi.group_by(DeTai.don_vi).all()
    
    # Thống kê sáng kiến theo đơn vị
    sangkien_donvi = db.session.query(
        SangKien.don_vi,
        func.count(SangKien.id).label('total')
    )
    if don_vi_filter:
        sangkien_donvi = sangkien_donvi.filter(SangKien.don_vi == don_vi_filter)
    if nam_filter:
        sangkien_donvi = sangkien_donvi.filter(func.extract('year', SangKien.ngay_bat_dau) == int(nam_filter))
    sangkien_donvi = sangkien_donvi.group_by(SangKien.don_vi).all()
    
    # Thống kê tài liệu theo đơn vị
    tailieu_donvi = db.session.query(
        TaiLieu.don_vi,
        func.count(TaiLieu.id).label('total')
    )
    if don_vi_filter:
        tailieu_donvi = tailieu_donvi.filter(TaiLieu.don_vi == don_vi_filter)
    if nam_filter:
        tailieu_donvi = tailieu_donvi.filter(func.extract('year', TaiLieu.ngay_bat_dau) == int(nam_filter))
    tailieu_donvi = tailieu_donvi.group_by(TaiLieu.don_vi).all()
    
    # Thống kê tài liệu theo loại
    tailieu_loai = db.session.query(
        TaiLieu.loai,
        func.count(TaiLieu.id).label('total')
    )
    if don_vi_filter:
        tailieu_loai = tailieu_loai.filter(TaiLieu.don_vi == don_vi_filter)
    if nam_filter:
        tailieu_loai = tailieu_loai.filter(func.extract('year', TaiLieu.ngay_bat_dau) == int(nam_filter))
    tailieu_loai = tailieu_loai.group_by(TaiLieu.loai).all()
    
    # Thống kê tiềm lực theo đơn vị
    tiemluc_donvi = db.session.query(
        TiemLuc.don_vi,
        func.count(TiemLuc.id).label('total')
    )
    if don_vi_filter:
        tiemluc_donvi = tiemluc_donvi.filter(TiemLuc.don_vi == don_vi_filter)
    tiemluc_donvi = tiemluc_donvi.group_by(TiemLuc.don_vi).all()
    
    # Thống kê trạng thái đề tài
    detai_trangthai = db.session.query(
        DeTai.trang_thai,
        func.count(DeTai.id).label('total')
    )
    if don_vi_filter:
        detai_trangthai = detai_trangthai.filter(DeTai.don_vi == don_vi_filter)
    if nam_filter:
        detai_trangthai = detai_trangthai.filter(func.extract('year', DeTai.ngay_bat_dau) == int(nam_filter))
    detai_trangthai = detai_trangthai.group_by(DeTai.trang_thai).all()
    
    # Thống kê theo cá nhân (chủ nhiệm)
    detai_canhan = []
    sangkien_canhan = []
    tailieu_canhan = []
    
    if loai_filter == 'canhan':
        # Top cá nhân có nhiều đề tài nhất
        detai_canhan_query = db.session.query(
            DeTai.chu_nhiem,
            DeTai.don_vi,
            func.count(DeTai.id).label('total')
        )
        if don_vi_filter:
            detai_canhan_query = detai_canhan_query.filter(DeTai.don_vi == don_vi_filter)
        if nam_filter:
            detai_canhan_query = detai_canhan_query.filter(func.extract('year', DeTai.ngay_bat_dau) == int(nam_filter))
        detai_canhan = detai_canhan_query.group_by(DeTai.chu_nhiem, DeTai.don_vi).order_by(func.count(DeTai.id).desc()).limit(20).all()
        
        # Top cá nhân có nhiều sáng kiến nhất
        sangkien_canhan_query = db.session.query(
            SangKien.chu_nhiem,
            SangKien.don_vi,
            func.count(SangKien.id).label('total')
        )
        if don_vi_filter:
            sangkien_canhan_query = sangkien_canhan_query.filter(SangKien.don_vi == don_vi_filter)
        if nam_filter:
            sangkien_canhan_query = sangkien_canhan_query.filter(func.extract('year', SangKien.ngay_bat_dau) == int(nam_filter))
        sangkien_canhan = sangkien_canhan_query.group_by(SangKien.chu_nhiem, SangKien.don_vi).order_by(func.count(SangKien.id).desc()).limit(20).all()
        
        # Top cá nhân có nhiều tài liệu nhất
        tailieu_canhan_query = db.session.query(
            TaiLieu.tac_gia,
            TaiLieu.don_vi,
            func.count(TaiLieu.id).label('total')
        )
        if don_vi_filter:
            tailieu_canhan_query = tailieu_canhan_query.filter(TaiLieu.don_vi == don_vi_filter)
        if nam_filter:
            tailieu_canhan_query = tailieu_canhan_query.filter(func.extract('year', TaiLieu.ngay_bat_dau) == int(nam_filter))
        tailieu_canhan = tailieu_canhan_query.group_by(TaiLieu.tac_gia, TaiLieu.don_vi).order_by(func.count(TaiLieu.id).desc()).limit(20).all()
    
    # Tổng số liệu
    total_detai = detai_query.count()
    total_sangkien = sangkien_query.count()
    total_tailieu = tailieu_query.count()
    total_tiemluc = tiemluc_query.count()
    
    return render_template('thongke.html',
                         detai_donvi=detai_donvi,
                         sangkien_donvi=sangkien_donvi,
                         tailieu_donvi=tailieu_donvi,
                         tailieu_loai=tailieu_loai,
                         tiemluc_donvi=tiemluc_donvi,
                         detai_trangthai=detai_trangthai,
                         detai_canhan=detai_canhan,
                         sangkien_canhan=sangkien_canhan,
                         tailieu_canhan=tailieu_canhan,
                         total_detai=total_detai,
                         total_sangkien=total_sangkien,
                         total_tailieu=total_tailieu,
                         total_tiemluc=total_tiemluc,
                         don_vi_list=DON_VI_LIST,
                         don_vi_filter=don_vi_filter,
                         nam_filter=nam_filter,
                         loai_filter=loai_filter)

@main_bp.route('/thongke/export-excel')
@login_required
def thongke_export_excel():
    """Xuất thống kê ra Excel"""
    # Lấy tham số lọc
    don_vi_filter = request.args.get('don_vi', '')
    nam_filter = request.args.get('nam', '')
    
    # Tạo workbook
    wb = Workbook()
    wb.remove(wb.active)  # Xóa sheet mặc định
    
    # Định dạng tiêu đề
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    # Sheet 1: Thống kê đề tài theo đơn vị
    ws1 = wb.create_sheet("Đề tài theo đơn vị")
    ws1.append(["STT", "Đơn vị", "Số lượng"])
    
    # Định dạng header
    for cell in ws1[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
    
    query = db.session.query(DeTai.don_vi, func.count(DeTai.id).label('total'))
    if don_vi_filter:
        query = query.filter(DeTai.don_vi == don_vi_filter)
    if nam_filter:
        query = query.filter(func.extract('year', DeTai.ngay_bat_dau) == int(nam_filter))
    detai_stats = query.group_by(DeTai.don_vi).all()
    
    for idx, (don_vi, total) in enumerate(detai_stats, 1):
        ws1.append([idx, don_vi or "Chưa xác định", total])
    
    ws1.column_dimensions['A'].width = 10
    ws1.column_dimensions['B'].width = 40
    ws1.column_dimensions['C'].width = 15
    
    # Sheet 2: Thống kê sáng kiến theo đơn vị
    ws2 = wb.create_sheet("Sáng kiến theo đơn vị")
    ws2.append(["STT", "Đơn vị", "Số lượng"])
    
    for cell in ws2[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
    
    query = db.session.query(SangKien.don_vi, func.count(SangKien.id).label('total'))
    if don_vi_filter:
        query = query.filter(SangKien.don_vi == don_vi_filter)
    if nam_filter:
        query = query.filter(func.extract('year', SangKien.ngay_bat_dau) == int(nam_filter))
    sangkien_stats = query.group_by(SangKien.don_vi).all()
    
    for idx, (don_vi, total) in enumerate(sangkien_stats, 1):
        ws2.append([idx, don_vi or "Chưa xác định", total])
    
    ws2.column_dimensions['A'].width = 10
    ws2.column_dimensions['B'].width = 40
    ws2.column_dimensions['C'].width = 15
    
    # Sheet 3: Thống kê tài liệu theo đơn vị
    ws3 = wb.create_sheet("Tài liệu theo đơn vị")
    ws3.append(["STT", "Đơn vị", "Số lượng"])
    
    for cell in ws3[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
    
    query = db.session.query(TaiLieu.don_vi, func.count(TaiLieu.id).label('total'))
    if don_vi_filter:
        query = query.filter(TaiLieu.don_vi == don_vi_filter)
    if nam_filter:
        query = query.filter(func.extract('year', TaiLieu.ngay_bat_dau) == int(nam_filter))
    tailieu_stats = query.group_by(TaiLieu.don_vi).all()
    
    for idx, (don_vi, total) in enumerate(tailieu_stats, 1):
        ws3.append([idx, don_vi or "Chưa xác định", total])
    
    ws3.column_dimensions['A'].width = 10
    ws3.column_dimensions['B'].width = 40
    ws3.column_dimensions['C'].width = 15
    
    # Sheet 4: Thống kê tiềm lực theo đơn vị
    ws4 = wb.create_sheet("Tiềm lực theo đơn vị")
    ws4.append(["STT", "Đơn vị", "Số lượng"])
    
    for cell in ws4[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
    
    query = db.session.query(TiemLuc.don_vi, func.count(TiemLuc.id).label('total'))
    if don_vi_filter:
        query = query.filter(TiemLuc.don_vi == don_vi_filter)
    tiemluc_stats = query.group_by(TiemLuc.don_vi).all()
    
    for idx, (don_vi, total) in enumerate(tiemluc_stats, 1):
        ws4.append([idx, don_vi or "Chưa xác định", total])
    
    ws4.column_dimensions['A'].width = 10
    ws4.column_dimensions['B'].width = 40
    ws4.column_dimensions['C'].width = 15
    
    # Sheet 5: Tổng hợp
    ws5 = wb.create_sheet("Tổng hợp", 0)
    ws5.append(["THỐNG KÊ KHOA HỌC"])
    ws5.merge_cells('A1:C1')
    ws5['A1'].font = Font(bold=True, size=16)
    ws5['A1'].alignment = Alignment(horizontal="center")
    
    ws5.append([])
    ws5.append(["Loại", "Số lượng"])
    
    for cell in ws5[3]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
    
    total_detai = DeTai.query.count() if not don_vi_filter and not nam_filter else len(detai_stats)
    total_sangkien = SangKien.query.count() if not don_vi_filter and not nam_filter else len(sangkien_stats)
    total_tailieu = TaiLieu.query.count() if not don_vi_filter and not nam_filter else len(tailieu_stats)
    total_tiemluc = TiemLuc.query.count() if not don_vi_filter else len(tiemluc_stats)
    
    ws5.append(["Đề tài", sum(total for _, total in detai_stats)])
    ws5.append(["Sáng kiến", sum(total for _, total in sangkien_stats)])
    ws5.append(["Tài liệu", sum(total for _, total in tailieu_stats)])
    ws5.append(["Tiềm lực", sum(total for _, total in tiemluc_stats)])
    
    ws5.column_dimensions['A'].width = 30
    ws5.column_dimensions['B'].width = 20
    
    # Lưu file vào memory
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    # Tạo tên file
    filename = f"thongke_khoa_hoc"
    if don_vi_filter:
        filename += f"_{don_vi_filter}"
    if nam_filter:
        filename += f"_{nam_filter}"
    filename += ".xlsx"
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

@main_bp.route('/thongke/export-word')
@login_required  
def thongke_export_word():
    """Xuất thống kê ra Word"""
    from flask import send_file
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    import io
    
    # Lấy tham số lọc
    don_vi_filter = request.args.get('don_vi', '')
    nam_filter = request.args.get('nam', '')
    
    # Tạo document
    doc = Document()
    
    # Tiêu đề
    heading = doc.add_heading('BÁO CÁO THỐNG KÊ HOẠT ĐỘNG KHOA HỌC', 0)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Thông tin bộ lọc
    if don_vi_filter or nam_filter:
        filter_info = doc.add_paragraph()
        filter_info.add_run('Bộ lọc: ').bold = True
        if don_vi_filter:
            filter_info.add_run(f'Đơn vị: {don_vi_filter}  ')
        if nam_filter:
            filter_info.add_run(f'Năm: {nam_filter}')
        filter_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # 1. Thống kê đề tài
    doc.add_heading('I. THỐNG KÊ ĐỀ TÀI THEO ĐƠN VỊ', 1)
    query = db.session.query(DeTai.don_vi, func.count(DeTai.id).label('total'))
    if don_vi_filter:
        query = query.filter(DeTai.don_vi == don_vi_filter)
    if nam_filter:
        query = query.filter(func.extract('year', DeTai.ngay_bat_dau) == int(nam_filter))
    detai_stats = query.group_by(DeTai.don_vi).all()
    
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Light Grid Accent 1'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'STT'
    hdr_cells[1].text = 'Đơn vị'
    hdr_cells[2].text = 'Số lượng'
    
    for idx, (don_vi, total) in enumerate(detai_stats, 1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = don_vi or "Chưa xác định"
        row_cells[2].text = str(total)
    
    doc.add_paragraph()
    
    # 2. Thống kê sáng kiến
    doc.add_heading('II. THỐNG KÊ SÁNG KIẾN THEO ĐƠN VỊ', 1)
    query = db.session.query(SangKien.don_vi, func.count(SangKien.id).label('total'))
    if don_vi_filter:
        query = query.filter(SangKien.don_vi == don_vi_filter)
    if nam_filter:
        query = query.filter(func.extract('year', SangKien.ngay_bat_dau) == int(nam_filter))
    sangkien_stats = query.group_by(SangKien.don_vi).all()
    
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Light Grid Accent 1'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'STT'
    hdr_cells[1].text = 'Đơn vị'
    hdr_cells[2].text = 'Số lượng'
    
    for idx, (don_vi, total) in enumerate(sangkien_stats, 1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = don_vi or "Chưa xác định"
        row_cells[2].text = str(total)
    
    doc.add_paragraph()
    
    # 3. Thống kê tài liệu
    doc.add_heading('III. THỐNG KÊ TÀI LIỆU THEO LOẠI', 1)
    query = db.session.query(TaiLieu.loai, func.count(TaiLieu.id).label('total'))
    if don_vi_filter:
        query = query.filter(TaiLieu.don_vi == don_vi_filter)
    if nam_filter:
        query = query.filter(func.extract('year', TaiLieu.ngay_bat_dau) == int(nam_filter))
    tailieu_stats = query.group_by(TaiLieu.loai).all()
    
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Light Grid Accent 1'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'STT'
    hdr_cells[1].text = 'Loại tài liệu'
    hdr_cells[2].text = 'Số lượng'
    
    for idx, (loai, total) in enumerate(tailieu_stats, 1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = loai or "Chưa phân loại"
        row_cells[2].text = str(total)
    
    doc.add_paragraph()
    
    # 4. Thống kê tiềm lực
    doc.add_heading('IV. THỐNG KÊ TIỀM LỰC KHOA HỌC THEO ĐƠN VỊ', 1)
    query = db.session.query(TiemLuc.don_vi, func.count(TiemLuc.id).label('total'))
    if don_vi_filter:
        query = query.filter(TiemLuc.don_vi == don_vi_filter)
    tiemluc_stats = query.group_by(TiemLuc.don_vi).all()
    
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Light Grid Accent 1'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'STT'
    hdr_cells[1].text = 'Đơn vị'
    hdr_cells[2].text = 'Số lượng'
    
    for idx, (don_vi, total) in enumerate(tiemluc_stats, 1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = don_vi or "Chưa xác định"
        row_cells[2].text = str(total)
    
    # Lưu vào memory
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    
    # Tạo tên file
    filename = f"baocao_thongke_khoa_hoc"
    if don_vi_filter:
        filename += f"_{don_vi_filter.replace(' ', '_')}"
    if nam_filter:
        filename += f"_{nam_filter}"
    filename += ".docx"
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name=filename
    )

# ==================== TÌM KIẾM ====================
@main_bp.route('/timkiem', methods=['GET', 'POST'])
@login_required
def timkiem():
    """Trang tìm kiếm tổng hợp"""
    results = {
        'detai': [],
        'sangkien': [],
        'tailieu': [],
        'tiemluc': []
    }
    search_query = ''
    search_type = 'all'
    
    if request.method == 'GET' and request.args.get('q'):
        search_query = request.args.get('q', '').strip()
        search_type = request.args.get('type', 'all')
        
        if search_query:
            # Tìm kiếm đề tài
            if search_type in ['all', 'detai']:
                results['detai'] = DeTai.query.filter(
                    db.or_(
                        DeTai.ten_de_tai.contains(search_query),
                        DeTai.ma_de_tai.contains(search_query),
                        DeTai.chu_nhiem.contains(search_query),
                        DeTai.mo_ta.contains(search_query),
                     #   DeTai.ban_thanh_vien.contains(search_query),
                        DeTai.linh_vuc.contains(search_query),
                        DeTai.don_vi.contains(search_query),
                     #   DeTai.hoi_dong_nghiem_thu.contains(search_query)
                    )
                ).limit(20).all()
            
            # Tìm kiếm sáng kiến
            if search_type in ['all', 'sangkien']:
                results['sangkien'] = SangKien.query.filter(
                    db.or_(
                        SangKien.ten_sang_kien.contains(search_query),
                        SangKien.ma_sang_kien.contains(search_query),
                        SangKien.chu_nhiem.contains(search_query),
                        SangKien.mo_ta.contains(search_query),
                       # SangKien.ban_thanh_vien.contains(search_query),
                        SangKien.linh_vuc.contains(search_query),
                        SangKien.don_vi.contains(search_query),
                      #  SangKien.hoi_dong_nghiem_thu.contains(search_query)
                    )
                ).limit(20).all()
            
            # Tìm kiếm tài liệu
            if search_type in ['all', 'tailieu']:
                results['tailieu'] = TaiLieu.query.filter(
                    db.or_(
                        TaiLieu.ten_tai_lieu.contains(search_query),
                        TaiLieu.ma_tai_lieu.contains(search_query),
                        TaiLieu.tac_gia.contains(search_query),
                        TaiLieu.mo_ta.contains(search_query)
                    )
                ).limit(20).all()
            
            # Tìm kiếm tiềm lực
            if search_type in ['all', 'tiemluc']:
                results['tiemluc'] = TiemLuc.query.filter(
                    db.or_(
                        TiemLuc.ho_ten.contains(search_query),
                        TiemLuc.ma_can_bo.contains(search_query),
                        TiemLuc.chuyen_mon.contains(search_query),
                        TiemLuc.don_vi.contains(search_query)
                    )
                ).limit(20).all()
    
    return render_template('timkiem.html',
                         results=results,
                         search_query=search_query,
                         search_type=search_type)
