from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_required, current_user
from models import db, DeTai, GiaoTrinh, TiemLuc, User, DON_VI_LIST
from datetime import datetime
from sqlalchemy import func

main_bp = Blueprint('main', __name__)

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
    total_giaotrinh = GiaoTrinh.query.count()
    total_tiemluc = TiemLuc.query.count()
    total_users = User.query.count()
    
    # Đề tài gần đây
    recent_detai = DeTai.query.order_by(DeTai.created_at.desc()).limit(5).all()
    
    return render_template('dashboard.html',
                         total_detai=total_detai,
                         total_giaotrinh=total_giaotrinh,
                         total_tiemluc=total_tiemluc,
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
            kinh_phi=float(request.form.get('kinh_phi', 0)),
            trang_thai=request.form.get('trang_thai'),
            created_by=current_user.id
        )
        
        # Xử lý ngày tháng
        ngay_bat_dau = request.form.get('ngay_bat_dau')
        ngay_ket_thuc = request.form.get('ngay_ket_thuc')
        if ngay_bat_dau:
            detai.ngay_bat_dau = datetime.strptime(ngay_bat_dau, '%Y-%m-%d').date()
        if ngay_ket_thuc:
            detai.ngay_ket_thuc = datetime.strptime(ngay_ket_thuc, '%Y-%m-%d').date()
        
        db.session.add(detai)
        db.session.commit()
        
        flash('Thêm đề tài thành công!', 'success')
        return redirect(url_for('main.detai_list'))
    
    return render_template('detai/add.html', don_vi_list=DON_VI_LIST)

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
        
        # Admin có thể đổi đơn vị
        if current_user.role == 'admin':
            detai.don_vi = request.form.get('don_vi')
        
        detai.kinh_phi = float(request.form.get('kinh_phi', 0))
        detai.trang_thai = request.form.get('trang_thai')
        detai.ket_qua = request.form.get('ket_qua')
        
        ngay_bat_dau = request.form.get('ngay_bat_dau')
        ngay_ket_thuc = request.form.get('ngay_ket_thuc')
        if ngay_bat_dau:
            detai.ngay_bat_dau = datetime.strptime(ngay_bat_dau, '%Y-%m-%d').date()
        if ngay_ket_thuc:
            detai.ngay_ket_thuc = datetime.strptime(ngay_ket_thuc, '%Y-%m-%d').date()
        
        db.session.commit()
        
        flash('Cập nhật đề tài thành công!', 'success')
        return redirect(url_for('main.detai_list'))
    
    return render_template('detai/edit.html', detai=detai, don_vi_list=DON_VI_LIST)

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

# ==================== QUẢN LÝ GIÁO TRÌNH ====================
@main_bp.route('/giaotrinh')
@login_required
def giaotrinh_list():
    """Danh sách giáo trình - Xem được tất cả nhưng hiển thị đơn vị"""
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '', type=str)
    
    query = GiaoTrinh.query
    if search:
        query = query.filter(
            db.or_(
                GiaoTrinh.ten_tai_lieu.contains(search),
                GiaoTrinh.ma_tai_lieu.contains(search),
                GiaoTrinh.tac_gia.contains(search),
                GiaoTrinh.don_vi.contains(search)
            )
        )
    
    pagination = query.order_by(GiaoTrinh.created_at.desc()).paginate(
        page=page, per_page=10, error_out=False
    )
    
    return render_template('giaotrinh/list.html', pagination=pagination, search=search)

@main_bp.route('/giaotrinh/add', methods=['GET', 'POST'])
@login_required
def giaotrinh_add():
    """Thêm giáo trình mới - Tự động gán đơn vị của user"""
    if request.method == 'POST':
        # Chỉ cho phép thêm với đơn vị của mình (nếu không phải admin)
        don_vi = current_user.don_vi if current_user.role != 'admin' else request.form.get('don_vi')
        
        giaotrinh = GiaoTrinh(
            ma_tai_lieu=request.form.get('ma_tai_lieu'),
            ten_tai_lieu=request.form.get('ten_tai_lieu'),
            loai=request.form.get('loai'),
            tac_gia=request.form.get('tac_gia'),
            nha_xuat_ban=request.form.get('nha_xuat_ban'),
            nam_xuat_ban=int(request.form.get('nam_xuat_ban', 0)) if request.form.get('nam_xuat_ban') else None,
            mon_hoc=request.form.get('mon_hoc'),
            mo_ta=request.form.get('mo_ta'),
            don_vi=don_vi,
            created_by=current_user.id
        )
        
        db.session.add(giaotrinh)
        db.session.commit()
        
        flash('Thêm tài liệu thành công!', 'success')
        return redirect(url_for('main.giaotrinh_list'))
    
    return render_template('giaotrinh/add.html', don_vi_list=DON_VI_LIST)

@main_bp.route('/giaotrinh/edit/<int:id>', methods=['GET', 'POST'])
@login_required
def giaotrinh_edit(id):
    """Sửa giáo trình - Chỉ sửa được tài liệu của đơn vị mình hoặc admin"""
    giaotrinh = GiaoTrinh.query.get_or_404(id)
    
    # Kiểm tra quyền: chỉ admin hoặc user cùng đơn vị mới sửa được
    if current_user.role != 'admin' and giaotrinh.don_vi != current_user.don_vi:
        flash('Bạn không có quyền chỉnh sửa tài liệu của đơn vị khác!', 'danger')
        return redirect(url_for('main.giaotrinh_list'))
    
    if request.method == 'POST':
        giaotrinh.ma_tai_lieu = request.form.get('ma_tai_lieu')
        giaotrinh.ten_tai_lieu = request.form.get('ten_tai_lieu')
        giaotrinh.loai = request.form.get('loai')
        giaotrinh.tac_gia = request.form.get('tac_gia')
        giaotrinh.nha_xuat_ban = request.form.get('nha_xuat_ban')
        giaotrinh.nam_xuat_ban = int(request.form.get('nam_xuat_ban', 0)) if request.form.get('nam_xuat_ban') else None
        giaotrinh.mon_hoc = request.form.get('mon_hoc')
        giaotrinh.mo_ta = request.form.get('mo_ta')
        
        # Admin có thể đổi đơn vị
        if current_user.role == 'admin':
            giaotrinh.don_vi = request.form.get('don_vi')
        
        db.session.commit()
        
        flash('Cập nhật tài liệu thành công!', 'success')
        return redirect(url_for('main.giaotrinh_list'))
    
    return render_template('giaotrinh/edit.html', giaotrinh=giaotrinh, don_vi_list=DON_VI_LIST)

@main_bp.route('/giaotrinh/delete/<int:id>')
@login_required
def giaotrinh_delete(id):
    """Xóa giáo trình - Chỉ xóa được tài liệu của đơn vị mình hoặc admin"""
    giaotrinh = GiaoTrinh.query.get_or_404(id)
    
    # Kiểm tra quyền
    if current_user.role != 'admin' and giaotrinh.don_vi != current_user.don_vi:
        flash('Bạn không có quyền xóa tài liệu của đơn vị khác!', 'danger')
        return redirect(url_for('main.giaotrinh_list'))
    
    db.session.delete(giaotrinh)
    db.session.commit()
    
    flash('Xóa tài liệu thành công!', 'success')
    return redirect(url_for('main.giaotrinh_list'))

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
