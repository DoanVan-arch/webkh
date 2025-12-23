from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_user, logout_user, login_required, current_user
from models import db, User, DON_VI_LIST
from werkzeug.security import generate_password_hash

auth_bp = Blueprint('auth', __name__)

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    """Đăng nhập"""
    if current_user.is_authenticated:
        return redirect(url_for('main.dashboard'))
    
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        remember = request.form.get('remember', False)
        
        user = User.query.filter_by(username=username).first()
        
        if user and user.check_password(password):
            login_user(user, remember=remember)
            next_page = request.args.get('next')
            return redirect(next_page if next_page else url_for('main.dashboard'))
        else:
            flash('Tên đăng nhập hoặc mật khẩu không đúng!', 'danger')
    
    return render_template('login.html')

@auth_bp.route('/logout')
@login_required
def logout():
    """Đăng xuất"""
    logout_user()
    flash('Đã đăng xuất thành công!', 'success')
    return redirect(url_for('auth.login'))

@auth_bp.route('/register', methods=['GET', 'POST'])
def register():
    """Đăng ký tài khoản mới"""
    if current_user.is_authenticated:
        return redirect(url_for('main.dashboard'))
    
    if request.method == 'POST':
        username = request.form.get('username')
        email = request.form.get('email')
        full_name = request.form.get('full_name')
        don_vi = request.form.get('don_vi')
        password = request.form.get('password')
        confirm_password = request.form.get('confirm_password')
        
        if password != confirm_password:
            flash('Mật khẩu xác nhận không khớp!', 'danger')
            return render_template('register.html', don_vi_list=DON_VI_LIST)
        
        if User.query.filter_by(username=username).first():
            flash('Tên đăng nhập đã tồn tại!', 'danger')
            return render_template('register.html', don_vi_list=DON_VI_LIST)
        
        if User.query.filter_by(email=email).first():
            flash('Email đã được sử dụng!', 'danger')
            return render_template('register.html', don_vi_list=DON_VI_LIST)
        
        user = User(
            username=username,
            email=email,
            full_name=full_name,
            don_vi=don_vi,
            role='user'
        )
        user.set_password(password)
        
        db.session.add(user)
        db.session.commit()
        
        flash('Đăng ký thành công! Vui lòng đăng nhập.', 'success')
        return redirect(url_for('auth.login'))
    
    return render_template('register.html', don_vi_list=DON_VI_LIST)
