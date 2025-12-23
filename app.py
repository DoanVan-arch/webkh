from flask import Flask
from flask_login import LoginManager
from models import db, User
from routes.auth import auth_bp
from routes.main import main_bp
from config import Config
import os

def create_app():
    """Tạo và cấu hình Flask app"""
    app = Flask(__name__)
    app.config.from_object(Config)
    
    # Khởi tạo database
    db.init_app(app)
    
    # Khởi tạo Login Manager
    login_manager = LoginManager()
    login_manager.init_app(app)
    login_manager.login_view = 'auth.login'
    login_manager.login_message = 'Vui lòng đăng nhập để truy cập trang này.'
    login_manager.login_message_category = 'warning'
    
    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))
    
    # Đăng ký blueprints
    app.register_blueprint(auth_bp)
    app.register_blueprint(main_bp)
    
    # Tạo database và thư mục uploads
    with app.app_context():
        db.create_all()
        
        # Tạo thư mục uploads nếu chưa có
        upload_folder = app.config['UPLOAD_FOLDER']
        if not os.path.exists(upload_folder):
            os.makedirs(upload_folder)
        
        # Tạo tài khoản admin mặc định nếu chưa có
        if User.query.count() == 0:
            admin = User(
                username='admin',
                email='admin@military.edu.vn',
                full_name='Quản trị viên',
                role='admin',
                don_vi='Phòng Khoa học quân sự (B4)'
            )
            admin.set_password('admin123')
            db.session.add(admin)
            db.session.commit()
            print('Đã tạo tài khoản admin mặc định: admin/admin123')
    
    return app

if __name__ == '__main__':
    app = create_app()
    # Tắt use_reloader để tránh lỗi watchdog
    app.run(debug=True, host='0.0.0.0', port=5000, use_reloader=False)
