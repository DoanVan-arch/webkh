import os
from werkzeug.utils import secure_filename
from docx import Document
import re

ALLOWED_EXTENSIONS = {'doc', 'docx', 'pdf', 'xls', 'xlsx', 'jpg', 'jpeg', 'png', 'zip', 'rar'}
UPLOAD_FOLDER = 'static/uploads'

# Định nghĩa các thư mục riêng cho từng loại
UPLOAD_FOLDERS = {
    'detai': 'static/uploads/detai',
    'tailieu': 'static/uploads/tailieu',
    'sangkien': 'static/uploads/sangkien',
    'tiemluc': 'static/uploads/tiemluc',
    'bien_ban': 'static/uploads/bien_ban',
    'other': 'static/uploads/other'
}

def allowed_file(filename):
    """Kiểm tra file có được phép upload không"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_upload_folder(category='other'):
    """
    Lấy đường dẫn thư mục upload theo loại
    
    Args:
        category: Loại file (detai, tailieu, sangkien, tiemluc, bien_ban)
    
    Returns:
        str: Đường dẫn thư mục upload
    """
    return UPLOAD_FOLDERS.get(category, UPLOAD_FOLDERS['other'])

def save_uploaded_file(file, prefix='', category='other'):
    """
    Lưu file upload vào thư mục riêng theo loại
    
    Args:
        file: File object từ request.files
        prefix: Tiền tố cho tên file (vd: 'detai_', 'bienban_')
        category: Loại file để xác định thư mục lưu (detai, tailieu, sangkien, tiemluc, bien_ban)
    
    Returns:
        str: relative path của file (để lưu vào database) hoặc None nếu lỗi
    """
    if file and file.filename and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        # Thêm timestamp để tránh trùng tên
        from datetime import datetime
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        name, ext = os.path.splitext(filename)
        filename = f"{prefix}{timestamp}_{name}{ext}"
        
        # Lấy thư mục upload theo loại
        upload_folder = get_upload_folder(category)
        
        # Tạo thư mục nếu chưa có
        if not os.path.exists(upload_folder):
            os.makedirs(upload_folder)
        
        filepath = os.path.join(upload_folder, filename)
        file.save(filepath)
        
        # Trả về relative path từ thư mục static
        return os.path.join(category, filename).replace('\\', '/')
    
    return None

def save_word_file(file, prefix='', category='other'):
    """
    Lưu file Word vào thư mục uploads và trích xuất nội dung
    
    Args:
        file: File object từ request.files
        prefix: Tiền tố cho tên file (vd: 'detai_', 'giaotrinh_')
        category: Loại file để xác định thư mục lưu
    
    Returns:
        tuple: (relative_path, content) hoặc (None, None) nếu lỗi
    """
    if file and allowed_file(file.filename):
        relative_path = save_uploaded_file(file, prefix, category)
        if (relative_path):
            # Tạo full path để đọc file
            upload_folder = get_upload_folder(category)
            filename = os.path.basename(relative_path)
            filepath = os.path.join(upload_folder, filename)
            
            # Đọc nội dung file Word
            content = extract_text_from_word(filepath)
            return relative_path, content
    
    return None, None

def extract_text_from_word(filepath):
    """
    Trích xuất nội dung text từ file Word
    
    Args:
        filepath: Đường dẫn đến file Word
    
    Returns:
        str: Nội dung text của file
    """
    try:
        doc = Document(filepath)
        full_text = []
        
        # Đọc từng đoạn văn
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text)
        
        # Đọc text trong bảng
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        full_text.append(cell.text)
        
        return '\n'.join(full_text)
    except Exception as e:
        print(f"Lỗi khi đọc file Word: {e}")
        return ""

def delete_uploaded_file(filename, category='other'):
    """
    Xóa file đã upload
    
    Args:
        filename: Relative path của file cần xóa (vd: 'detai/20231223_120000_file.pdf')
        category: Loại file (nếu filename không chứa thư mục con)
    
    Returns:
        bool: True nếu xóa thành công, False nếu lỗi
    """
    try:
        if filename:
            # Nếu filename đã chứa thư mục con (vd: detai/file.pdf)
            if '/' in filename or '\\' in filename:
                filepath = os.path.join('static/uploads', filename)
            else:
                # Nếu chỉ có tên file, dùng category
                upload_folder = get_upload_folder(category)
                filepath = os.path.join(upload_folder, filename)
            
            if os.path.exists(filepath):
                os.remove(filepath)
                return True
    except Exception as e:
        print(f"Lỗi khi xóa file: {e}")
    return False

def delete_word_file(filename):
    """Wrapper cho delete_uploaded_file để tương thích với code cũ"""
    return delete_uploaded_file(filename)

def search_in_word_content(query, content):
    """
    Tìm kiếm trong nội dung file Word
    
    Args:
        query: Từ khóa tìm kiếm
        content: Nội dung file Word
    
    Returns:
        bool: True nếu tìm thấy, False nếu không
    """
    if not query or not content:
        return False
    
    # Tìm kiếm không phân biệt hoa thường
    return query.lower() in content.lower()
