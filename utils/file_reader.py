import os
from docx import Document
import PyPDF2

def read_docx(file_path):
    """Đọc nội dung file DOCX"""
    try:
        doc = Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)
    except Exception as e:
        print(f"Error reading DOCX {file_path}: {e}")
        return ""

def read_pdf(file_path):
    """Đọc nội dung file PDF"""
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            full_text = []
            for page in pdf_reader.pages:
                full_text.append(page.extract_text())
            return '\n'.join(full_text)
    except Exception as e:
        print(f"Error reading PDF {file_path}: {e}")
        return ""

def read_file_content(file_path):
    """Đọc nội dung file dựa trên extension"""
    if not os.path.exists(file_path):
        return ""
    
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext in ['.docx', '.doc']:
        return read_docx(file_path)
    elif ext == '.pdf':
        return read_pdf(file_path)
    elif ext == '.txt':
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        except:
            return ""
    else:
        return ""

def search_in_file(file_path, search_query):
    """Tìm kiếm trong nội dung file"""
    content = read_file_content(file_path)
    if not content:
        return False, ""
    
    # Tìm kiếm không phân biệt hoa thường
    content_lower = content.lower()
    query_lower = search_query.lower()
    
    if query_lower in content_lower:
        # Lấy context xung quanh từ khóa (200 ký tự)
        index = content_lower.index(query_lower)
        start = max(0, index - 100)
        end = min(len(content), index + len(search_query) + 100)
        snippet = content[start:end]
        
        # Thêm ... nếu không phải đầu/cuối
        if start > 0:
            snippet = "..." + snippet
        if end < len(content):
            snippet = snippet + "..."
            
        return True, snippet
    
    return False, ""