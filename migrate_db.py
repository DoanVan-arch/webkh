import sqlite3
import os

def migrate_database():
    """Thêm các cột còn thiếu vào các bảng"""
    db_path = 'instance/khoa_hoc_quan_su.db'
    
    if not os.path.exists(db_path):
        print(f"Không tìm thấy database tại {db_path}")
        return
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # ===== Migration cho bảng de_tai =====
    print("\n=== Migrating bảng de_tai ===")
    de_tai_columns = [
        ('file_word', 'VARCHAR(255)'),
        ('file_word_content', 'TEXT'),
        ('bien_ban_ban', 'VARCHAR(255)'),
        ('bien_ban_de_cuong', 'VARCHAR(255)'),
        ('bien_ban_nghiem_thu', 'VARCHAR(255)'),
        ('file_sau_sua', 'VARCHAR(255)'),
        ('ban_de_tai', 'VARCHAR(255)'),
        ('hoi_dong_nghiem_thu', 'VARCHAR(255)'),
        ('file_bien_ban', 'VARCHAR(255)'),
        ('ket_qua', 'VARCHAR(50)'),
        ('cap_do', 'VARCHAR(50)'),
    ]
    
    for column_name, column_type in de_tai_columns:
        try:
            cursor.execute(f"PRAGMA table_info(de_tai)")
            existing_columns = [row[1] for row in cursor.fetchall()]
            
            if column_name not in existing_columns:
                print(f"Đang thêm cột '{column_name}'...")
                cursor.execute(f"ALTER TABLE de_tai ADD COLUMN {column_name} {column_type}")
                conn.commit()
                print(f"✓ Đã thêm cột '{column_name}' thành công")
            else:
                print(f"- Cột '{column_name}' đã tồn tại, bỏ qua")
        except Exception as e:
            print(f"✗ Lỗi khi thêm cột '{column_name}': {str(e)}")
            conn.rollback()
    
    # ===== Migration cho bảng sang_kien =====
    print("\n=== Migrating bảng sang_kien ===")
    sang_kien_columns = [
        ('kinh_phi', 'FLOAT DEFAULT 0'),
        ('file_word', 'VARCHAR(255)'),
        ('file_word_content', 'TEXT'),
        ('bien_ban_ban', 'VARCHAR(255)'),
        ('bien_ban_de_cuong', 'VARCHAR(255)'),
        ('bien_ban_nghiem_thu', 'VARCHAR(255)'),
        ('file_sau_sua', 'VARCHAR(255)'),
        ('cap_do', 'VARCHAR(50)'),
    ]
    
    for column_name, column_type in sang_kien_columns:
        try:
            cursor.execute(f"PRAGMA table_info(sang_kien)")
            existing_columns = [row[1] for row in cursor.fetchall()]
            
            if column_name not in existing_columns:
                print(f"Đang thêm cột '{column_name}'...")
                cursor.execute(f"ALTER TABLE sang_kien ADD COLUMN {column_name} {column_type}")
                conn.commit()
                print(f"✓ Đã thêm cột '{column_name}' thành công")
            else:
                print(f"- Cột '{column_name}' đã tồn tại, bỏ qua")
        except Exception as e:
            print(f"✗ Lỗi khi thêm cột '{column_name}': {str(e)}")
            conn.rollback()
    
    # ===== Migration cho bảng tai_lieu =====
    print("\n=== Migrating bảng tai_lieu ===")
    tai_lieu_columns = [
        ('kinh_phi', 'FLOAT DEFAULT 0'),
        ('file_word', 'VARCHAR(255)'),
        ('file_word_content', 'TEXT'),
        ('bien_ban_ban', 'VARCHAR(255)'),
        ('bien_ban_de_cuong', 'VARCHAR(255)'),
        ('bien_ban_nghiem_thu', 'VARCHAR(255)'),
        ('file_sau_sua', 'VARCHAR(255)'),
    ]
    
    for column_name, column_type in tai_lieu_columns:
        try:
            cursor.execute(f"PRAGMA table_info(tai_lieu)")
            existing_columns = [row[1] for row in cursor.fetchall()]
            
            if column_name not in existing_columns:
                print(f"Đang thêm cột '{column_name}'...")
                cursor.execute(f"ALTER TABLE tai_lieu ADD COLUMN {column_name} {column_type}")
                conn.commit()
                print(f"✓ Đã thêm cột '{column_name}' thành công")
            else:
                print(f"- Cột '{column_name}' đã tồn tại, bỏ qua")
        except Exception as e:
            print(f"✗ Lỗi khi thêm cột '{column_name}': {str(e)}")
            conn.rollback()
    
    # ===== Migration cho bảng tiem_luc =====
    print("\n=== Migrating bảng tiem_luc ===")
    tiem_luc_columns = [
        ('chuc_danh', 'VARCHAR(100)'),
        ('nha_giao', 'VARCHAR(100)'),
    ]
    
    for column_name, column_type in tiem_luc_columns:
        try:
            cursor.execute(f"PRAGMA table_info(tiem_luc)")
            existing_columns = [row[1] for row in cursor.fetchall()]
            
            if column_name not in existing_columns:
                print(f"Đang thêm cột '{column_name}'...")
                cursor.execute(f"ALTER TABLE tiem_luc ADD COLUMN {column_name} {column_type}")
                conn.commit()
                print(f"✓ Đã thêm cột '{column_name}' thành công")
            else:
                print(f"- Cột '{column_name}' đã tồn tại, bỏ qua")
        except Exception as e:
            print(f"✗ Lỗi khi thêm cột '{column_name}': {str(e)}")
            conn.rollback()
    
    conn.close()
    print("\n=== Migration hoàn tất ===")

if __name__ == '__main__':
    migrate_database()
