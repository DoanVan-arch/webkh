"""
Script migration để cập nhật database:
1. Đổi tên cột chu_nhiem thành tac_gia trong bảng tai_lieu
2. Thêm cột file_tai_lieu vào bảng tai_lieu
"""
import sqlite3
import os

# Đường dẫn đến database
DB_PATH = os.path.join('instance', 'khoa_hoc_quan_su.db')

def migrate():
    """Thực hiện migration"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    try:
        # Kiểm tra xem cột tac_gia đã tồn tại chưa
        cursor.execute("PRAGMA table_info(tai_lieu)")
        columns = [col[1] for col in cursor.fetchall()]
        
        # 1. Đổi tên cột chu_nhiem thành tac_gia
        if 'chu_nhiem' in columns and 'tac_gia' not in columns:
            print("Đang đổi tên cột chu_nhiem thành tac_gia...")
            # SQLite không hỗ trợ ALTER COLUMN, phải tạo lại bảng
            # Nhưng vì chỉ đổi tên, ta có thể dùng cách đơn giản:
            cursor.execute("ALTER TABLE tai_lieu RENAME COLUMN chu_nhiem TO tac_gia")
            print("✓ Đã đổi tên cột chu_nhiem thành tac_gia")
        elif 'tac_gia' in columns:
            print("✓ Cột tac_gia đã tồn tại, bỏ qua bước này")
        
        # 2. Thêm cột file_tai_lieu nếu chưa có
        if 'file_tai_lieu' not in columns:
            print("Đang thêm cột file_tai_lieu...")
            cursor.execute("ALTER TABLE tai_lieu ADD COLUMN file_tai_lieu VARCHAR(255)")
            print("✓ Đã thêm cột file_tai_lieu")
        else:
            print("✓ Cột file_tai_lieu đã tồn tại, bỏ qua bước này")
        
        conn.commit()
        print("\n✅ Migration hoàn tất!")
        
    except Exception as e:
        conn.rollback()
        print(f"\n❌ Lỗi migration: {e}")
    finally:
        conn.close()

if __name__ == '__main__':
    print("=== BẮT ĐẦU MIGRATION TÀI LIỆU ===\n")
    migrate()
    print("\n=== KẾT THÚC MIGRATION ===")
