import sqlite3

class Database:
    def __init__(self, db_path): 
        self.db_path = db_path 
        self.conn = sqlite3.connect(self.db_path)
        self.cursor = self.conn.cursor()
        self.create_table()

    def create_table(self):
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS cong_van (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                parent_id INTEGER DEFAULT 0,
                ten_cong_viec TEXT,
                noi_dung TEXT,
                tai_lieu_dinh_kem TEXT,
                nguoi_xu_ly TEXT,
                nguoi_trinh_ky TEXT,
                ngay_nhan TEXT,
                ngay_bat_dau TEXT,
                ngay_hoan_thanh TEXT,
                trang_thai TEXT,
                file_path TEXT
            )
        ''')
        self.conn.commit()

    def insert_cong_van(self, data):
        self.cursor.execute('''
            INSERT INTO cong_van (parent_id, ten_cong_viec, noi_dung, tai_lieu_dinh_kem, nguoi_xu_ly, nguoi_trinh_ky, ngay_nhan, ngay_bat_dau, ngay_hoan_thanh, trang_thai, file_path)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', data)
        self.conn.commit()

    def update_cong_van(self, record_id, data):
        self.cursor.execute('''
            UPDATE cong_van 
            SET parent_id=?, ten_cong_viec=?, noi_dung=?, tai_lieu_dinh_kem=?, nguoi_xu_ly=?, nguoi_trinh_ky=?, ngay_nhan=?, ngay_bat_dau=?, ngay_hoan_thanh=?, trang_thai=?, file_path=?
            WHERE id=?
        ''', (*data, record_id))
        self.conn.commit()

    def update_trang_thai(self, record_id, trang_thai):
        self.cursor.execute('UPDATE cong_van SET trang_thai=? WHERE id=?', (trang_thai, record_id))
        self.conn.commit()

    def delete_cong_van(self, record_id):
        self.cursor.execute('DELETE FROM cong_van WHERE id=? OR parent_id=?', (record_id, record_id))
        self.conn.commit()

    def get_all_cong_van(self):
        self.cursor.execute('''
            SELECT id, parent_id, ten_cong_viec, noi_dung, tai_lieu_dinh_kem, nguoi_xu_ly, nguoi_trinh_ky, ngay_nhan, ngay_bat_dau, ngay_hoan_thanh, trang_thai 
            FROM cong_van ORDER BY id ASC
        ''')
        return self.cursor.fetchall()

    def get_all_for_excel(self):
        self.cursor.execute('''
            SELECT id, parent_id, ten_cong_viec, noi_dung, tai_lieu_dinh_kem, nguoi_xu_ly, nguoi_trinh_ky, ngay_nhan, ngay_bat_dau, ngay_hoan_thanh, trang_thai, file_path 
            FROM cong_van ORDER BY id ASC
        ''')
        return self.cursor.fetchall()

    def get_main_tasks(self):
        self.cursor.execute("SELECT id, ten_cong_viec FROM cong_van WHERE parent_id=0 ORDER BY id ASC")
        return self.cursor.fetchall()

    def close(self):
        self.conn.close()