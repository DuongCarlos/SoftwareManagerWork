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
        
        # MIGRATION: Tự động cấy thêm các cột mới vào DB cũ an toàn
        self.cursor.execute("PRAGMA table_info(cong_van)")
        columns = [col[1] for col in self.cursor.fetchall()]
        if "ngay_ket_thuc" not in columns:
            self.cursor.execute("ALTER TABLE cong_van ADD COLUMN ngay_ket_thuc TEXT")
        if "ngay_tam_dung" not in columns:
            self.cursor.execute("ALTER TABLE cong_van ADD COLUMN ngay_tam_dung TEXT DEFAULT ''")
        self.conn.commit()

        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS viec_dinh_ky (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ten_cong_viec TEXT,
                noi_dung TEXT,
                nguoi_xu_ly TEXT,
                nguoi_trinh_ky TEXT,
                loai_lap TEXT,
                chi_tiet_lap TEXT,
                ngay_tao_gan_nhat TEXT
            )
        ''')
        self.conn.commit()
        
        self.cursor.execute("PRAGMA table_info(viec_dinh_ky)")
        cols_viec = [col[1] for col in self.cursor.fetchall()]
        if "tai_lieu_dinh_kem" not in cols_viec:
            self.cursor.execute("ALTER TABLE viec_dinh_ky ADD COLUMN tai_lieu_dinh_kem TEXT DEFAULT ''")
            self.cursor.execute("ALTER TABLE viec_dinh_ky ADD COLUMN so_ngay_han INTEGER DEFAULT 0")
            self.cursor.execute("ALTER TABLE viec_dinh_ky ADD COLUMN ngay_bat_dau_lap TEXT DEFAULT ''")
            self.cursor.execute("ALTER TABLE viec_dinh_ky ADD COLUMN file_path TEXT DEFAULT ''")
            self.conn.commit()

    def insert_cong_van(self, data):
        self.cursor.execute('''
            INSERT INTO cong_van (parent_id, ten_cong_viec, noi_dung, tai_lieu_dinh_kem, nguoi_xu_ly, nguoi_trinh_ky, ngay_nhan, ngay_bat_dau, ngay_ket_thuc, ngay_hoan_thanh, trang_thai, file_path, ngay_tam_dung)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', data)
        self.conn.commit()

    def update_cong_van(self, record_id, data):
        self.cursor.execute('''
            UPDATE cong_van 
            SET parent_id=?, ten_cong_viec=?, noi_dung=?, tai_lieu_dinh_kem=?, nguoi_xu_ly=?, nguoi_trinh_ky=?, ngay_nhan=?, ngay_bat_dau=?, ngay_ket_thuc=?, ngay_hoan_thanh=?, trang_thai=?, file_path=?, ngay_tam_dung=?
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
            SELECT id, parent_id, ten_cong_viec, noi_dung, tai_lieu_dinh_kem, nguoi_xu_ly, nguoi_trinh_ky, ngay_nhan, ngay_bat_dau, ngay_ket_thuc, ngay_hoan_thanh, trang_thai 
            FROM cong_van ORDER BY id ASC
        ''')
        return self.cursor.fetchall()

    def get_all_for_excel(self):
        self.cursor.execute('''
            SELECT id, parent_id, ten_cong_viec, noi_dung, tai_lieu_dinh_kem, nguoi_xu_ly, nguoi_trinh_ky, ngay_nhan, ngay_bat_dau, ngay_ket_thuc, ngay_hoan_thanh, trang_thai, file_path 
            FROM cong_van ORDER BY id ASC
        ''')
        return self.cursor.fetchall()

    def get_main_tasks(self):
        self.cursor.execute("SELECT id, ten_cong_viec FROM cong_van WHERE parent_id=0 ORDER BY id ASC")
        return self.cursor.fetchall()

    def insert_viec_dinh_ky(self, data):
        self.cursor.execute('''
            INSERT INTO viec_dinh_ky (ten_cong_viec, noi_dung, nguoi_xu_ly, nguoi_trinh_ky, loai_lap, chi_tiet_lap, ngay_tao_gan_nhat, tai_lieu_dinh_kem, so_ngay_han, ngay_bat_dau_lap, file_path)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', data)
        self.conn.commit()

    def update_viec_dinh_ky(self, record_id, data):
        self.cursor.execute('''
            UPDATE viec_dinh_ky 
            SET ten_cong_viec=?, noi_dung=?, nguoi_xu_ly=?, nguoi_trinh_ky=?, loai_lap=?, chi_tiet_lap=?, ngay_tao_gan_nhat=?, tai_lieu_dinh_kem=?, so_ngay_han=?, ngay_bat_dau_lap=?, file_path=?
            WHERE id=?
        ''', (*data, record_id))
        self.conn.commit()

    def get_all_viec_dinh_ky(self):
        self.cursor.execute('''
            SELECT id, ten_cong_viec, noi_dung, nguoi_xu_ly, nguoi_trinh_ky, loai_lap, chi_tiet_lap, ngay_tao_gan_nhat, tai_lieu_dinh_kem, so_ngay_han, ngay_bat_dau_lap, file_path 
            FROM viec_dinh_ky ORDER BY id ASC
        ''')
        return self.cursor.fetchall()

    def delete_viec_dinh_ky(self, record_id):
        self.cursor.execute("DELETE FROM viec_dinh_ky WHERE id=?", (record_id,))
        self.conn.commit()

    def update_ngay_tao_gan_nhat(self, record_id, ngay_tao):
        self.cursor.execute("UPDATE viec_dinh_ky SET ngay_tao_gan_nhat=? WHERE id=?", (ngay_tao, record_id))
        self.conn.commit()

    def close(self):
        self.conn.close()