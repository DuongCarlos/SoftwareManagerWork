import mysql.connector

class Database:
    def __init__(self, host, user, password, database):
        self.host = host
        self.user = user
        self.password = password
        self.database = database
        
        self.conn = mysql.connector.connect(
            host=self.host,
            user=self.user,
            password=self.password,
            database=self.database,
            port=3306,
            autocommit=True
        )
        self.cursor = self.conn.cursor(buffered=True) 
        self.create_table()

    def create_table(self):
        # Bảng Công Việc
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS cong_van (
                id INT AUTO_INCREMENT PRIMARY KEY,
                parent_id INT DEFAULT 0,
                ten_cong_viec TEXT,
                noi_dung TEXT,
                tai_lieu_dinh_kem TEXT,
                nguoi_xu_ly VARCHAR(255),
                nguoi_trinh_ky VARCHAR(255),
                ngay_nhan VARCHAR(50),
                ngay_bat_dau VARCHAR(50),
                ngay_ket_thuc VARCHAR(50),
                ngay_hoan_thanh VARCHAR(50),
                trang_thai VARCHAR(100),
                file_path TEXT,
                ngay_tam_dung VARCHAR(50) DEFAULT ''
            )
        ''')
        
        # Bảng Việc Định Kỳ
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS viec_dinh_ky (
                id INT AUTO_INCREMENT PRIMARY KEY,
                ten_cong_viec TEXT,
                noi_dung TEXT,
                nguoi_xu_ly VARCHAR(255),
                nguoi_trinh_ky VARCHAR(255),
                loai_lap VARCHAR(50),
                chi_tiet_lap VARCHAR(50),
                ngay_tao_gan_nhat VARCHAR(50),
                tai_lieu_dinh_kem TEXT,
                so_ngay_han INT DEFAULT 0,
                ngay_bat_dau_lap VARCHAR(50),
                file_path TEXT
            )
        ''')

        # BẢNG TÀI KHOẢN
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                username VARCHAR(50) PRIMARY KEY,
                password VARCHAR(100),
                role VARCHAR(20),
                fullname VARCHAR(255) DEFAULT ''
            )
        ''')
        
        try:
            self.cursor.execute("ALTER TABLE users ADD COLUMN fullname VARCHAR(255) DEFAULT ''")
        except:
            pass
            
        # NÂNG CẤP DB TỰ ĐỘNG: Tự động thăng cấp tài khoản 'admin' mặc định cũ lên 'super_admin'
        try:
            self.cursor.execute("UPDATE users SET role='super_admin' WHERE username='admin'")
        except:
            pass
        
        # Tự động tạo tài khoản Super Admin mặc định nếu chưa có
        self.cursor.execute("SELECT * FROM users WHERE username='admin'")
        if not self.cursor.fetchone():
            self.cursor.execute("INSERT INTO users (username, password, role, fullname) VALUES ('admin', 'admin', 'super_admin', 'Sếp Quản Trị')")

    # ================= QUẢN LÝ TÀI KHOẢN =================
    def check_login(self, username, password):
        self.cursor.execute("SELECT role FROM users WHERE username=%s AND password=%s", (username, password))
        res = self.cursor.fetchone()
        return res[0] if res else None

    def get_all_users(self):
        self.cursor.execute("SELECT username, password, role, fullname FROM users")
        return self.cursor.fetchall()

    def add_user(self, username, password, role, fullname):
        try:
            self.cursor.execute("INSERT INTO users (username, password, role, fullname) VALUES (%s, %s, %s, %s)", (username, password, role, fullname))
            return True
        except:
            return False

    # HÀM MỚI: SỬA TÀI KHOẢN
    def update_user(self, username, password, role, fullname):
        try:
            self.cursor.execute("UPDATE users SET password=%s, role=%s, fullname=%s WHERE username=%s", (password, role, fullname, username))
            return True
        except:
            return False

    def delete_user(self, username):
        self.cursor.execute("DELETE FROM users WHERE username=%s", (username,))

    # ================= CÁC HÀM CŨ GIỮ NGUYÊN =================
    def insert_cong_van(self, data):
        self.cursor.execute('''
            INSERT INTO cong_van (parent_id, ten_cong_viec, noi_dung, tai_lieu_dinh_kem, nguoi_xu_ly, nguoi_trinh_ky, ngay_nhan, ngay_bat_dau, ngay_ket_thuc, ngay_hoan_thanh, trang_thai, file_path, ngay_tam_dung)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ''', data)

    def update_cong_van(self, record_id, data):
        self.cursor.execute('''
            UPDATE cong_van 
            SET parent_id=%s, ten_cong_viec=%s, noi_dung=%s, tai_lieu_dinh_kem=%s, nguoi_xu_ly=%s, nguoi_trinh_ky=%s, ngay_nhan=%s, ngay_bat_dau=%s, ngay_ket_thuc=%s, ngay_hoan_thanh=%s, trang_thai=%s, file_path=%s, ngay_tam_dung=%s
            WHERE id=%s
        ''', (*data, record_id))

    def update_trang_thai(self, record_id, trang_thai):
        self.cursor.execute('UPDATE cong_van SET trang_thai=%s WHERE id=%s', (trang_thai, record_id))

    def delete_cong_van(self, record_id):
        self.cursor.execute('DELETE FROM cong_van WHERE id=%s OR parent_id=%s', (record_id, record_id))

    def get_all_cong_van(self):
        self.cursor.execute('SELECT id, parent_id, ten_cong_viec, noi_dung, tai_lieu_dinh_kem, nguoi_xu_ly, nguoi_trinh_ky, ngay_nhan, ngay_bat_dau, ngay_ket_thuc, ngay_hoan_thanh, trang_thai FROM cong_van ORDER BY id ASC')
        return self.cursor.fetchall()

    def get_all_for_excel(self):
        self.cursor.execute('SELECT id, parent_id, ten_cong_viec, noi_dung, tai_lieu_dinh_kem, nguoi_xu_ly, nguoi_trinh_ky, ngay_nhan, ngay_bat_dau, ngay_ket_thuc, ngay_hoan_thanh, trang_thai, file_path FROM cong_van ORDER BY id ASC')
        return self.cursor.fetchall()

    def insert_viec_dinh_ky(self, data):
        self.cursor.execute('''
            INSERT INTO viec_dinh_ky (ten_cong_viec, noi_dung, nguoi_xu_ly, nguoi_trinh_ky, loai_lap, chi_tiet_lap, ngay_tao_gan_nhat, tai_lieu_dinh_kem, so_ngay_han, ngay_bat_dau_lap, file_path)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ''', data)

    def update_viec_dinh_ky(self, record_id, data):
        self.cursor.execute('''
            UPDATE viec_dinh_ky 
            SET ten_cong_viec=%s, noi_dung=%s, nguoi_xu_ly=%s, nguoi_trinh_ky=%s, loai_lap=%s, chi_tiet_lap=%s, ngay_tao_gan_nhat=%s, tai_lieu_dinh_kem=%s, so_ngay_han=%s, ngay_bat_dau_lap=%s, file_path=%s
            WHERE id=%s
        ''', (*data, record_id))

    def get_all_viec_dinh_ky(self):
        self.cursor.execute('SELECT id, ten_cong_viec, noi_dung, nguoi_xu_ly, nguoi_trinh_ky, loai_lap, chi_tiet_lap, ngay_tao_gan_nhat, tai_lieu_dinh_kem, so_ngay_han, ngay_bat_dau_lap, file_path FROM viec_dinh_ky ORDER BY id ASC')
        return self.cursor.fetchall()

    def delete_viec_dinh_ky(self, record_id):
        self.cursor.execute("DELETE FROM viec_dinh_ky WHERE id=%s", (record_id,))

    def update_ngay_tao_gan_nhat(self, record_id, ngay_tao):
        self.cursor.execute("UPDATE viec_dinh_ky SET ngay_tao_gan_nhat=%s WHERE id=%s", (ngay_tao, record_id))

    def close(self):
        self.conn.close()