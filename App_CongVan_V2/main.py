import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import shutil
import os
import sys 
import json
from datetime import datetime, timedelta
import pandas as pd 
from tkcalendar import DateEntry 
from gui import MainWindowUI
from database import Database

# ================= MÀN HÌNH ĐĂNG NHẬP =================
class LoginDialog(tk.Toplevel):
    def __init__(self, parent, db):
        super().__init__(parent)
        self.title("Đăng nhập Hệ thống")
        self.geometry("380x250")
        self.resizable(False, False)
        self.db = db
        self.result = None
        
        ttk.Label(self, text="🔑 ĐĂNG NHẬP PHẦN MỀM", font=("Segoe UI", 14, "bold")).pack(pady=(20, 15))
        
        frame = ttk.Frame(self)
        frame.pack(pady=5)
        ttk.Label(frame, text="Tài khoản:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.txt_user = ttk.Entry(frame, width=25)
        self.txt_user.grid(row=0, column=1, padx=5, pady=5)
        self.txt_user.focus()
        
        ttk.Label(frame, text="Mật khẩu:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.txt_pass = ttk.Entry(frame, width=25, show="*")
        self.txt_pass.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Button(self, text="Đăng Nhập", command=self.login).pack(pady=15)
        
        self.bind('<Return>', lambda e: self.login())
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.transient(parent)
        self.grab_set()

    def login(self):
        u = self.txt_user.get().strip()
        p = self.txt_pass.get().strip()
        role = self.db.check_login(u, p)
        if role:
            self.result = {"username": u, "role": role}
            self.destroy()
        else:
            messagebox.showerror("Lỗi", "Sai tên đăng nhập hoặc mật khẩu!", parent=self)
            
    def on_close(self):
        self.destroy()


class AppController:
    def __init__(self, root):
        self.root = root
        self.ui = MainWindowUI(root)
        
        if getattr(sys, 'frozen', False):
            self.app_dir = os.path.dirname(sys.executable)
        else:
            self.app_dir = os.path.dirname(os.path.abspath(__file__))

        self.kho_tai_lieu = os.path.join(self.app_dir, "Kho_Tai_Lieu")
        self.server_config = os.path.join(self.app_dir, "server_config.json")
        if not os.path.exists(self.kho_tai_lieu):
            os.makedirs(self.kho_tai_lieu)
            
        # ================= KẾT NỐI SERVER =================
        self.server_ip = "localhost"
        if os.path.exists(self.server_config):
            try:
                with open(self.server_config, 'r') as f:
                    data = json.load(f)
                    self.server_ip = data.get("host", "localhost")
            except: pass

        self.connect_to_server(self.server_ip, is_startup=True)
        
        # ================= GỌI MÀN HÌNH ĐĂNG NHẬP =================
        login_window = LoginDialog(self.root, self.db)
        self.root.wait_window(login_window)
        
        if not login_window.result:
            sys.exit()
            
        self.current_username = login_window.result['username']
        self.current_role = login_window.result['role']

        self.user_to_fullname = {} 
        self.fullname_to_user = {} 
        
        self.update_window_title()

        self.all_users_display = []
        self.admin_users_display = []
        self.assignable_by_admin = []
        self.load_users_to_comboboxes()

        self.filter_employee = "" 
        self.current_is_boss_task = False

        if self.current_role == 'super_admin':
            self.btn_manage_users = ttk.Button(self.ui.btn_frame, text="👥 Quản lý Nhân viên", command=self.open_user_management)
            self.btn_manage_users.pack(side=tk.RIGHT, padx=5)
            
        if self.current_role in ['admin', 'super_admin']:
            self.ui.sidebar_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(5, 5), pady=5, before=self.ui.tree_frame)
            self.load_sidebar_users()
            self.ui.listbox_nhan_vien.bind("<<ListboxSelect>>", self.on_sidebar_select)
            self.auto_spawn_scheduled_tasks()
            self.auto_scan_overdue_tasks()

        # ================= KHỞI TẠO NÚT BẤM =================
        self.current_selected_id = None
        self.current_parent_id = 0 
        self.current_attached_files = [] 

        self.ui.btn_add_new.config(command=self.open_add_popup) 
        self.ui.btn_edit.config(command=self.handle_edit_click)
        self.ui.btn_update.config(command=self.handle_update_click)
        self.ui.btn_delete.config(command=self.handle_delete_click)
        self.ui.btn_cancel.config(command=self.handle_cancel_click)
        self.ui.btn_upload.config(command=self.handle_upload)
        self.ui.btn_download.config(command=self.handle_download)
        self.ui.btn_remove_file.config(command=self.handle_remove_file)
        self.ui.btn_export_excel.config(command=self.handle_export_excel)
        self.ui.btn_search.config(command=self.handle_search)
        self.ui.btn_clear_search.config(command=self.handle_clear_search)
        self.ui.btn_schedule.config(command=self.open_schedule_popup) 
        self.ui.btn_config_ip.config(command=self.change_server_ip)

        self.ui.tree.bind("<<TreeviewSelect>>", self.handle_tree_click)
        self.ui.entries["Trạng Thái:"].bind("<<ComboboxSelected>>", self.on_main_status_change)
        
        self.set_form_state("disabled")
        self.ui.btn_edit.config(state=tk.DISABLED)
        self.refresh_table()

    # ================= HÀM XỬ LÝ SIDEBAR NHÂN VIÊN =================
    def load_sidebar_users(self):
        self.ui.listbox_nhan_vien.delete(0, tk.END)
        self.ui.listbox_nhan_vien.insert(tk.END, "🌟 Tất cả nhân viên")
        for disp_name in self.all_users_display:
            self.ui.listbox_nhan_vien.insert(tk.END, f"👤 {disp_name}")

    def on_sidebar_select(self, event):
        sel = self.ui.listbox_nhan_vien.curselection()
        if not sel: return
        val = self.ui.listbox_nhan_vien.get(sel[0])
        
        if "Tất cả" in val:
            self.filter_employee = ""
        else:
            disp_val = val.replace("👤 ", "").strip()
            self.filter_employee = self.fullname_to_user.get(disp_val, disp_val)
            
        self.refresh_table(search_kw=self.ui.entry_search_date.get().strip(), 
                           search_status=self.ui.combo_search_status.get())

    # ================= HÀM HỖ TRỢ NẠP & TÌM KIẾM NHÂN VIÊN =================
    def setup_user_combobox(self, cb, user_list):
        cb['values'] = user_list
        def search(event):
            typed = cb.get().strip().lower()
            if not typed:
                cb['values'] = user_list
            else:
                matches = [u for u in user_list if typed in u.lower()]
                cb['values'] = matches
            try: cb.event_generate('<Down>') 
            except: pass
        cb.bind("<Return>", search)

    def load_users_to_comboboxes(self):
        try:
            users_data = self.db.get_all_users() 
            self.user_to_fullname = {}
            self.fullname_to_user = {}
            
            for u in users_data:
                uname = u[0]
                fname = u[3]
                disp_str = fname if (fname and fname.strip() != "") else uname
                self.user_to_fullname[uname] = disp_str
                self.fullname_to_user[disp_str] = uname

            self.all_users_display = [self.user_to_fullname[u[0]] for u in users_data]
            self.admin_users_display = [self.user_to_fullname[u[0]] for u in users_data if u[2] in ['admin', 'super_admin']]
            self.assignable_by_admin = [self.user_to_fullname[u[0]] for u in users_data if u[2] != 'super_admin']

            if self.current_role == 'super_admin':
                self.setup_user_combobox(self.ui.entries["Người Xử Lý:"], self.all_users_display)
            elif self.current_role == 'admin':
                self.setup_user_combobox(self.ui.entries["Người Xử Lý:"], self.assignable_by_admin)
            else:
                self.setup_user_combobox(self.ui.entries["Người Xử Lý:"], self.all_users_display)
            
            # ĐÃ CHỈNH SỬA: Ô BÁO CÁO LUÔN LUÔN LÀ ADMIN_USERS_DISPLAY (CHỈ CÓ LÃNH ĐẠO)
            self.setup_user_combobox(self.ui.entries["Báo Cáo:"], self.admin_users_display)
            if self.current_role not in ['admin', 'super_admin'] and self.admin_users_display:
                self.ui.entries["Báo Cáo:"].set(self.admin_users_display[0])

        except Exception as e:
            print(f"Lỗi tải danh sách: {e}")

    # ================= MÀN HÌNH QUẢN LÝ TÀI KHOẢN =================
    def open_user_management(self):
        pop = tk.Toplevel(self.root)
        pop.title("Quản lý Tài khoản Nhân viên")
        pop.geometry("680x480")
        pop.transient(self.root)
        pop.grab_set()

        frame = ttk.LabelFrame(pop, text="Thông Tin Tài Khoản")
        frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(frame, text="Tài khoản:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        txt_u = ttk.Entry(frame, width=20)
        txt_u.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Mật khẩu:").grid(row=0, column=2, padx=5, pady=5, sticky='w')
        txt_p = ttk.Entry(frame, width=20, show="*")
        txt_p.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame, text="Tên đầy đủ:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        txt_f = ttk.Entry(frame, width=20)
        txt_f.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Quyền:").grid(row=1, column=2, padx=5, pady=5, sticky='w')
        allowed_roles = ["user", "admin", "super_admin"]
        combo_role = ttk.Combobox(frame, width=17, state="readonly", values=allowed_roles)
        combo_role.set("user")
        combo_role.grid(row=1, column=3, padx=5, pady=5)

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=2, column=0, columnspan=4, pady=10)

        def add_user():
            u = txt_u.get().strip()
            p = txt_p.get().strip()
            f = txt_f.get().strip()
            r = combo_role.get()
            if not u or not p:
                messagebox.showwarning("Thiếu thông tin", "Vui lòng nhập Tài khoản và Mật khẩu!")
                return
            if self.db.add_user(u, p, r, f):
                refresh_user_list()
                clear_form()
            else:
                messagebox.showerror("Lỗi", "Tên tài khoản đã tồn tại!")

        def update_user():
            u = txt_u.get().strip()
            p = txt_p.get().strip()
            f = txt_f.get().strip()
            r = combo_role.get()
            
            if txt_u['state'] == tk.NORMAL or not u:
                messagebox.showwarning("Cảnh báo", "Vui lòng CHỌN 1 tài khoản từ bảng bên dưới để Sửa!")
                return
            if not p:
                messagebox.showwarning("Thiếu thông tin", "Mật khẩu không được để trống!")
                return

            if self.db.update_user(u, p, r, f):
                refresh_user_list()
                clear_form()
                messagebox.showinfo("Thành công", "Đã cập nhật thông tin tài khoản!")
            else:
                messagebox.showerror("Lỗi", "Cập nhật thất bại!")

        def clear_form():
            txt_u.config(state=tk.NORMAL)
            txt_u.delete(0, tk.END)
            txt_p.delete(0, tk.END)
            txt_f.delete(0, tk.END)
            combo_role.set("user")

        ttk.Button(btn_frame, text="➕ Thêm Mới", command=add_user).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="💾 Sửa / Lưu Cập Nhật", command=update_user).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔄 Làm Mới Form", command=clear_form).pack(side=tk.LEFT, padx=5)

        list_frame = ttk.Frame(pop)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        cols = ("u", "f", "p", "r")
        tree_user = ttk.Treeview(list_frame, columns=cols, show="headings", height=8)
        tree_user.heading("u", text="Tài Khoản")
        tree_user.heading("f", text="Tên Đầy Đủ")
        tree_user.heading("p", text="Mật Khẩu")
        tree_user.heading("r", text="Quyền")
        tree_user.column("u", width=120)
        tree_user.column("f", width=160)
        tree_user.column("p", width=100, anchor="center")
        tree_user.column("r", width=100, anchor="center")
        tree_user.pack(fill=tk.BOTH, expand=True)

        def on_tree_select(event):
            sel = tree_user.selection()
            if not sel: return
            u_id = sel[0]
            for r in self.db.get_all_users():
                if r[0] == u_id:
                    txt_u.config(state=tk.NORMAL)
                    txt_u.delete(0, tk.END)
                    txt_u.insert(0, r[0])
                    txt_u.config(state="disabled") 
                    txt_p.delete(0, tk.END)
                    txt_p.insert(0, r[1])
                    txt_f.delete(0, tk.END)
                    txt_f.insert(0, r[3] if r[3] else "")
                    combo_role.set(r[2])
                    break

        tree_user.bind("<<TreeviewSelect>>", on_tree_select)

        def refresh_user_list():
            for row in tree_user.get_children(): tree_user.delete(row)
            for r in self.db.get_all_users():
                tree_user.insert("", tk.END, iid=r[0], values=(r[0], r[3], "******", r[2]))
            self.load_users_to_comboboxes()
            if self.current_role in ['admin', 'super_admin']:
                self.load_sidebar_users()
            self.refresh_table()
            
        refresh_user_list()

        def del_user():
            sel = tree_user.selection()
            if not sel:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn 1 tài khoản để Xóa!")
                return
            u = sel[0]
            if u == self.current_username:
                messagebox.showerror("Lỗi", "Không thể tự xóa tài khoản của chính mình đang đăng nhập!")
                return
                
            if messagebox.askyesno("Xác nhận", f"Bạn có chắc chắn muốn xóa tài khoản: {u}?"):
                self.db.delete_user(u)
                refresh_user_list()
                clear_form()

        ttk.Button(pop, text="🗑️ Xóa Tài Khoản Đang Chọn", command=del_user).pack(pady=10)

    def update_window_title(self):
        if self.current_role == "super_admin": role_str = "Sếp Tổng (Super Admin)"
        elif self.current_role == "admin": role_str = "Lãnh đạo (Admin)"
        else: role_str = "Nhân viên"
            
        disp_name = self.user_to_fullname.get(self.current_username, self.current_username) if hasattr(self, 'user_to_fullname') else self.current_username
        self.root.title(f"Quản lý Công việc - [{role_str}: {disp_name}] - Máy chủ: {self.server_ip}")

    def connect_to_server(self, ip, is_startup=False):
        try:
            if hasattr(self, 'db') and self.db:
                self.db.close()
            self.db = Database(host=ip, user="admin", password="123456", database="db_qlcongviec")
            self.server_ip = ip
            with open(self.server_config, 'w') as f:
                json.dump({"host": ip}, f)
            if not is_startup:
                messagebox.showinfo("Thành công", f"Đã kết nối Máy chủ: {ip}")
                self.refresh_table()
        except Exception as e:
            if is_startup:
                new_ip = simpledialog.askstring("Mất kết nối Máy Chủ", f"Không thể kết nối đến Máy chủ [{ip}].\nNhập IP Máy chủ nội bộ để tiếp tục:\n(Ví dụ: 192.168.1.45)", parent=self.root)
                if new_ip: self.connect_to_server(new_ip.strip(), is_startup=True)
                else: sys.exit() 
            else:
                messagebox.showerror("Lỗi kết nối", f"Lỗi: {e}")

    def change_server_ip(self):
        new_ip = simpledialog.askstring("Cấu hình Máy Chủ", f"IP Máy chủ hiện tại: {self.server_ip}\nNhập địa chỉ IP mới:", initialvalue=self.server_ip, parent=self.root)
        if new_ip and new_ip.strip() != self.server_ip:
            self.connect_to_server(new_ip.strip())

    def build_tree_data(self, rows):
        tree = {}
        row_dict = {}
        for r in rows:
            row_dict[r[0]] = r
            if r[1] not in tree: tree[r[1]] = []
            tree[r[1]].append(r)
        return tree, row_dict

    def handle_search(self):
        kw = self.ui.entry_search_date.get().strip()
        status = self.ui.combo_search_status.get()
        self.refresh_table(search_kw=kw, search_status=status)

    def handle_clear_search(self):
        self.ui.entry_search_date.delete(0, tk.END)
        self.ui.combo_search_status.set("Tất cả")
        if self.current_role in ['admin', 'super_admin']:
            self.filter_employee = ""
            self.ui.listbox_nhan_vien.selection_clear(0, tk.END)
        self.refresh_table()

    def get_tag_color(self, trang_thai):
        if trang_thai == "Đã hoàn thành": return 'hoan_thanh'
        elif trang_thai == "Đang xử lý": return 'dang_xu_ly'
        elif trang_thai == "Đang đợi phản hồi": return 'doi_phan_hoi'
        elif trang_thai == "Chậm tiến độ": return 'cham_tien_do'
        elif trang_thai == "Hoàn thành chậm": return 'hoan_thanh_cham'
        return ''

    def refresh_table(self, search_kw="", search_status="Tất cả"):
        for item in self.ui.tree.get_children():
            self.ui.tree.delete(item)
            
        rows = self.db.get_all_cong_van()
        tree_data, row_dict = self.build_tree_data(rows)
        
        progress_map = {}
        for r in rows:
            if r[0] in tree_data and len(tree_data[r[0]]) > 0: 
                total_children = len(tree_data[r[0]])
                completed_children = sum(1 for c in tree_data[r[0]] if c[11] in ["Đã hoàn thành", "Hoàn thành chậm"])
                pct = int((completed_children / total_children) * 100)
                progress_map[r[0]] = f"{pct}%"
            else: 
                progress_map[r[0]] = "" 
        
        valid_ids = set()
        is_searching = True
        
        for r in rows:
            has_permission = False
            nxl = str(r[5] or "").lower()
            ntk = str(r[6] or "").lower()

            if self.current_role in ['admin', 'super_admin']:
                if hasattr(self, 'filter_employee') and self.filter_employee:
                    if self.filter_employee.lower() in nxl or self.filter_employee.lower() in ntk:
                        has_permission = True
                else:
                    has_permission = True
            else:
                uname = self.current_username.lower()
                if uname in nxl or uname in ntk:
                    has_permission = True
                    
            if not has_permission: continue
                
            match_kw = True
            if search_kw:
                kw_lower = search_kw.lower()
                id_str = str(r[0])
                ten_str = str(r[2] or "").lower()
                nd_str = str(r[3] or "").lower()
                d1, d2, d3, d4 = str(r[7] or ""), str(r[8] or ""), str(r[9] or ""), str(r[10] or "")
                
                nxl_disp = self.user_to_fullname.get(r[5], r[5]).lower() if r[5] else ""
                ntk_disp = self.user_to_fullname.get(r[6], r[6]).lower() if r[6] else ""
                
                if (kw_lower not in id_str and kw_lower not in ten_str and kw_lower not in nd_str and 
                    kw_lower not in d1 and kw_lower not in d2 and kw_lower not in d3 and kw_lower not in d4 and
                    kw_lower not in nxl_disp and kw_lower not in ntk_disp):
                    match_kw = False
            
            match_status = True
            if search_status != "Tất cả":
                if str(r[11]) != search_status: match_status = False
                    
            if match_kw and match_status: valid_ids.add(r[0]) 
        
        added_new = True
        while added_new:
            added_new = False
            for r in rows:
                if r[0] in valid_ids and r[1] != 0 and r[1] not in valid_ids:
                    valid_ids.add(r[1])
                    added_new = True
        
        stt_map = {}
        stt_counter = 1 
        for r in rows:
            if r[1] == 0: 
                stt_map[r[0]] = stt_counter
                stt_counter += 1
                
        def insert_node(parent_id, current_iid, level):
            if parent_id not in tree_data: return
            for r in tree_data[parent_id]:
                if is_searching and r[0] not in valid_ids: continue 
                    
                tag = self.get_tag_color(r[11])
                ty_le = progress_map.get(r[0], "")
                
                if level == 0:
                    stt = stt_map.get(r[0], "")
                    ten_cv = r[2]
                else:
                    stt = ""
                    indent = "   " * level
                    ten_cv = f"{indent}↳ {r[2]}"
                    
                nxl_disp = self.user_to_fullname.get(r[5], r[5]) if r[5] else ""
                ntk_disp = self.user_to_fullname.get(r[6], r[6]) if r[6] else ""
                
                display_values = (stt, ten_cv, r[3], r[4], nxl_disp, ntk_disp, r[7], r[8], r[9], r[10], r[11], ty_le)
                
                try:
                    self.ui.tree.insert(current_iid, tk.END, iid=str(r[0]), values=display_values, tags=(tag,))
                    if current_iid != "": self.ui.tree.item(current_iid, open=True) 
                except tk.TclError:
                    self.ui.tree.insert("", tk.END, iid=str(r[0]), values=display_values, tags=(tag,))
                
                insert_node(r[0], str(r[0]), level + 1)
                
        insert_node(0, "", 0)

    # ================= TỰ ĐỘNG SINH MẪU ĐỊNH KỲ =================
    def auto_spawn_scheduled_tasks(self):
        today = datetime.now()
        today_str = today.strftime("%d/%m/%Y")
        weekdays = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ nhật"]
        today_weekday = weekdays[today.weekday()]
        today_day = str(today.day)
        
        tasks = self.db.get_all_viec_dinh_ky()
        for t in tasks:
            t_id, t_ten, t_nd, t_nxl, t_ntk, t_loai, t_chitiet, t_ngaytao, t_tldk, t_songay, t_ngaybd, t_filepath = t
            if t_ngaybd and t_ngaybd != "":
                if self.is_date_after(t_ngaybd, today_str): continue 
            should_create = False
            if str(t_ngaytao) != today_str:
                if t_loai == "Hàng ngày": should_create = True
                elif t_loai == "Hàng tuần" and str(t_chitiet) == today_weekday: should_create = True
                elif t_loai == "Hàng tháng" and str(t_chitiet) == today_day: should_create = True
            if should_create:
                try: days = int(t_songay)
                except: days = 0
                ngay_kt_str = (today + timedelta(days=days)).strftime("%d/%m/%Y")
                new_cv = (0, f"[Định kỳ] {t_ten}", str(t_nd), str(t_tldk), str(t_nxl), str(t_ntk), today_str, today_str, ngay_kt_str, "", "Đang xử lý", str(t_filepath), "")
                self.db.insert_cong_van(new_cv)
                self.db.update_ngay_tao_gan_nhat(t_id, today_str)

    def open_schedule_popup(self):
        self.sched_pop = tk.Toplevel(self.root)
        self.sched_pop.title("Quản lý Công việc Định kỳ")
        self.sched_pop.geometry("1100x850")
        self.sched_pop.transient(self.root) 
        self.sched_pop.grab_set() 
        self.sch_attached_files = []
        self.current_sch_id = None
        
        form_frame = ttk.LabelFrame(self.sched_pop, text="Cài đặt mẫu công việc định kỳ")
        form_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(form_frame, text="Tên công việc:").grid(row=0, column=0, padx=5, pady=5, sticky="nw")
        self.sch_ten = tk.Text(form_frame, height=2, width=45, wrap=tk.WORD)
        self.sch_ten.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        
        ttk.Label(form_frame, text="Người Xử Lý:").grid(row=0, column=3, padx=5, pady=5, sticky="nw")
        self.sch_nxl = ttk.Combobox(form_frame, width=23)
        self.sch_nxl.grid(row=0, column=4, padx=5, pady=5, sticky="w")
        
        if self.current_role == 'super_admin':
            self.setup_user_combobox(self.sch_nxl, self.all_users_display)
        elif self.current_role == 'admin':
            self.setup_user_combobox(self.sch_nxl, self.assignable_by_admin)
        else:
            self.setup_user_combobox(self.sch_nxl, self.all_users_display)
            
        ttk.Label(form_frame, text="Báo Cáo:").grid(row=0, column=5, padx=5, pady=5, sticky="nw")
        self.sch_ntk = ttk.Combobox(form_frame, width=23)
        self.sch_ntk.grid(row=0, column=6, padx=5, pady=5, sticky="w")
        
        # ĐÃ CHỈNH SỬA: Ô BÁO CÁO LUÔN LUÔN LÀ ADMIN_USERS_DISPLAY (CHỈ CÓ LÃNH ĐẠO)
        self.setup_user_combobox(self.sch_ntk, self.admin_users_display)
        if self.current_role not in ['admin', 'super_admin'] and self.admin_users_display:
            self.sch_ntk.insert(0, self.admin_users_display[0])
        
        ttk.Label(form_frame, text="Nội dung:").grid(row=1, column=0, padx=5, pady=5, sticky="nw")
        self.sch_nd = tk.Text(form_frame, height=3, width=105, wrap=tk.WORD)
        self.sch_nd.grid(row=1, column=1, columnspan=6, padx=5, pady=5, sticky="w")
        
        ttk.Label(form_frame, text="Chu kỳ lặp:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.sch_loai = ttk.Combobox(form_frame, state="readonly", values=["Hàng ngày", "Hàng tuần", "Hàng tháng"], width=15)
        self.sch_loai.set("Hàng ngày")
        self.sch_loai.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(form_frame, text="Chi tiết:").grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.sch_chitiet = ttk.Combobox(form_frame, state="disabled", width=15)
        self.sch_chitiet.grid(row=2, column=3, padx=5, pady=5, sticky="w")
        ttk.Label(form_frame, text="Ngày BĐ Áp dụng:").grid(row=2, column=4, padx=5, pady=5, sticky="e")
        self.sch_ngaybd = DateEntry(form_frame, width=15, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.sch_ngaybd.grid(row=2, column=5, padx=5, pady=5, sticky="w")
        ttk.Label(form_frame, text="Hạn (số ngày):").grid(row=2, column=6, padx=5, pady=5, sticky="w")
        self.sch_songay = ttk.Entry(form_frame, width=10)
        self.sch_songay.insert(0, "0")
        self.sch_songay.grid(row=2, column=6, padx=(100, 5), pady=5, sticky="w")
        
        def on_loai_change(event):
            val = self.sch_loai.get()
            if val == "Hàng ngày":
                self.sch_chitiet.set("")
                self.sch_chitiet.config(state="disabled")
            elif val == "Hàng tuần":
                self.sch_chitiet.config(state="readonly")
                self.sch_chitiet['values'] = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ nhật"]
                self.sch_chitiet.set("Thứ 2")
            elif val == "Hàng tháng":
                self.sch_chitiet.config(state="readonly")
                self.sch_chitiet['values'] = [str(i) for i in range(1, 32)]
                self.sch_chitiet.set("1")
        self.sch_loai.bind("<<ComboboxSelected>>", on_loai_change)
        
        file_frame = ttk.Frame(form_frame)
        file_frame.grid(row=3, column=0, columnspan=7, pady=10, sticky="w", padx=5)
        ttk.Label(file_frame, text="File đính kèm:").pack(side=tk.LEFT, anchor="n")
        self.sch_listbox_files = tk.Listbox(file_frame, height=3, width=75, selectmode=tk.EXTENDED)
        self.sch_listbox_files.pack(side=tk.LEFT, padx=10)
        btn_box = ttk.Frame(file_frame)
        btn_box.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_box, text="➕ Chọn File", command=self.sch_handle_upload).pack(fill=tk.X, pady=1)
        ttk.Button(btn_box, text="❌ Xóa File", command=self.sch_handle_remove).pack(fill=tk.X, pady=1)
        
        btn_frame = ttk.Frame(self.sched_pop)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(btn_frame, text="❌ Hủy chọn", command=self.sch_handle_cancel).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🗑️ Xóa Mẫu", command=self.delete_schedule).pack(side=tk.LEFT, padx=5)
        self.btn_sch_save = ttk.Button(btn_frame, text="💾 LƯU MẪU MỚI", command=self.save_schedule)
        self.btn_sch_save.pack(side=tk.RIGHT, padx=5)
        self.btn_sch_update = ttk.Button(btn_frame, text="🔄 CẬP NHẬT MẪU", command=self.update_schedule, state=tk.DISABLED)
        self.btn_sch_update.pack(side=tk.RIGHT, padx=5)
        
        list_frame = ttk.LabelFrame(self.sched_pop, text="Danh sách Mẫu đang chạy tự động")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        cols = ("id", "ten", "loai", "chitiet", "ngaybd", "songay", "nxl", "ngaytao")
        self.tree_sch = ttk.Treeview(list_frame, columns=cols, show="headings", height=10)
        self.tree_sch.heading("id", text="ID")
        self.tree_sch.heading("ten", text="Tên Công Việc")
        self.tree_sch.heading("loai", text="Chu Kỳ")
        self.tree_sch.heading("chitiet", text="Chi Tiết")
        self.tree_sch.heading("ngaybd", text="Ngày BĐ")
        self.tree_sch.heading("songay", text="Hạn(Ngày)")
        self.tree_sch.heading("nxl", text="Người XL")
        self.tree_sch.heading("ngaytao", text="Tạo Gần Nhất")
        self.tree_sch.column("id", width=40, anchor="center")
        self.tree_sch.column("ten", width=250)
        self.tree_sch.column("loai", width=100, anchor="center")
        self.tree_sch.column("chitiet", width=80, anchor="center")
        self.tree_sch.column("ngaybd", width=100, anchor="center")
        self.tree_sch.column("songay", width=80, anchor="center")
        self.tree_sch.column("nxl", width=120, anchor="center")
        self.tree_sch.column("ngaytao", width=120, anchor="center")
        self.tree_sch.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.tree_sch.bind("<<TreeviewSelect>>", self.on_sch_tree_select)
        self.refresh_schedule_table()

    def sch_handle_upload(self):
        file_paths = filedialog.askopenfilenames(title="Chọn file đính kèm", filetypes=[("Tất cả", "*.*")])
        for path in file_paths:
            if path not in self.sch_attached_files:
                self.sch_attached_files.append(path)
                self.sch_listbox_files.insert(tk.END, os.path.basename(path))

    def sch_handle_remove(self):
        selected = self.sch_listbox_files.curselection()
        if not selected: return
        for idx in reversed(selected):
            self.sch_listbox_files.delete(idx)
            self.sch_attached_files.pop(idx)

    def sch_handle_cancel(self):
        self.current_sch_id = None
        self.sch_attached_files = []
        self.sch_listbox_files.delete(0, tk.END)
        self.sch_ten.delete("1.0", tk.END)
        self.sch_nd.delete("1.0", tk.END)
        self.sch_nxl.set('')
        self.sch_ntk.set('')
        self.sch_songay.delete(0, tk.END)
        self.sch_songay.insert(0, "0")
        self.btn_sch_save.config(state=tk.NORMAL)
        self.btn_sch_update.config(state=tk.DISABLED)

    def on_sch_tree_select(self, event):
        sel = self.tree_sch.selection()
        if not sel: return
        self.sch_handle_cancel()
        self.current_sch_id = int(sel[0])
        self.btn_sch_save.config(state=tk.DISABLED)
        self.btn_sch_update.config(state=tk.NORMAL)
        tasks = self.db.get_all_viec_dinh_ky()
        for t in tasks:
            if t[0] == self.current_sch_id:
                self.sch_ten.insert(tk.END, t[1])
                self.sch_nd.insert(tk.END, t[2])
                
                self.sch_nxl.set(self.user_to_fullname.get(t[3], t[3]) if t[3] else "")
                self.sch_ntk.set(self.user_to_fullname.get(t[4], t[4]) if t[4] else "")
                
                self.sch_loai.set(t[5])
                if t[5] == "Hàng ngày":
                    self.sch_chitiet.config(state="disabled")
                    self.sch_chitiet.set("")
                elif t[5] == "Hàng tuần":
                    self.sch_chitiet.config(state="readonly")
                    self.sch_chitiet['values'] = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ nhật"]
                    self.sch_chitiet.set(t[6])
                elif t[5] == "Hàng tháng":
                    self.sch_chitiet.config(state="readonly")
                    self.sch_chitiet['values'] = [str(i) for i in range(1, 32)]
                    self.sch_chitiet.set(t[6])
                if t[8] and t[8] != "None":
                    self.sch_attached_files = t[11].split("|") if t[11] else []
                    for f in self.sch_attached_files:
                        if f.strip(): self.sch_listbox_files.insert(tk.END, os.path.basename(f))
                self.sch_songay.delete(0, tk.END)
                self.sch_songay.insert(0, str(t[9]))
                if t[10]: 
                    self.sch_ngaybd.delete(0, tk.END)
                    self.sch_ngaybd.insert(0, t[10])
                break

    def get_sch_form_data(self):
        ten = self.sch_ten.get("1.0", tk.END).strip()
        if not ten: return None
        nd = self.sch_nd.get("1.0", tk.END).strip()
        
        nxl_disp = self.sch_nxl.get().strip()
        ntk_disp = self.sch_ntk.get().strip()
        nxl = self.fullname_to_user.get(nxl_disp, nxl_disp)
        ntk = self.fullname_to_user.get(ntk_disp, ntk_disp)
        
        loai = self.sch_loai.get()
        chitiet = self.sch_chitiet.get()
        ngaybd = self.sch_ngaybd.get()
        try: songay = int(self.sch_songay.get().strip())
        except: songay = 0
        safe_paths = []
        for path in self.sch_attached_files:
            safe_p = self.process_attached_file(path)
            if safe_p: safe_paths.append(safe_p)
        safe_file_path_str = "|".join(safe_paths)
        tai_lieu_dinh_kem_str = " | ".join([os.path.basename(p) for p in safe_paths])
        return (ten, nd, nxl, ntk, loai, chitiet, tai_lieu_dinh_kem_str, songay, ngaybd, safe_file_path_str)

    def save_schedule(self):
        data = self.get_sch_form_data()
        if not data: return
        ten, nd, nxl, ntk, loai, chitiet, tldk, songay, ngaybd, fpath = data
        self.db.insert_viec_dinh_ky((ten, nd, nxl, ntk, loai, chitiet, "", tldk, songay, ngaybd, fpath))
        self.sch_handle_cancel()
        self.refresh_schedule_table()

    def update_schedule(self):
        if not self.current_sch_id: return
        data = self.get_sch_form_data()
        if not data: return
        ten, nd, nxl, ntk, loai, chitiet, tldk, songay, ngaybd, fpath = data
        tasks = self.db.get_all_viec_dinh_ky()
        ngaytao = ""
        for t in tasks:
            if t[0] == self.current_sch_id:
                ngaytao = t[7]
                break
        self.db.update_viec_dinh_ky(self.current_sch_id, (ten, nd, nxl, ntk, loai, chitiet, ngaytao, tldk, songay, ngaybd, fpath))
        self.sch_handle_cancel()
        self.refresh_schedule_table()

    def delete_schedule(self):
        sel = self.tree_sch.selection()
        if not sel: return
        if messagebox.askyesno("Xác nhận", "Hủy bỏ tự động hóa công việc này?"):
            self.db.delete_viec_dinh_ky(sel[0])
            self.sch_handle_cancel()
            self.refresh_schedule_table()

    def refresh_schedule_table(self):
        for row in self.tree_sch.get_children(): self.tree_sch.delete(row)
        for t in self.db.get_all_viec_dinh_ky():
            nxl_disp = self.user_to_fullname.get(t[3], t[3]) if t[3] else ""
            self.tree_sch.insert("", tk.END, iid=str(t[0]), values=(t[0], t[1], t[5], t[6], t[10], t[9], nxl_disp, t[7]))

    # ================= NÂNG CẤP BẢO MẬT: KHÓA FORM NGHIÊM NGẶT =================
    def set_form_state(self, state, is_boss_task=False):
        is_user = self.current_role not in ['admin', 'super_admin']
        lock_all = is_user and is_boss_task
        
        text_state = tk.DISABLED if lock_all else (tk.NORMAL if state == "normal" else tk.DISABLED)
        self.ui.txt_ten_cv.config(state=text_state)
        self.ui.txt_noi_dung.config(state=text_state)
        
        for field, entry in self.ui.entries.items():
            if field == "Trạng Thái:":
                entry.config(state="readonly" if state == "normal" else "disabled")
            elif field == "Ngày Hoàn Thành:":
                entry.config(state="disabled") 
            elif field == "Người Xử Lý:" and is_user:
                entry.config(state="disabled") 
            elif field == "Báo Cáo:" and is_user:
                if lock_all:
                    entry.config(state="disabled")
                else:
                    entry.config(state=tk.NORMAL if state == "normal" else tk.DISABLED)
            else:
                entry_state = tk.DISABLED if lock_all else (tk.NORMAL if state == "normal" else tk.DISABLED)
                if isinstance(entry, ttk.Combobox) and entry_state == tk.NORMAL:
                    entry_state = "readonly" if field == "Trạng Thái:" else tk.NORMAL
                entry.config(state=entry_state)
                
        btn_state = tk.DISABLED if lock_all else (tk.NORMAL if state == "normal" else tk.DISABLED)
        self.ui.btn_upload.config(state=btn_state)
        self.ui.btn_remove_file.config(state=btn_state)
        self.ui.btn_update.config(state=tk.NORMAL if state == "normal" else tk.DISABLED) 

    def handle_edit_click(self):
        if not self.current_selected_id:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn 1 công việc để chỉnh sửa!")
            return
        
        self.set_form_state("normal", is_boss_task=getattr(self, 'current_is_boss_task', False))
        self.ui.btn_edit.config(state=tk.DISABLED)

    def on_main_status_change(self, event=None):
        val = self.ui.entries["Trạng Thái:"].get()
        self.ui.entries["Ngày Hoàn Thành:"].config(state="normal")
        if val in ["Đã hoàn thành", "Hoàn thành chậm"]:
            self.ui.entries["Ngày Hoàn Thành:"].set_date(datetime.now())
        else:
            self.ui.entries["Ngày Hoàn Thành:"].delete(0, tk.END)
        self.ui.entries["Ngày Hoàn Thành:"].config(state="disabled")

    # ================= POPUP THÊM MỚI (TẠO CÔNG VIỆC) =================
    def open_add_popup(self):
        self.popup = tk.Toplevel(self.root)
        self.popup.title("Thêm Mới Công Việc")
        self.popup.geometry("900x700") 
        self.popup.transient(self.root) 
        self.popup.grab_set() 
        self.popup_attached_files = []
        
        type_frame = ttk.LabelFrame(self.popup, text="Loại Công Việc")
        type_frame.pack(fill=tk.X, padx=15, pady=10)
        self.task_type_var = tk.IntVar(value=0) 
        
        rows = self.db.get_all_cong_van()
        tree_data, row_dict = self.build_tree_data(rows)
        
        self.all_task_strs = []
        self.main_task_mapping = {}
        
        stt_counter = 1
        def build_combo(p_id, level):
            nonlocal stt_counter
            if p_id not in tree_data: return
            for r in tree_data[p_id]:
                if level == 0:
                    name = f"[STT: {stt_counter}] {r[2]} (ID:{r[0]})"
                    stt_counter += 1
                else:
                    name = f"{'   ' * level}↳ {r[2]} (ID:{r[0]})"
                self.all_task_strs.append(name)
                self.main_task_mapping[name] = r[0]
                build_combo(r[0], level + 1)
                
        build_combo(0, 0)
        
        ttk.Label(type_frame, text="Thuộc công việc:").grid(row=0, column=2, padx=10)
        self.cb_parent = ttk.Combobox(type_frame, state="disabled", width=55)
        self.cb_parent.grid(row=0, column=3, padx=5, pady=10)
        
        PLACEHOLDER = "Gõ STT, Tên hoặc ID rồi nhấn Enter..."
        def toggle_parent_cb():
            if self.task_type_var.get() == 1:
                self.cb_parent.config(state="normal") 
                self.cb_parent['values'] = self.all_task_strs
                self.cb_parent.set(PLACEHOLDER) 
            else:
                self.cb_parent.set("")
                self.cb_parent.config(state="disabled")

        def on_focus_in(event):
            if self.cb_parent.get() == PLACEHOLDER: self.cb_parent.set("") 

        def on_focus_out(event):
            if self.cb_parent.get().strip() == "" and self.task_type_var.get() == 1: self.cb_parent.set(PLACEHOLDER) 

        self.cb_parent.bind("<FocusIn>", on_focus_in)
        self.cb_parent.bind("<FocusOut>", on_focus_out)

        def trigger_search(event):
            typed = self.cb_parent.get().strip().lower()
            if typed == "" or typed == PLACEHOLDER.lower():
                self.cb_parent['values'] = self.all_task_strs
            else:
                matches = []
                if typed.isdigit():
                    exact_stt = f"[stt: {typed}]"
                    exact_id = f"(id:{typed})"
                    for s in self.all_task_strs:
                        if exact_stt in s.lower() or exact_id in s.lower(): matches.append(s)
                if not matches:
                    matches = [s for s in self.all_task_strs if typed in s.lower()]
                self.cb_parent['values'] = matches
                if not matches: self.cb_parent.set("Không tìm thấy!")
            self.cb_parent.event_generate('<Down>')
            
        self.cb_parent.bind("<Return>", trigger_search)
        ttk.Radiobutton(type_frame, text="Công việc mới", variable=self.task_type_var, value=0, command=toggle_parent_cb).grid(row=0, column=0, padx=15, pady=10)
        ttk.Radiobutton(type_frame, text="Công việc con", variable=self.task_type_var, value=1, command=toggle_parent_cb).grid(row=0, column=1, padx=15, pady=10)
            
        form_frame = ttk.LabelFrame(self.popup, text="Thông Tin Chi Tiết")
        form_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        
        ttk.Label(form_frame, text="Tên Công Việc:").grid(row=0, column=0, sticky="nw", padx=10, pady=5)
        self.pop_txt_ten = tk.Text(form_frame, height=2, width=80, wrap=tk.WORD)
        self.pop_txt_ten.grid(row=0, column=1, columnspan=3, padx=10, pady=5, sticky="w")
        ttk.Label(form_frame, text="Nội Dung Chi Tiết:").grid(row=1, column=0, sticky="nw", padx=10, pady=5)
        self.pop_txt_nd = tk.Text(form_frame, height=3, width=80, wrap=tk.WORD)
        self.pop_txt_nd.grid(row=1, column=1, columnspan=3, padx=10, pady=5, sticky="w")
        
        self.pop_entries = {}
        fields = ["Người Xử Lý:", "Báo Cáo:", "Ngày Nhận:", "Ngày Bắt Đầu:", "Ngày Kết Thúc:", "Ngày Hoàn Thành:", "Trạng Thái:"]
        for idx, field in enumerate(fields):
            row = 2 + (idx // 2)
            col = (idx % 2) * 2
            ttk.Label(form_frame, text=field).grid(row=row, column=col, sticky="w", padx=10, pady=5)
            if field == "Trạng Thái:":
                entry = ttk.Combobox(form_frame, width=28, state="readonly", values=["Đang xử lý", "Đang đợi phản hồi", "Đã hoàn thành", "Chậm tiến độ", "Hoàn thành chậm"])
                entry.set("Đang xử lý")
                def on_pop_status_change(event):
                    val = self.pop_entries["Trạng Thái:"].get()
                    self.pop_entries["Ngày Hoàn Thành:"].config(state="normal")
                    if val in ["Đã hoàn thành", "Hoàn thành chậm"]:
                        self.pop_entries["Ngày Hoàn Thành:"].set_date(datetime.now())
                    else:
                        self.pop_entries["Ngày Hoàn Thành:"].delete(0, tk.END)
                    self.pop_entries["Ngày Hoàn Thành:"].config(state="disabled")
                entry.bind("<<ComboboxSelected>>", on_pop_status_change)
            elif field in ["Ngày Nhận:", "Ngày Bắt Đầu:", "Ngày Kết Thúc:", "Ngày Hoàn Thành:"]:
                entry = DateEntry(form_frame, width=29, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
                entry.delete(0, tk.END)
                if field == "Ngày Hoàn Thành:":
                    entry.config(state="disabled") 
            elif field in ["Người Xử Lý:", "Báo Cáo:"]:
                entry = ttk.Combobox(form_frame, width=29)
                
                if field == "Người Xử Lý:":
                    if self.current_role == 'super_admin':
                        self.setup_user_combobox(entry, self.all_users_display)
                    elif self.current_role == 'admin':
                        self.setup_user_combobox(entry, self.assignable_by_admin)
                    else:
                        self.setup_user_combobox(entry, self.all_users_display)
                        disp_name = self.user_to_fullname.get(self.current_username, self.current_username)
                        entry.insert(0, disp_name)
                        entry.config(state="disabled") 
                else: 
                    # ĐÃ CHỈNH SỬA: Ô BÁO CÁO Ở TẠO MỚI CŨNG CHỈ CÓ LÃNH ĐẠO
                    self.setup_user_combobox(entry, self.admin_users_display)
                    if self.current_role not in ['admin', 'super_admin'] and self.admin_users_display:
                        entry.insert(0, self.admin_users_display[0])
            else:
                entry = ttk.Entry(form_frame, width=31)
                
            entry.grid(row=row, column=col+1, padx=10, pady=5, sticky="w")
            self.pop_entries[field] = entry
            
        file_frame = ttk.Frame(form_frame)
        file_frame.grid(row=6, column=0, columnspan=4, pady=10, sticky="w", padx=10)
        ttk.Label(file_frame, text="Danh Sách File:").pack(side=tk.LEFT, anchor="n")
        self.pop_listbox_files = tk.Listbox(file_frame, height=3, width=55, selectmode=tk.EXTENDED)
        self.pop_listbox_files.pack(side=tk.LEFT, padx=10)
        btn_box = ttk.Frame(file_frame)
        btn_box.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_box, text="➕ Chọn File", command=self.pop_handle_upload).pack(fill=tk.X, pady=1)
        ttk.Button(btn_box, text="❌ Xóa File", command=self.pop_handle_remove_file).pack(fill=tk.X, pady=1)
        btn_frame = ttk.Frame(self.popup)
        btn_frame.pack(fill=tk.X, padx=15, pady=15)
        ttk.Button(btn_frame, text="💾 LƯU VÀO HỆ THỐNG", command=self.pop_handle_save).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="Hủy bỏ", command=self.popup.destroy).pack(side=tk.RIGHT, padx=5)

    def pop_handle_upload(self):
        file_paths = filedialog.askopenfilenames(title="Chọn file", filetypes=[("Tất cả", "*.*")])
        for path in file_paths:
            if path not in self.popup_attached_files:
                self.popup_attached_files.append(path)
                self.pop_listbox_files.insert(tk.END, os.path.basename(path))

    def pop_handle_remove_file(self):
        selected = self.pop_listbox_files.curselection()
        if not selected: return
        for idx in reversed(selected):
            self.pop_listbox_files.delete(idx)
            self.popup_attached_files.pop(idx)

    def pop_handle_save(self):
        ten_cv = self.pop_txt_ten.get("1.0", tk.END).strip()
        noi_dung = self.pop_txt_nd.get("1.0", tk.END).strip()
        if not ten_cv:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập 'Tên Công Việc'!")
            return

        if self.current_role in ['admin', 'super_admin']:
            disp_name = self.user_to_fullname.get(self.current_username, self.current_username)
            if not ten_cv.startswith("⭐"):
                ten_cv = f"⭐[{disp_name}] {ten_cv}"
            elif ten_cv.startswith("⭐") and "⭐[" not in ten_cv:
                clean_name = ten_cv.lstrip("⭐ ").strip()
                ten_cv = f"⭐[{disp_name}] {clean_name}"

        safe_paths = []
        for path in self.popup_attached_files:
            safe_p = self.process_attached_file(path)
            if safe_p: safe_paths.append(safe_p)
        safe_file_path_str = "|".join(safe_paths)
        tai_lieu_dinh_kem_str = " | ".join([os.path.basename(p) for p in safe_paths])
        
        nxl_disp = self.pop_entries["Người Xử Lý:"].get().strip()
        ntk_disp = self.pop_entries["Báo Cáo:"].get().strip()
        nxl_val = self.fullname_to_user.get(nxl_disp, nxl_disp)
        ntk_val = self.fullname_to_user.get(ntk_disp, ntk_disp)
        
        form_data = (
            ten_cv, noi_dung, tai_lieu_dinh_kem_str,
            nxl_val, ntk_val,
            self.pop_entries["Ngày Nhận:"].get().strip(),
            self.pop_entries["Ngày Bắt Đầu:"].get().strip(),
            self.pop_entries["Ngày Kết Thúc:"].get().strip(),
            self.pop_entries["Ngày Hoàn Thành:"].get().strip(),
            self.pop_entries["Trạng Thái:"].get().strip()
        )
        form_data = list(self.auto_correct_status(form_data))
        
        new_status = form_data[9]
        ngay_tam_dung = datetime.now().strftime("%d/%m/%Y") if new_status == "Đang đợi phản hồi" else ""
        
        if self.task_type_var.get() == 0:
            parent_id = 0
        else:
            selected_str = self.cb_parent.get() 
            if not selected_str or "Gõ STT hoặc Tên" in selected_str or selected_str not in self.main_task_mapping:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn Công việc chứa nó một cách hợp lệ!")
                return
            parent_id = self.main_task_mapping.get(selected_str, 0)
            
        full_data = (parent_id,) + tuple(form_data) + (safe_file_path_str, ngay_tam_dung)
        self.db.insert_cong_van(full_data)
        messagebox.showinfo("Thành công", "Đã lưu công việc thành công!")
        self.popup.destroy() 
        self.refresh_table()

    def calculate_hold_logic(self, old_status, new_status, old_ngay_tam_dung, current_ngay_kt):
        today = datetime.now()
        today_str = today.strftime("%d/%m/%Y")
        ngay_tam_dung = old_ngay_tam_dung
        ngay_kt = current_ngay_kt
        
        if new_status == "Đang đợi phản hồi" and old_status != "Đang đợi phản hồi":
            ngay_tam_dung = today_str
        elif old_status == "Đang đợi phản hồi" and new_status != "Đang đợi phản hồi":
            if old_ngay_tam_dung:
                try:
                    hold_date = datetime.strptime(old_ngay_tam_dung, "%d/%m/%Y")
                    days_held = (today - hold_date).days
                    if days_held > 0 and current_ngay_kt:
                        kt_date = datetime.strptime(current_ngay_kt, "%d/%m/%Y")
                        ngay_kt = (kt_date + timedelta(days=days_held)).strftime("%d/%m/%Y")
                except Exception: pass
            ngay_tam_dung = ""
        elif new_status != "Đang đợi phản hồi":
            ngay_tam_dung = ""
            
        return ngay_tam_dung, ngay_kt

    def is_date_after(self, d1_str, d2_str):
        try:
            d1 = datetime.strptime(d1_str, "%d/%m/%Y").date()
            d2 = datetime.strptime(d2_str, "%d/%m/%Y").date()
            return d1 > d2
        except Exception: return False 

    def auto_correct_status(self, form_tuple):
        data_list = list(form_tuple)
        ngay_kt = data_list[7] 
        ngay_ht = data_list[8] 
        trang_thai = data_list[9] 
        today_str = datetime.now().strftime("%d/%m/%Y")
        
        if trang_thai in ["Đã hoàn thành", "Hoàn thành chậm"]:
            if not ngay_ht:
                ngay_ht = today_str
                data_list[8] = ngay_ht
            if self.is_date_after(ngay_ht, ngay_kt): 
                data_list[9] = "Hoàn thành chậm"
            else: 
                data_list[9] = "Đã hoàn thành"
        elif trang_thai in ["Đang xử lý", "Chậm tiến độ"]:
            data_list[8] = ""
            if self.is_date_after(today_str, ngay_kt): 
                data_list[9] = "Chậm tiến độ"
            else: 
                data_list[9] = "Đang xử lý"
        elif trang_thai == "Đang đợi phản hồi":
            data_list[8] = ""
            
        return tuple(data_list)

    def auto_scan_overdue_tasks(self):
        rows = self.db.get_all_cong_van()
        today_str = datetime.now().strftime("%d/%m/%Y")
        for row in rows:
            record_id = row[0]
            ngay_kt = row[9] 
            trang_thai = row[11]
            if trang_thai == "Đang xử lý" and self.is_date_after(today_str, ngay_kt):
                self.db.update_trang_thai(record_id, "Chậm tiến độ")

    def process_attached_file(self, original_path):
        if not original_path or not os.path.exists(original_path): return ""
        if self.kho_tai_lieu in original_path: return original_path
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_")
        safe_filename = timestamp + os.path.basename(original_path)
        dest_path = os.path.join(self.kho_tai_lieu, safe_filename)
        try:
            shutil.copy2(original_path, dest_path)
            return dest_path 
        except Exception: return original_path

    def handle_upload(self):
        file_paths = filedialog.askopenfilenames(title="Chọn file đính kèm", filetypes=[("Tất cả", "*.*")])
        for path in file_paths:
            if path not in self.current_attached_files:
                self.current_attached_files.append(path)
                self.ui.listbox_files.insert(tk.END, os.path.basename(path))

    def handle_remove_file(self):
        selected = self.ui.listbox_files.curselection()
        if not selected: return
        for idx in reversed(selected):
            self.ui.listbox_files.delete(idx)
            self.current_attached_files.pop(idx)

    def handle_download(self):
        selected = self.ui.listbox_files.curselection()
        if not selected: return
        if len(selected) == 1:
            idx = selected[0]
            target_file = self.current_attached_files[idx]
            if not os.path.exists(target_file): return
            file_ext = os.path.splitext(target_file)[1]
            dest_path = filedialog.asksaveasfilename(defaultextension=file_ext, filetypes=[("File gốc", f"*{file_ext}")], initialfile=os.path.basename(target_file))
            if dest_path:
                shutil.copy2(target_file, dest_path)
                os.startfile(dest_path) 
        else:
            dest_dir = filedialog.askdirectory(title="Chọn THƯ MỤC tải về")
            if dest_dir:
                for idx in selected:
                    target_file = self.current_attached_files[idx]
                    if os.path.exists(target_file):
                        dest_path = os.path.join(dest_dir, os.path.basename(target_file))
                        shutil.copy2(target_file, dest_path)
                os.startfile(dest_dir)

    def get_main_form_data(self):
        ten_cv = self.ui.txt_ten_cv.get("1.0", tk.END).strip()
        noi_dung = self.ui.txt_noi_dung.get("1.0", tk.END).strip()
        filenames = [os.path.basename(p) for p in self.current_attached_files]
        tai_lieu_dinh_kem = " | ".join(filenames)
        
        nxl_disp = self.ui.entries["Người Xử Lý:"].get().strip()
        ntk_disp = self.ui.entries["Báo Cáo:"].get().strip()
        nxl_val = self.fullname_to_user.get(nxl_disp, nxl_disp)
        ntk_val = self.fullname_to_user.get(ntk_disp, ntk_disp)
        
        return (
            ten_cv, noi_dung, tai_lieu_dinh_kem,
            nxl_val, ntk_val,
            self.ui.entries["Ngày Nhận:"].get().strip(),
            self.ui.entries["Ngày Bắt Đầu:"].get().strip(),
            self.ui.entries["Ngày Kết Thúc:"].get().strip(),
            self.ui.entries["Ngày Hoàn Thành:"].get().strip(),
            self.ui.entries["Trạng Thái:"].get().strip()
        )

    def handle_update_click(self):
        if not self.current_selected_id: return
        
        self.db.cursor.execute("SELECT trang_thai, ngay_tam_dung FROM cong_van WHERE id=%s", (self.current_selected_id,))
        old_data = self.db.cursor.fetchone()
        old_status = old_data[0] if old_data else ""
        old_ngay_tam_dung = old_data[1] if (old_data and len(old_data) > 1) else ""
        
        form_data = self.get_main_form_data()
        if not form_data[0]: return
            
        safe_paths = []
        for path in self.current_attached_files:
            safe_p = self.process_attached_file(path)
            if safe_p: safe_paths.append(safe_p)
            
        safe_file_path_str = "|".join(safe_paths)
        tai_lieu_dinh_kem_str = " | ".join([os.path.basename(p) for p in safe_paths])
        data_list = list(form_data)
        data_list[2] = tai_lieu_dinh_kem_str 

        if self.current_role in ['admin', 'super_admin']:
            disp_name = self.user_to_fullname.get(self.current_username, self.current_username)
            if not data_list[0].startswith("⭐"):
                data_list[0] = f"⭐[{disp_name}] {data_list[0]}"
            elif data_list[0].startswith("⭐") and "⭐[" not in data_list[0]:
                clean_name = data_list[0].lstrip("⭐ ").strip()
                data_list[0] = f"⭐[{disp_name}] {clean_name}"
        
        new_ngay_tam_dung, new_ngay_kt = self.calculate_hold_logic(old_status, data_list[9], old_ngay_tam_dung, data_list[7])
        data_list[7] = new_ngay_kt
        
        form_data = self.auto_correct_status(tuple(data_list))
        full_data = (self.current_parent_id,) + form_data + (safe_file_path_str, new_ngay_tam_dung)
        self.db.update_cong_van(self.current_selected_id, full_data)
        messagebox.showinfo("Thành công", "Cập nhật thành công!")
        self.handle_cancel_click() 
        self.refresh_table()

    def handle_delete_click(self):
        if not self.current_selected_id: return
        if messagebox.askyesno("Xác nhận xóa", "Cảnh báo: Hệ thống sẽ XÓA TOÀN BỘ các Việc con/cháu bên trong.\nBạn có chắc chắn muốn xóa?"):
            rows = self.db.get_all_cong_van()
            tree_data, _ = self.build_tree_data(rows)
            ids_to_delete = [self.current_selected_id]
            
            def collect_descendants(p_id):
                if p_id in tree_data:
                    for c in tree_data[p_id]:
                        ids_to_delete.append(c[0])
                        collect_descendants(c[0])
                        
            collect_descendants(self.current_selected_id)
            for d_id in ids_to_delete: self.db.delete_cong_van(d_id) 
            messagebox.showinfo("Thành công", "Đã xóa toàn bộ nhánh thành công!")
            self.handle_cancel_click() 
            self.refresh_table()

    def handle_cancel_click(self):
        self.current_selected_id = None 
        self.current_parent_id = 0
        self.current_attached_files = []
        self.current_is_boss_task = False
        
        self.set_form_state("normal") 
        self.ui.listbox_files.delete(0, tk.END)
        self.ui.txt_ten_cv.delete("1.0", tk.END)
        self.ui.txt_noi_dung.delete("1.0", tk.END)
        for field, entry in self.ui.entries.items():
            if field == "Trạng Thái:": entry.set("Đang xử lý") 
            else: 
                if field == "Ngày Hoàn Thành:":
                    entry.config(state="normal")
                    entry.delete(0, tk.END)
                    entry.config(state="disabled")
                else:
                    entry.set('') if isinstance(entry, ttk.Combobox) else entry.delete(0, tk.END)
                    
        self.set_form_state("disabled") 
        self.ui.btn_edit.config(state=tk.DISABLED)

    def handle_tree_click(self, event):
        selected_item = self.ui.tree.selection()
        if selected_item:
            self.handle_cancel_click() 
            self.current_selected_id = int(selected_item[0]) 
            self.db.cursor.execute('''
                SELECT id, parent_id, ten_cong_viec, noi_dung, tai_lieu_dinh_kem, 
                       nguoi_xu_ly, nguoi_trinh_ky, ngay_nhan, ngay_bat_dau, ngay_ket_thuc, 
                       ngay_hoan_thanh, trang_thai, file_path 
                FROM cong_van WHERE id=%s
            ''', (self.current_selected_id,))
            full_data = self.db.cursor.fetchone()
            
            if full_data:
                self.current_is_boss_task = str(full_data[2]).startswith("⭐")
                
                self.set_form_state("normal") 
                self.current_parent_id = full_data[1] 
                self.ui.txt_ten_cv.insert(tk.END, str(full_data[2]) if str(full_data[2]) != "None" else "")
                self.ui.txt_noi_dung.insert(tk.END, str(full_data[3]) if str(full_data[3]) != "None" else "")
                
                nxl_disp = self.user_to_fullname.get(full_data[5], full_data[5]) if full_data[5] else ""
                ntk_disp = self.user_to_fullname.get(full_data[6], full_data[6]) if full_data[6] else ""
                
                self.ui.entries["Người Xử Lý:"].set(nxl_disp)
                self.ui.entries["Báo Cáo:"].set(ntk_disp)
                
                self.ui.entries["Ngày Nhận:"].delete(0, tk.END)
                self.ui.entries["Ngày Nhận:"].insert(0, str(full_data[7]) if str(full_data[7]) != "None" else "")
                self.ui.entries["Ngày Bắt Đầu:"].delete(0, tk.END)
                self.ui.entries["Ngày Bắt Đầu:"].insert(0, str(full_data[8]) if str(full_data[8]) != "None" else "")
                self.ui.entries["Ngày Kết Thúc:"].delete(0, tk.END)
                self.ui.entries["Ngày Kết Thúc:"].insert(0, str(full_data[9]) if str(full_data[9]) != "None" else "")
                
                self.ui.entries["Ngày Hoàn Thành:"].config(state="normal")
                self.ui.entries["Ngày Hoàn Thành:"].delete(0, tk.END)
                if full_data[10] and str(full_data[10]) != "None":
                    self.ui.entries["Ngày Hoàn Thành:"].insert(0, str(full_data[10]))
                self.ui.entries["Ngày Hoàn Thành:"].config(state="disabled")
                
                self.ui.entries["Trạng Thái:"].set(str(full_data[11]))
                
                file_path_str = str(full_data[12]) if full_data[12] and full_data[12] != "None" else ""
                self.current_attached_files = file_path_str.split("|") if file_path_str else []
                for p in self.current_attached_files:
                    if p.strip(): self.ui.listbox_files.insert(tk.END, os.path.basename(p))
                
                self.set_form_state("disabled") 
                self.ui.btn_edit.config(state=tk.NORMAL) 

    def handle_export_excel(self):
        dest_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile="Danh_Sach_Cong_Viec.xlsx")
        if dest_path:
            try:
                rows = self.db.get_all_for_excel()
                tree_data, _ = self.build_tree_data(rows)
                progress_map = {}
                for r in rows:
                    if r[0] in tree_data and len(tree_data[r[0]]) > 0:
                        total = len(tree_data[r[0]])
                        comp = sum(1 for c in tree_data[r[0]] if c[11] in ["Đã hoàn thành", "Hoàn thành chậm"])
                        progress_map[r[0]] = f"{int((comp/total)*100)}%"
                    else:
                        progress_map[r[0]] = ""
                
                rows_cleaned = []
                def append_excel_node(p_id, level):
                    if p_id not in tree_data: return
                    for r in tree_data[p_id]:
                        has_perm = (self.current_role in ['admin', 'super_admin']) or (self.current_username.lower() in str(r[5] or "").lower()) or (self.current_username.lower() in str(r[6] or "").lower())
                        if not has_perm: continue

                        base = list(r[:-1]) 
                        if level > 0: base[2] = f"{'   ' * level}↳ {base[2]}" 
                        
                        base[5] = self.user_to_fullname.get(base[5], base[5]) if base[5] else ""
                        base[6] = self.user_to_fullname.get(base[6], base[6]) if base[6] else ""
                        
                        base.append(progress_map.get(r[0], ""))
                        rows_cleaned.append(base)
                        append_excel_node(r[0], level + 1)
                        
                append_excel_node(0, 0)
                columns = ["ID", "ID Việc Chính", "Tên Công Việc", "Nội Dung", "Tài Liệu", "Người XL", "Báo Cáo", "Ngày Nhận", "Bắt Đầu", "Kết Thúc", "Hoàn Thành", "Trạng Thái", "Tiến Độ (%)"]
                df = pd.DataFrame(rows_cleaned, columns=columns)
                df.to_excel(dest_path, index=False, engine='openpyxl')
                messagebox.showinfo("Thành công", f"Đã xuất báo cáo Excel tại:\n{dest_path}")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi xuất Excel:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppController(root)
    root.mainloop()