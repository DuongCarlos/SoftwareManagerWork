import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import shutil
import os
import sys 
from datetime import datetime, timedelta
import pandas as pd 
from tkcalendar import DateEntry 
from gui import MainWindowUI
from database import Database

class AppController:
    def __init__(self, root):
        self.root = root
        self.ui = MainWindowUI(root)
        
        if getattr(sys, 'frozen', False):
            self.app_dir = os.path.dirname(sys.executable)
        else:
            self.app_dir = os.path.dirname(os.path.abspath(__file__))

        self.db_folder = os.path.join(self.app_dir, "Database")
        self.kho_tai_lieu = os.path.join(self.app_dir, "Kho_Tai_Lieu")
        
        if not os.path.exists(self.db_folder):
            os.makedirs(self.db_folder)
        if not os.path.exists(self.kho_tai_lieu):
            os.makedirs(self.kho_tai_lieu)
            
        self.current_db_path = os.path.join(self.db_folder, "congvan_v6.db") 
        self.db = Database(self.current_db_path)
        
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
        self.ui.btn_open_db.config(command=self.handle_open_db)
        self.ui.btn_export_excel.config(command=self.handle_export_excel)
        
        self.ui.btn_search.config(command=self.handle_search)
        self.ui.btn_clear_search.config(command=self.handle_clear_search)
        self.ui.btn_schedule.config(command=self.open_schedule_popup) 

        self.ui.tree.bind("<<TreeviewSelect>>", self.handle_tree_click)
        self.ui.entries["Trạng Thái:"].bind("<<ComboboxSelected>>", self.on_main_status_change)

        self.update_window_title()
        self.auto_spawn_scheduled_tasks()
        self.auto_scan_overdue_tasks()
        
        self.set_form_state("disabled")
        self.ui.btn_edit.config(state=tk.DISABLED)
        self.refresh_table()

    # ================= 1. LOGIC TÌM KIẾM TRÊN BẢNG VÀ TÍNH % TỰ ĐỘNG =================
    def handle_search(self):
        kw = self.ui.entry_search_date.get().strip()
        status = self.ui.combo_search_status.get()
        self.refresh_table(search_kw=kw, search_status=status)

    def handle_clear_search(self):
        self.ui.entry_search_date.delete(0, tk.END)
        self.ui.combo_search_status.set("Tất cả")
        self.refresh_table()

    def get_tag_color(self, trang_thai):
        if trang_thai == "Đã hoàn thành": return 'hoan_thanh'
        elif trang_thai == "Đang xử lý": return 'dang_xu_ly'
        elif trang_thai == "Chậm tiến độ": return 'cham_tien_do'
        elif trang_thai == "Hoàn thành chậm": return 'hoan_thanh_cham'
        return ''

    # Hàm vệ tinh: Tính tự động phần trăm (%)
    def get_task_progress_map(self, rows):
        progress_map = {}
        subtask_stats = {}
        # Quét lần 1: Đếm số lượng Việc con của từng Việc chính
        for row in rows:
            if row[1] != 0: # Đây là Việc con
                p_id = row[1]
                if p_id not in subtask_stats:
                    subtask_stats[p_id] = {'total': 0, 'completed': 0}
                subtask_stats[p_id]['total'] += 1
                if row[11] in ["Đã hoàn thành", "Hoàn thành chậm"]:
                    subtask_stats[p_id]['completed'] += 1

        # Quét lần 2: Đúc kết Tỷ lệ % cho Việc chính
        for row in rows:
            t_id = row[0]
            if row[1] == 0: # Đây là Việc chính
                if t_id in subtask_stats: # Nếu nó có Việc con
                    total = subtask_stats[t_id]['total']
                    comp = subtask_stats[t_id]['completed']
                    pct = int((comp / total) * 100)
                    progress_map[t_id] = f"{pct}%"
                else: # Nếu nó KHÔNG có Việc con
                    if row[11] in ["Đã hoàn thành", "Hoàn thành chậm"]:
                        progress_map[t_id] = "100%"
                    else:
                        progress_map[t_id] = "0%"
            else: # Việc con thì để trống ô %
                progress_map[t_id] = ""
                
        return progress_map

    def refresh_table(self, search_kw="", search_status="Tất cả"):
        for item in self.ui.tree.get_children():
            self.ui.tree.delete(item)
            
        rows = self.db.get_all_cong_van()
        
        # Gọi hàm tính toán % thông minh (Chạy ngầm, không sửa DB)
        progress_map = self.get_task_progress_map(rows)
        
        valid_parent_ids = set()
        valid_node_ids = set()
        is_searching = bool(search_kw) or (search_status != "Tất cả")
        
        if is_searching:
            for row in rows:
                match_kw = True
                if search_kw:
                    d1 = str(row[7]) if row[7] else ""
                    d2 = str(row[8]) if row[8] else ""
                    d3 = str(row[9]) if row[9] else ""
                    d4 = str(row[10]) if row[10] else ""
                    if search_kw not in d1 and search_kw not in d2 and search_kw not in d3 and search_kw not in d4:
                        match_kw = False
                
                match_status = True
                if search_status != "Tất cả":
                    if str(row[11]) != search_status:
                        match_status = False
                        
                if match_kw and match_status:
                    valid_node_ids.add(row[0]) 
                    if row[1] == 0:
                        valid_parent_ids.add(row[0]) 
                    else:
                        valid_parent_ids.add(row[1]) 
        
        stt_chinh = 1 
        for row in rows:
            if row[1] == 0: 
                if is_searching and row[0] not in valid_parent_ids:
                    stt_chinh += 1 
                    continue
                    
                tag = self.get_tag_color(row[11])
                ty_le = progress_map.get(row[0], "")
                display_values = (stt_chinh, row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], ty_le)
                self.ui.tree.insert("", tk.END, iid=str(row[0]), values=display_values, tags=(tag,))
                stt_chinh += 1 

        for row in rows:
            if row[1] != 0: 
                if is_searching and row[0] not in valid_node_ids:
                    continue
                    
                tag = self.get_tag_color(row[11])
                ty_le = progress_map.get(row[0], "")
                display_values = ("", f"   ↳ {row[2]}", row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], ty_le)
                try:
                    self.ui.tree.insert(str(row[1]), tk.END, iid=str(row[0]), values=display_values, tags=(tag,))
                    self.ui.tree.item(str(row[1]), open=True) 
                except tk.TclError:
                    pass 


    # ================= 2. LOGIC AUTO SPAWN TASKS ĐỊNH KỲ =================
    def auto_spawn_scheduled_tasks(self):
        today = datetime.now()
        today_str = today.strftime("%d/%m/%Y")
        weekdays = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ nhật"]
        today_weekday = weekdays[today.weekday()]
        today_day = str(today.day)
        
        tasks = self.db.get_all_viec_dinh_ky()
        spawned_count = 0
        
        for t in tasks:
            t_id, t_ten, t_nd, t_nxl, t_ntk, t_loai, t_chitiet, t_ngaytao, t_tldk, t_songay, t_ngaybd, t_filepath = t
            if t_ngaybd and t_ngaybd != "":
                if self.is_date_after(t_ngaybd, today_str):
                    continue 
            should_create = False
            if str(t_ngaytao) != today_str:
                if t_loai == "Hàng ngày":
                    should_create = True
                elif t_loai == "Hàng tuần" and str(t_chitiet) == today_weekday:
                    should_create = True
                elif t_loai == "Hàng tháng" and str(t_chitiet) == today_day:
                    should_create = True
            if should_create:
                try: days = int(t_songay)
                except: days = 0
                ngay_kt_str = (today + timedelta(days=days)).strftime("%d/%m/%Y")
                new_cv = (0, f"[Định kỳ] {t_ten}", str(t_nd), str(t_tldk), str(t_nxl), str(t_ntk), today_str, today_str, ngay_kt_str, "", "Đang xử lý", str(t_filepath))
                self.db.insert_cong_van(new_cv)
                self.db.update_ngay_tao_gan_nhat(t_id, today_str)
                spawned_count += 1
                
        if spawned_count > 0:
            messagebox.showinfo("Thông báo", f"Hệ thống vừa tự động sinh ra {spawned_count} công việc định kỳ cho hôm nay!")

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
        self.sch_nxl = ttk.Entry(form_frame, width=25)
        self.sch_nxl.grid(row=0, column=4, padx=5, pady=5, sticky="w")
        ttk.Label(form_frame, text="Người Trình Ký:").grid(row=0, column=5, padx=5, pady=5, sticky="nw")
        self.sch_ntk = ttk.Entry(form_frame, width=25)
        self.sch_ntk.grid(row=0, column=6, padx=5, pady=5, sticky="w")
        
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
        self.sch_nxl.delete(0, tk.END)
        self.sch_ntk.delete(0, tk.END)
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
                self.sch_nxl.insert(0, t[3])
                self.sch_ntk.insert(0, t[4])
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
        nxl = self.sch_nxl.get().strip()
        ntk = self.sch_ntk.get().strip()
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
        if not data:
            messagebox.showwarning("Lỗi", "Nhập tên công việc!")
            return
        ten, nd, nxl, ntk, loai, chitiet, tldk, songay, ngaybd, fpath = data
        self.db.insert_viec_dinh_ky((ten, nd, nxl, ntk, loai, chitiet, "", tldk, songay, ngaybd, fpath))
        messagebox.showinfo("Thành công", "Đã thiết lập tự động hóa!")
        self.sch_handle_cancel()
        self.refresh_schedule_table()

    def update_schedule(self):
        if not self.current_sch_id: return
        data = self.get_sch_form_data()
        if not data:
            messagebox.showwarning("Lỗi", "Nhập tên công việc!")
            return
        ten, nd, nxl, ntk, loai, chitiet, tldk, songay, ngaybd, fpath = data
        tasks = self.db.get_all_viec_dinh_ky()
        ngaytao = ""
        for t in tasks:
            if t[0] == self.current_sch_id:
                ngaytao = t[7]
                break
        self.db.update_viec_dinh_ky(self.current_sch_id, (ten, nd, nxl, ntk, loai, chitiet, ngaytao, tldk, songay, ngaybd, fpath))
        messagebox.showinfo("Thành công", "Đã cập nhật mẫu định kỳ!")
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
        for row in self.tree_sch.get_children():
            self.tree_sch.delete(row)
        for t in self.db.get_all_viec_dinh_ky():
            self.tree_sch.insert("", tk.END, iid=str(t[0]), values=(t[0], t[1], t[5], t[6], t[10], t[9], t[3], t[7]))


    # ================= 3. LOGIC GIAO DIỆN CHÍNH / THÊM MỚI =================

    def set_form_state(self, state):
        text_state = tk.NORMAL if state == "normal" else tk.DISABLED
        self.ui.txt_ten_cv.config(state=text_state)
        self.ui.txt_noi_dung.config(state=text_state)
        for field, entry in self.ui.entries.items():
            if field == "Trạng Thái:":
                entry.config(state="readonly" if state == "normal" else "disabled")
            else:
                entry.config(state=state)
        btn_state = tk.NORMAL if state == "normal" else tk.DISABLED
        self.ui.btn_upload.config(state=btn_state)
        self.ui.btn_remove_file.config(state=btn_state)
        self.ui.btn_update.config(state=btn_state) 

    def handle_edit_click(self):
        if not self.current_selected_id:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn 1 công việc để chỉnh sửa!")
            return
        self.set_form_state("normal")
        self.ui.btn_edit.config(state=tk.DISABLED)

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
        
        main_tasks = self.db.get_main_tasks()
        self.main_task_mapping = {f"[STT: {i+1}] {t[1]}": t[0] for i, t in enumerate(main_tasks)}
        self.all_task_strs = list(self.main_task_mapping.keys())
        
        ttk.Label(type_frame, text="Thuộc công việc:").grid(row=0, column=2, padx=10)
        self.cb_parent = ttk.Combobox(type_frame, state="disabled", width=55)
        self.cb_parent.grid(row=0, column=3, padx=5, pady=10)
        
        PLACEHOLDER = "Gõ STT hoặc Tên rồi nhấn Enter..."
        def toggle_parent_cb():
            if self.task_type_var.get() == 1:
                self.cb_parent.config(state="normal") 
                self.cb_parent['values'] = self.all_task_strs
                self.cb_parent.set(PLACEHOLDER) 
            else:
                self.cb_parent.set("")
                self.cb_parent.config(state="disabled")

        def on_focus_in(event):
            if self.cb_parent.get() == PLACEHOLDER:
                self.cb_parent.set("") 

        def on_focus_out(event):
            if self.cb_parent.get().strip() == "" and self.task_type_var.get() == 1:
                self.cb_parent.set(PLACEHOLDER) 

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
                    for s in self.all_task_strs:
                        if s.lower().startswith(exact_stt):
                            matches.append(s)
                if not matches:
                    matches = [s for s in self.all_task_strs if typed in s.lower()]
                self.cb_parent['values'] = matches
                if not matches:
                    self.cb_parent.set("Không tìm thấy!")
            self.cb_parent.event_generate('<Down>')
            
        self.cb_parent.bind("<Return>", trigger_search)
        ttk.Radiobutton(type_frame, text="Công việc chính", variable=self.task_type_var, value=0, command=toggle_parent_cb).grid(row=0, column=0, padx=15, pady=10)
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
        fields = ["Người Xử Lý:", "Người Trình ký VB:", "Ngày Nhận:", "Ngày Bắt Đầu:", "Ngày Kết Thúc:", "Ngày Hoàn Thành:", "Trạng Thái:"]
        for idx, field in enumerate(fields):
            row = 2 + (idx // 2)
            col = (idx % 2) * 2
            ttk.Label(form_frame, text=field).grid(row=row, column=col, sticky="w", padx=10, pady=5)
            if field == "Trạng Thái:":
                entry = ttk.Combobox(form_frame, width=28, state="readonly", values=["Đang xử lý", "Đã hoàn thành", "Chậm tiến độ", "Hoàn thành chậm"])
                entry.set("Đang xử lý")
                def on_pop_status_change(event):
                    if self.pop_entries["Trạng Thái:"].get() in ["Đã hoàn thành", "Hoàn thành chậm"]:
                        self.pop_entries["Ngày Hoàn Thành:"].set_date(datetime.now())
                entry.bind("<<ComboboxSelected>>", on_pop_status_change)
            elif field in ["Ngày Nhận:", "Ngày Bắt Đầu:", "Ngày Kết Thúc:", "Ngày Hoàn Thành:"]:
                entry = DateEntry(form_frame, width=29, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
                entry.delete(0, tk.END)
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
        safe_paths = []
        for path in self.popup_attached_files:
            safe_p = self.process_attached_file(path)
            if safe_p: safe_paths.append(safe_p)
        safe_file_path_str = "|".join(safe_paths)
        tai_lieu_dinh_kem_str = " | ".join([os.path.basename(p) for p in safe_paths])
        form_data = (
            ten_cv, noi_dung, tai_lieu_dinh_kem_str,
            self.pop_entries["Người Xử Lý:"].get().strip(),
            self.pop_entries["Người Trình ký VB:"].get().strip(),
            self.pop_entries["Ngày Nhận:"].get().strip(),
            self.pop_entries["Ngày Bắt Đầu:"].get().strip(),
            self.pop_entries["Ngày Kết Thúc:"].get().strip(),
            self.pop_entries["Ngày Hoàn Thành:"].get().strip(),
            self.pop_entries["Trạng Thái:"].get().strip()
        )
        form_data = self.auto_correct_status(form_data)
        if self.task_type_var.get() == 0:
            parent_id = 0
        else:
            selected_str = self.cb_parent.get().strip()
            if not selected_str or "Gõ STT hoặc Tên" in selected_str or selected_str not in self.main_task_mapping:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn Việc chính hợp lệ!")
                return
            parent_id = self.main_task_mapping.get(selected_str, 0)
            
        full_data = (parent_id,) + form_data + (safe_file_path_str,)
        self.db.insert_cong_van(full_data)
        messagebox.showinfo("Thành công", "Đã lưu công việc thành công!")
        self.popup.destroy() 
        self.refresh_table()

    def update_window_title(self):
        db_name = os.path.basename(self.current_db_path)
        self.root.title(f"Phần mềm Quản lý Công việc - Mở DB: [{db_name}]")

    def is_date_after(self, d1_str, d2_str):
        try:
            d1 = datetime.strptime(d1_str, "%d/%m/%Y").date()
            d2 = datetime.strptime(d2_str, "%d/%m/%Y").date()
            return d1 > d2
        except Exception:
            return False 

    def auto_scan_overdue_tasks(self):
        rows = self.db.get_all_cong_van()
        today_str = datetime.now().strftime("%d/%m/%Y")
        for row in rows:
            record_id = row[0]
            ngay_kt = row[9] 
            trang_thai = row[11]
            if trang_thai == "Đang xử lý" and self.is_date_after(today_str, ngay_kt):
                self.db.update_trang_thai(record_id, "Chậm tiến độ")

    def auto_correct_status(self, form_tuple):
        data_list = list(form_tuple)
        ngay_kt = data_list[7] 
        ngay_ht = data_list[8] 
        trang_thai = data_list[9] 
        today_str = datetime.now().strftime("%d/%m/%Y")
        
        if trang_thai in ["Đã hoàn thành", "Hoàn thành chậm"]:
            if self.is_date_after(ngay_ht, ngay_kt):
                data_list[9] = "Hoàn thành chậm"
            else:
                data_list[9] = "Đã hoàn thành"
        elif trang_thai in ["Đang xử lý", "Chậm tiến độ"]:
            if self.is_date_after(today_str, ngay_kt):
                data_list[9] = "Chậm tiến độ"
            else:
                data_list[9] = "Đang xử lý"
        return tuple(data_list)

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

    def on_main_status_change(self, event):
        val = self.ui.entries["Trạng Thái:"].get()
        if val in ["Đã hoàn thành", "Hoàn thành chậm"]:
            self.ui.entries["Ngày Hoàn Thành:"].set_date(datetime.now())

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
        return (
            ten_cv, noi_dung, tai_lieu_dinh_kem,
            self.ui.entries["Người Xử Lý:"].get().strip(),
            self.ui.entries["Người Trình ký VB:"].get().strip(),
            self.ui.entries["Ngày Nhận:"].get().strip(),
            self.ui.entries["Ngày Bắt Đầu:"].get().strip(),
            self.ui.entries["Ngày Kết Thúc:"].get().strip(),
            self.ui.entries["Ngày Hoàn Thành:"].get().strip(),
            self.ui.entries["Trạng Thái:"].get().strip()
        )

    def handle_update_click(self):
        if not self.current_selected_id: return
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
        form_data = self.auto_correct_status(tuple(data_list))
        
        full_data = (self.current_parent_id,) + form_data + (safe_file_path_str,)
        self.db.update_cong_van(self.current_selected_id, full_data)
        messagebox.showinfo("Thành công", "Cập nhật thành công!")
        self.handle_cancel_click() 
        self.refresh_table()

    def handle_delete_click(self):
        if not self.current_selected_id: return
        if messagebox.askyesno("Xác nhận xóa", "Xóa công việc này (và cả việc con nếu có)?"):
            self.db.delete_cong_van(self.current_selected_id)
            self.handle_cancel_click() 
            self.refresh_table()

    def handle_cancel_click(self):
        self.current_selected_id = None 
        self.current_parent_id = 0
        self.current_attached_files = []
        self.set_form_state("normal") 
        self.ui.listbox_files.delete(0, tk.END)
        self.ui.txt_ten_cv.delete("1.0", tk.END)
        self.ui.txt_noi_dung.delete("1.0", tk.END)
        for field, entry in self.ui.entries.items():
            if field == "Trạng Thái:":
                entry.set("Đang xử lý") 
            else:
                entry.delete(0, tk.END)
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
                FROM cong_van WHERE id=?
            ''', (self.current_selected_id,))
            full_data = self.db.cursor.fetchone()
            
            if full_data:
                self.set_form_state("normal") 
                self.current_parent_id = full_data[1] 
                self.ui.txt_ten_cv.insert(tk.END, str(full_data[2]) if str(full_data[2]) != "None" else "")
                self.ui.txt_noi_dung.insert(tk.END, str(full_data[3]) if str(full_data[3]) != "None" else "")
                self.ui.entries["Người Xử Lý:"].insert(0, str(full_data[5]) if str(full_data[5]) != "None" else "")
                self.ui.entries["Người Trình ký VB:"].insert(0, str(full_data[6]) if str(full_data[6]) != "None" else "")
                self.ui.entries["Ngày Nhận:"].delete(0, tk.END)
                self.ui.entries["Ngày Nhận:"].insert(0, str(full_data[7]) if str(full_data[7]) != "None" else "")
                self.ui.entries["Ngày Bắt Đầu:"].delete(0, tk.END)
                self.ui.entries["Ngày Bắt Đầu:"].insert(0, str(full_data[8]) if str(full_data[8]) != "None" else "")
                self.ui.entries["Ngày Kết Thúc:"].delete(0, tk.END)
                self.ui.entries["Ngày Kết Thúc:"].insert(0, str(full_data[9]) if str(full_data[9]) != "None" else "")
                self.ui.entries["Ngày Hoàn Thành:"].delete(0, tk.END)
                self.ui.entries["Ngày Hoàn Thành:"].insert(0, str(full_data[10]) if str(full_data[10]) != "None" else "")
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
                progress_map = self.get_task_progress_map(rows)
                
                rows_cleaned = []
                for row in rows:
                    base = list(row[:-1])
                    base.append(progress_map.get(row[0], ""))
                    rows_cleaned.append(base)
                    
                columns = ["ID", "ID Việc Chính", "Tên Công Việc", "Nội Dung", "Tài Liệu", "Người XL", "Người TK", "Ngày Nhận", "Bắt Đầu", "Kết Thúc", "Hoàn Thành", "Trạng Thái", "Tiến Độ (%)"]
                df = pd.DataFrame(rows_cleaned, columns=columns)
                df.to_excel(dest_path, index=False, engine='openpyxl')
                messagebox.showinfo("Thành công", f"Đã xuất báo cáo Excel tại:\n{dest_path}")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi xuất Excel:\n{e}")

    def handle_open_db(self):
        file_path = filedialog.askopenfilename(filetypes=[("SQLite DB", "*.db"), ("Tất cả", "*.*")])
        if file_path:
            self.db.close() 
            self.current_db_path = file_path
            self.db = Database(self.current_db_path) 
            self.update_window_title()
            self.refresh_table()

if __name__ == "__main__":
    root = tk.Tk()
    app = AppController(root)
    root.mainloop()