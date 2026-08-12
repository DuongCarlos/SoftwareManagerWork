import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import shutil
import os
import sys 
from datetime import datetime
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
        
        self.ui.tree.bind("<<TreeviewSelect>>", self.handle_tree_click)

        self.update_window_title()
        self.auto_scan_overdue_tasks()
        
        self.set_form_state("disabled")
        self.ui.btn_edit.config(state=tk.DISABLED)
        
        self.refresh_table()

    # ================= LOGIC TÌM KIẾM TRÊN BẢNG =================
    
    def handle_search(self):
        kw = self.ui.entry_search_date.get().strip()
        self.refresh_table(search_kw=kw)

    def handle_clear_search(self):
        self.ui.entry_search_date.delete(0, tk.END)
        self.refresh_table()

    # ================= LOGIC KHÓA/MỞ FORM =================

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

    # ================= LOGIC POP-UP THÊM MỚI =================

    def open_add_popup(self):
        self.popup = tk.Toplevel(self.root)
        self.popup.title("Thêm Mới Công Việc")
        self.popup.geometry("900x700") 
        self.popup.transient(self.root) 
        self.popup.grab_set() 
        
        self.popup_attached_files = []
        
        # Nhãn đã rút gọn gọn gàng
        type_frame = ttk.LabelFrame(self.popup, text="Loại Công Việc")
        type_frame.pack(fill=tk.X, padx=15, pady=10)
        
        self.task_type_var = tk.IntVar(value=0) 
        
        main_tasks = self.db.get_main_tasks()
        self.main_task_mapping = {f"[STT: {i+1}] {t[1]}": t[0] for i, t in enumerate(main_tasks)}
        self.all_task_strs = list(self.main_task_mapping.keys())
        
        ttk.Label(type_frame, text="Thuộc công việc:").grid(row=0, column=2, padx=10)
        
        self.cb_parent = ttk.Combobox(type_frame, state="disabled", width=55)
        self.cb_parent.grid(row=0, column=3, padx=5, pady=10)
        
        # Khai báo dòng Placeholder
        PLACEHOLDER = "Gõ STT hoặc Tên rồi nhấn Enter..."
        
        def toggle_parent_cb():
            if self.task_type_var.get() == 1:
                self.cb_parent.config(state="normal") 
                self.cb_parent['values'] = self.all_task_strs
                self.cb_parent.set(PLACEHOLDER) # Tự động điền chữ gợi ý
            else:
                self.cb_parent.set("")
                self.cb_parent.config(state="disabled")

        # LOGIC XỬ LÝ SỰ KIỆN CLICK CHUỘT VÀO / RA
        def on_focus_in(event):
            if self.cb_parent.get() == PLACEHOLDER:
                self.cb_parent.set("") # Xóa mờ đi khi người dùng click vào

        def on_focus_out(event):
            if self.cb_parent.get().strip() == "" and self.task_type_var.get() == 1:
                self.cb_parent.set(PLACEHOLDER) # Hiện lại nếu bỏ trống

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
        fields = ["Người Xử Lý:", "Người Trình ký VB:", "Ngày Nhận:", "Ngày Bắt Đầu:", "Ngày Hoàn Thành:", "Trạng Thái:"]
        for idx, field in enumerate(fields):
            row = 2 + (idx // 2)
            col = (idx % 2) * 2
            ttk.Label(form_frame, text=field).grid(row=row, column=col, sticky="w", padx=10, pady=5)
            
            if field == "Trạng Thái:":
                entry = ttk.Combobox(form_frame, width=28, state="readonly", values=["Đang xử lý", "Đã hoàn thành", "Chậm tiến độ", "Hoàn thành chậm"])
                entry.set("Đang xử lý")
            elif field in ["Ngày Nhận:", "Ngày Bắt Đầu:", "Ngày Hoàn Thành:"]:
                entry = DateEntry(form_frame, width=29, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
            else:
                entry = ttk.Entry(form_frame, width=31)
                
            entry.grid(row=row, column=col+1, padx=10, pady=5, sticky="w")
            self.pop_entries[field] = entry
            
        file_frame = ttk.Frame(form_frame)
        file_frame.grid(row=5, column=0, columnspan=4, pady=10, sticky="w", padx=10)
        
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
        file_paths = filedialog.askopenfilenames(
            title="Chọn file đính kèm",
            filetypes=[("PDF/Word/Excel", "*.pdf *.docx *.doc *.xlsx"), ("Tất cả", "*.*")]
        )
        for path in file_paths:
            if path not in self.popup_attached_files:
                self.popup_attached_files.append(path)
                self.pop_listbox_files.insert(tk.END, os.path.basename(path))

    def pop_handle_remove_file(self):
        selected = self.pop_listbox_files.curselection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Hãy chọn file để xóa!")
            return
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
            self.pop_entries["Ngày Hoàn Thành:"].get().strip(),
            self.pop_entries["Trạng Thái:"].get().strip()
        )
        form_data = self.auto_correct_status(form_data)
        
        if self.task_type_var.get() == 0:
            parent_id = 0
        else:
            selected_str = self.cb_parent.get().strip()
            
            # KIỂM TRA: Tránh việc người dùng bấm Lưu ngay khi ô vẫn còn chữ Gợi ý
            if not selected_str or selected_str == "Gõ STT hoặc Tên rồi nhấn Enter..." or selected_str not in self.main_task_mapping:
                messagebox.showwarning("Cảnh báo", "Vui lòng CLICK CHỌN một Công việc chính HỢP LỆ từ danh sách xổ xuống!")
                return
                
            parent_id = self.main_task_mapping.get(selected_str, 0)
            
        full_data = (parent_id,) + form_data + (safe_file_path_str,)
        self.db.insert_cong_van(full_data)
        messagebox.showinfo("Thành công", "Đã lưu công việc thành công!")
        self.popup.destroy() 
        self.refresh_table()

    # ================= CÁC HÀM TIỆN ÍCH DÙNG CHUNG =================

    def update_window_title(self):
        db_name = os.path.basename(self.current_db_path)
        self.root.title(f"Phần mềm Quản lý Công việc - Mở DB: [{db_name}]")

    def is_late(self, date_str):
        try:
            date_obj = datetime.strptime(date_str, "%d/%m/%Y").date()
            return datetime.now().date() > date_obj
        except Exception:
            return False 

    def auto_scan_overdue_tasks(self):
        rows = self.db.get_all_cong_van()
        for row in rows:
            record_id = row[0]
            ngay_ht = row[9] 
            trang_thai = row[10]
            if trang_thai == "Đang xử lý" and self.is_late(ngay_ht):
                self.db.update_trang_thai(record_id, "Chậm tiến độ")

    def auto_correct_status(self, form_tuple):
        data_list = list(form_tuple)
        ngay_ht = data_list[7] 
        trang_thai = data_list[8] 
        if trang_thai == "Đã hoàn thành" and self.is_late(ngay_ht):
            data_list[8] = "Hoàn thành chậm"
        elif trang_thai == "Đang xử lý" and self.is_late(ngay_ht):
            data_list[8] = "Chậm tiến độ"
        return tuple(data_list)

    def process_attached_file(self, original_path):
        if not original_path or not os.path.exists(original_path):
            return ""
        if self.kho_tai_lieu in original_path:
            return original_path
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_")
        safe_filename = timestamp + os.path.basename(original_path)
        dest_path = os.path.join(self.kho_tai_lieu, safe_filename)
        try:
            shutil.copy2(original_path, dest_path)
            return dest_path 
        except Exception:
            return original_path

    # ================= LOGIC GIAO DIỆN CHÍNH =================

    def handle_upload(self):
        file_paths = filedialog.askopenfilenames(
            title="Chọn một hoặc nhiều file đính kèm",
            filetypes=[("PDF/Word/Excel", "*.pdf *.docx *.doc *.xlsx"), ("Tất cả", "*.*")]
        )
        if file_paths:
            for path in file_paths:
                if path not in self.current_attached_files:
                    self.current_attached_files.append(path)
                    self.ui.listbox_files.insert(tk.END, os.path.basename(path))

    def handle_remove_file(self):
        selected = self.ui.listbox_files.curselection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Hãy click chọn (bôi đen) file trong danh sách để xóa!")
            return
        for idx in reversed(selected):
            self.ui.listbox_files.delete(idx)
            self.current_attached_files.pop(idx)

    def handle_download(self):
        selected = self.ui.listbox_files.curselection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất 1 file để tải!")
            return
        
        if len(selected) == 1:
            idx = selected[0]
            target_file = self.current_attached_files[idx]
            if not os.path.exists(target_file):
                messagebox.showwarning("Lỗi", "Không tìm thấy file vật lý trong kho.")
                return
            file_ext = os.path.splitext(target_file)[1]
            dest_path = filedialog.asksaveasfilename(
                title="Tải xuống File",
                defaultextension=file_ext,
                filetypes=[("File gốc", f"*{file_ext}")],
                initialfile=os.path.basename(target_file)
            )
            if dest_path:
                try:
                    shutil.copy2(target_file, dest_path)
                    if messagebox.askyesno("Thành công", "Đã tải file thành công!\nBạn có muốn mở file ngay bây giờ?"):
                        os.startfile(dest_path) 
                except Exception as e:
                    messagebox.showerror("Lỗi", f"Không thể tải file:\n{e}")
        else:
            dest_dir = filedialog.askdirectory(title="Chọn THƯ MỤC để tải tất cả file đã chọn về")
            if dest_dir:
                success_count = 0
                error_msgs = []
                for idx in selected:
                    target_file = self.current_attached_files[idx]
                    if os.path.exists(target_file):
                        dest_path = os.path.join(dest_dir, os.path.basename(target_file))
                        try:
                            shutil.copy2(target_file, dest_path)
                            success_count += 1
                        except Exception as e:
                            error_msgs.append(f"Lỗi tải {os.path.basename(target_file)}: {e}")
                    else:
                        error_msgs.append(f"Mất file gốc: {os.path.basename(target_file)}")
                
                if error_msgs:
                    messagebox.showwarning("Hoàn tất có lỗi", f"Đã tải {success_count}/{len(selected)} file.\n\nChi tiết lỗi:\n" + "\n".join(error_msgs))
                else:
                    if messagebox.askyesno("Thành công", f"Đã tải thành công toàn bộ {success_count} file!\nBạn có muốn mở thư mục chứa file lên không?"):
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
            self.ui.entries["Ngày Hoàn Thành:"].get().strip(),
            self.ui.entries["Trạng Thái:"].get().strip()
        )

    def handle_update_click(self):
        if not self.current_selected_id:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn 1 công việc dưới bảng để cập nhật!")
            return
        
        form_data = self.get_main_form_data()
        if not form_data[0]: 
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập 'Tên Công Việc'!")
            return
            
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
        messagebox.showinfo("Thành công", "Đã cập nhật thay đổi thành công!")
        
        self.handle_cancel_click() 
        self.refresh_table()

    def handle_delete_click(self):
        if not self.current_selected_id:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn 1 công việc để xóa!")
            return
        
        msg = "Bạn có chắc muốn xóa công việc này không?"
        if self.current_parent_id == 0:
            msg += "\n(Lưu ý: Các công việc con bên trong cũng sẽ bị xóa theo!)"
            
        if messagebox.askyesno("Xác nhận xóa", msg):
            self.db.delete_cong_van(self.current_selected_id)
            messagebox.showinfo("Thành công", "Đã xóa thành công!")
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
            
            self.db.cursor.execute("SELECT * FROM cong_van WHERE id=?", (self.current_selected_id,))
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
                self.ui.entries["Ngày Hoàn Thành:"].delete(0, tk.END)
                self.ui.entries["Ngày Hoàn Thành:"].insert(0, str(full_data[9]) if str(full_data[9]) != "None" else "")
                self.ui.entries["Trạng Thái:"].set(str(full_data[10]))
                
                file_path_str = str(full_data[11]) if full_data[11] and full_data[11] != "None" else ""
                self.current_attached_files = file_path_str.split("|") if file_path_str else []
                for p in self.current_attached_files:
                    if p.strip():
                        self.ui.listbox_files.insert(tk.END, os.path.basename(p))
                        
                self.set_form_state("disabled") 
                self.ui.btn_edit.config(state=tk.NORMAL) 

    # ================= LOGIC XUẤT DB & EXCEL =================

    def handle_export_excel(self):
        dest_path = filedialog.asksaveasfilename(
            title="Xuất Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="Danh_Sach_Cong_Viec.xlsx"
        )
        if dest_path:
            try:
                rows = self.db.get_all_for_excel()
                rows_cleaned = [row[:-1] for row in rows]
                columns = [
                    "ID", "ID Việc Chính (Nếu =0 là việc chính)", "Tên Công Việc", "Nội Dung Chi Tiết", "Tài Liệu Đính Kèm", 
                    "Người Xử Lý", "Người Trình Ký VB", "Ngày Nhận", "Ngày Bắt Đầu", "Ngày Hoàn Thành", "Trạng Thái"
                ]
                df = pd.DataFrame(rows_cleaned, columns=columns)
                df.to_excel(dest_path, index=False, engine='openpyxl')
                messagebox.showinfo("Thành công", f"Đã xuất báo cáo Excel tại:\n{dest_path}")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi xuất Excel:\n{e}")

    def handle_open_db(self):
        file_path = filedialog.askopenfilename(title="Chọn DB", filetypes=[("SQLite DB", "*.db"), ("Tất cả", "*.*")])
        if file_path:
            try:
                self.db.close() 
                self.current_db_path = file_path
                self.db = Database(self.current_db_path) 
                self.update_window_title()
                self.auto_scan_overdue_tasks()
                self.handle_cancel_click()
                self.refresh_table()
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể mở DB.\nLỗi: {e}")

    def handle_export_db(self):
        dest_path = filedialog.asksaveasfilename(
            title="Sao lưu Cơ sở dữ liệu",
            defaultextension=".db",
            filetypes=[("SQLite DB", "*.db")],
            initialfile="backup_congvan.db"
        )
        if dest_path:
            try:
                self.db.close() 
                shutil.copy(self.current_db_path, dest_path) 
                self.db = Database(self.current_db_path) 
                messagebox.showinfo("Thành công", f"Đã xuất DB ra:\n{dest_path}")
            except Exception as e:
                self.db = Database(self.current_db_path) 
                messagebox.showerror("Lỗi", f"Không thể sao lưu Database:\n{e}")

    def get_tag_color(self, trang_thai):
        if trang_thai == "Đã hoàn thành": return 'hoan_thanh'
        elif trang_thai == "Đang xử lý": return 'dang_xu_ly'
        elif trang_thai == "Chậm tiến độ": return 'cham_tien_do'
        elif trang_thai == "Hoàn thành chậm": return 'hoan_thanh_cham'
        return ''

    def refresh_table(self, search_kw=""):
        for item in self.ui.tree.get_children():
            self.ui.tree.delete(item)
            
        rows = self.db.get_all_cong_van()
        
        valid_parent_ids = set()
        if search_kw:
            for row in rows:
                d1 = str(row[7]) if row[7] else ""
                d2 = str(row[8]) if row[8] else ""
                d3 = str(row[9]) if row[9] else ""
                
                if search_kw in d1 or search_kw in d2 or search_kw in d3:
                    if row[1] == 0:
                        valid_parent_ids.add(row[0]) 
                    else:
                        valid_parent_ids.add(row[1]) 
        
        stt_chinh = 1 
        
        for row in rows:
            if row[1] == 0: 
                if search_kw and row[0] not in valid_parent_ids:
                    stt_chinh += 1 
                    continue
                    
                tag = self.get_tag_color(row[10])
                display_values = (stt_chinh, row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10])
                self.ui.tree.insert("", tk.END, iid=str(row[0]), values=display_values, tags=(tag,))
                stt_chinh += 1 

        for row in rows:
            if row[1] != 0: 
                if search_kw and row[1] not in valid_parent_ids:
                    continue
                    
                tag = self.get_tag_color(row[10])
                display_values = ("", f"   ↳ {row[2]}", row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10])
                try:
                    self.ui.tree.insert(str(row[1]), tk.END, iid=str(row[0]), values=display_values, tags=(tag,))
                    self.ui.tree.item(str(row[1]), open=True) 
                except tk.TclError:
                    pass 

if __name__ == "__main__":
    root = tk.Tk()
    app = AppController(root)
    root.mainloop()