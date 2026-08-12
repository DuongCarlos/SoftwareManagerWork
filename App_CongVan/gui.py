import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry

class MainWindowUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Phần mềm Quản lý Công việc & Công văn")
        self.root.geometry("1300x880") 

        self.setup_top_panel()
        self.setup_bottom_panel()

    def setup_top_panel(self):
        self.top_frame = ttk.LabelFrame(self.root, text="Xem chi tiết & Cập nhật Công Việc")
        self.top_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(self.top_frame, text="Tên Công Việc:").grid(row=0, column=0, sticky="nw", padx=10, pady=5)
        self.txt_ten_cv = tk.Text(self.top_frame, height=2, width=95, wrap=tk.WORD)
        self.txt_ten_cv.grid(row=0, column=1, columnspan=3, padx=10, pady=5, sticky="w")

        ttk.Label(self.top_frame, text="Nội Dung Chi Tiết:").grid(row=1, column=0, sticky="nw", padx=10, pady=5)
        self.txt_noi_dung = tk.Text(self.top_frame, height=3, width=95, wrap=tk.WORD)
        self.txt_noi_dung.grid(row=1, column=1, columnspan=3, padx=10, pady=5, sticky="w")

        self.entries = {}
        fields = ["Người Xử Lý:", "Người Trình ký VB:", "Ngày Nhận:", "Ngày Bắt Đầu:", "Ngày Hoàn Thành:", "Trạng Thái:"]
        
        for idx, field in enumerate(fields):
            row = 2 + (idx // 2)
            col = (idx % 2) * 2
            ttk.Label(self.top_frame, text=field).grid(row=row, column=col, sticky="w", padx=10, pady=5)
            
            if field == "Trạng Thái:":
                entry = ttk.Combobox(self.top_frame, width=32, state="readonly", 
                                     values=["Đang xử lý", "Đã hoàn thành", "Chậm tiến độ", "Hoàn thành chậm"])
                entry.set("Đang xử lý")
            elif field in ["Ngày Nhận:", "Ngày Bắt Đầu:", "Ngày Hoàn Thành:"]:
                entry = DateEntry(self.top_frame, width=33, background='darkblue', 
                                  foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
                entry.delete(0, tk.END)
            else:
                entry = ttk.Entry(self.top_frame, width=35)
                
            entry.grid(row=row, column=col+1, padx=10, pady=5, sticky="w")
            self.entries[field] = entry

        file_frame = ttk.Frame(self.top_frame)
        file_frame.grid(row=5, column=0, columnspan=4, pady=10, sticky="w", padx=10)
        
        ttk.Label(file_frame, text="Danh Sách File:").pack(side=tk.LEFT, anchor="n")
        self.listbox_files = tk.Listbox(file_frame, height=3, width=65, selectmode=tk.EXTENDED)
        self.listbox_files.pack(side=tk.LEFT, padx=10)
        
        btn_box = ttk.Frame(file_frame)
        btn_box.pack(side=tk.LEFT, padx=5)
        self.btn_upload = ttk.Button(btn_box, text="➕ Bổ sung File")
        self.btn_upload.pack(fill=tk.X, pady=1)
        self.btn_download = ttk.Button(btn_box, text="⬇️ Tải xuống File đang chọn")
        self.btn_download.pack(fill=tk.X, pady=1)
        self.btn_remove_file = ttk.Button(btn_box, text="❌ Xóa File khỏi danh sách")
        self.btn_remove_file.pack(fill=tk.X, pady=1)

        self.btn_frame = ttk.Frame(self.top_frame)
        self.btn_frame.grid(row=6, column=0, columnspan=4, pady=10, sticky="ew")

        # Đã bỏ chữ (Pop-up)
        self.btn_add_new = ttk.Button(self.btn_frame, text="➕ Thêm Mới")
        self.btn_add_new.pack(side=tk.LEFT, padx=5)

        self.btn_edit = ttk.Button(self.btn_frame, text="✏️ Bật Chỉnh Sửa")
        self.btn_edit.pack(side=tk.LEFT, padx=5)

        self.btn_update = ttk.Button(self.btn_frame, text="💾 Lưu Cập Nhật")
        self.btn_update.pack(side=tk.LEFT, padx=5)

        self.btn_delete = ttk.Button(self.btn_frame, text="🗑️ Xóa")
        self.btn_delete.pack(side=tk.LEFT, padx=5)

        self.btn_cancel = ttk.Button(self.btn_frame, text="❌ Hủy Chọn / Reset Form")
        self.btn_cancel.pack(side=tk.LEFT, padx=5)

        self.btn_export_excel = ttk.Button(self.btn_frame, text="📊 Xuất Excel")
        self.btn_export_excel.pack(side=tk.RIGHT, padx=5)

        self.btn_open_db = ttk.Button(self.btn_frame, text="📂 Mở DB khác")
        self.btn_open_db.pack(side=tk.RIGHT, padx=5)

    def setup_bottom_panel(self):
        self.bottom_frame = ttk.LabelFrame(self.root, text="Danh sách Công việc")
        self.bottom_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        search_frame = ttk.Frame(self.bottom_frame)
        search_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(search_frame, text="🔍 Tìm kiếm theo Ngày/Tháng (VD: 12/08/2026 hoặc 08/2026):").pack(side=tk.LEFT, padx=5)
        
        self.entry_search_date = ttk.Entry(search_frame, width=20)
        self.entry_search_date.pack(side=tk.LEFT, padx=5)
        
        self.btn_search = ttk.Button(search_frame, text="Lọc Dữ Liệu")
        self.btn_search.pack(side=tk.LEFT, padx=5)
        
        self.btn_clear_search = ttk.Button(search_frame, text="Hiển thị tất cả")
        self.btn_clear_search.pack(side=tk.LEFT, padx=5)

        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Segoe UI', 9, 'bold'))
        style.configure("Treeview", rowheight=25)

        y_scroll = ttk.Scrollbar(self.bottom_frame, orient=tk.VERTICAL)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        x_scroll = ttk.Scrollbar(self.bottom_frame, orient=tk.HORIZONTAL)
        x_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        columns = ("id", "ten_cong_viec", "noi_dung", "tai_lieu_dinh_kem", "nguoi_xu_ly", "nguoi_trinh_ky", "ngay_nhan", "ngay_bat_dau", "ngay_hoan_thanh", "trang_thai")
        self.tree = ttk.Treeview(self.bottom_frame, columns=columns, show="headings",
                                 yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        
        y_scroll.config(command=self.tree.yview)
        x_scroll.config(command=self.tree.xview)

        self.tree.heading("id", text="STT", anchor="center")
        self.tree.heading("ten_cong_viec", text="Tên Công Việc", anchor="w")
        self.tree.heading("noi_dung", text="Nội Dung Chi Tiết", anchor="w")
        self.tree.heading("tai_lieu_dinh_kem", text="Tài Liệu Đính Kèm", anchor="w")
        self.tree.heading("nguoi_xu_ly", text="Người Xử Lý", anchor="center")
        self.tree.heading("nguoi_trinh_ky", text="Người Trình Ký", anchor="center")
        self.tree.heading("ngay_nhan", text="Ngày Nhận", anchor="center")
        self.tree.heading("ngay_bat_dau", text="Ngày Bắt Đầu", anchor="center")
        self.tree.heading("ngay_hoan_thanh", text="Ngày Hoàn Thành", anchor="center")
        self.tree.heading("trang_thai", text="Trạng Thái", anchor="center")

        self.tree.column("id", width=40, anchor="center", stretch=False)
        self.tree.column("ten_cong_viec", width=220, anchor="w", stretch=False)
        self.tree.column("noi_dung", width=250, anchor="w", stretch=False)
        self.tree.column("tai_lieu_dinh_kem", width=180, anchor="w", stretch=False)
        self.tree.column("nguoi_xu_ly", width=120, anchor="center", stretch=False)
        self.tree.column("nguoi_trinh_ky", width=120, anchor="center", stretch=False)
        self.tree.column("ngay_nhan", width=90, anchor="center", stretch=False)
        self.tree.column("ngay_bat_dau", width=90, anchor="center", stretch=False)
        self.tree.column("ngay_hoan_thanh", width=100, anchor="center", stretch=False)
        self.tree.column("trang_thai", width=110, anchor="center", stretch=False)

        self.tree.tag_configure('hoan_thanh', background='#d4edda')       
        self.tree.tag_configure('dang_xu_ly', background='#fff3cd')       
        self.tree.tag_configure('cham_tien_do', background='#f8d7da')     
        self.tree.tag_configure('hoan_thanh_cham', background='#cce5ff')  

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)