"""
ỨNG DỤNG GIAO DIỆN DESKTOP (GUI APP)
Tự Động Hóa Xử Lý Chứng Từ Xuất Nhập Khẩu & Điền Tờ Khai ECUS5-VNACCS
- Chọn upload 1 hoặc nhiều file Excel cùng lúc.
- Bóc tách, tổng hợp dữ liệu vào Template.
- Đóng gói toàn bộ file kết quả vào 1 file ZIP và hỏi vị trí lưu trên máy người dùng.
- Tab 2: Xem trước dữ liệu và tự động điền vào ECUS5 (Windows) hoặc mô phỏng (macOS).
"""

import os
import sys
import threading
import time
import zipfile
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

# Import logic lõi từ 2 module có sẵn
import main as etl_module
import ecus_fill_demo as ecus_module

# Cấu hình giao diện CustomTkinter
ctk.set_appearance_mode("System")  # Tự động theo Dark/Light mode của máy
ctk.set_default_color_theme("blue")

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
IS_WINDOWS = sys.platform == "win32"


class EcusApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("ECUS5 Helper - Tự Động Hóa Chứng Từ & Tờ Khai Hải Quan")
        self.geometry("1060x780")
        self.minsize(920, 680)

        # Trạng thái dữ liệu
        self.selected_files = []  # Danh sách các file IV + PL được chọn
        self.template_file = BASE_DIR / "Template.xlsx"
        self.output_folder = BASE_DIR / "output"
        self.output_folder.mkdir(exist_ok=True)
        self.current_template_values = {}

        self._auto_detect_initial_files()
        self._build_ui()
        self._refresh_output_files()

    def _auto_detect_initial_files(self):
        """Tự động nạp sẵn các file trong thư mục input nếu có sẵn."""
        input_dir = BASE_DIR / "input"
        if input_dir.exists():
            files = [
                f for f in sorted(input_dir.glob("IV + PL G-*.xlsx"))
                if not f.name.startswith("~$")
            ]
            if files:
                self.selected_files = files

    def _build_ui(self):
        # 1. HEADER
        header_frame = ctk.CTkFrame(self, corner_radius=10)
        header_frame.pack(fill="x", padx=15, pady=(15, 10))

        title_label = ctk.CTkLabel(
            header_frame,
            text="HỆ THỐNG TỰ ĐỘNG HÓA TỜ KHAI HẢI QUAN ECUS5-VNACCS",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_label.pack(anchor="w", padx=15, pady=(10, 2))

        sub_text = "Công Ty TNHH ALDILA COMPOSITE PRODUCTS  |  Mã loại hình: E42 (DNCX)  |  Chi cục: 43NF"
        sub_label = ctk.CTkLabel(header_frame, text=sub_text, font=ctk.CTkFont(size=12), text_color="gray")
        sub_label.pack(anchor="w", padx=15, pady=(0, 8))

        # Badge hệ điều hành
        os_badge_text = "🖥️ HỆ ĐIỀU HÀNH: WINDOWS (SẴN SÀNG ĐIỀN THẬT VÀO ECUS5)" if IS_WINDOWS else "🍎 HỆ ĐIỀU HÀNH: MACOS (CHẾ ĐỘ XỬ LÝ EXCEL & MÔ PHỎNG DỮ LIỆU)"
        os_badge_color = "#2b7a4b" if IS_WINDOWS else "#8a5d14"
        badge = ctk.CTkLabel(
            header_frame,
            text=os_badge_text,
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=os_badge_color,
            corner_radius=6,
            text_color="white",
            padx=10, pady=3
        )
        badge.pack(anchor="w", padx=15, pady=(0, 10))

        # 2. TABS
        self.tabview = ctk.CTkTabview(self, corner_radius=10)
        self.tabview.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        self.tab_etl = self.tabview.add("1. Chọn File & Xuất ZIP")
        self.tab_ecus = self.tabview.add("2. Điền Tờ Khai ECUS5")

        self._build_tab_etl()
        self._build_tab_ecus()

    # =========================================================================
    # TAB 1: CHỌN FILE, XỬ LÝ & XUẤT ZIP
    # =========================================================================
    def _build_tab_etl(self):
        container = self.tab_etl

        # Box 1: Chọn file đầu vào
        upload_box = ctk.CTkFrame(container, corner_radius=8)
        upload_box.pack(fill="x", padx=10, pady=(10, 5))

        top_action_bar = ctk.CTkFrame(upload_box, fg_color="transparent")
        top_action_bar.pack(fill="x", padx=10, pady=(10, 5))

        btn_pick_files = ctk.CTkButton(
            top_action_bar,
            text="📂 Chọn 1 hoặc nhiều file hóa đơn (IV + PL)...",
            font=ctk.CTkFont(weight="bold"),
            height=34,
            command=self._on_pick_multiple_files
        )
        btn_pick_files.pack(side="left", padx=(0, 10))

        btn_pick_dir = ctk.CTkButton(
            top_action_bar,
            text="📁 Chọn cả thư mục...",
            width=150,
            height=34,
            fg_color="#3a3a3a",
            hover_color="#4a4a4a",
            command=self._on_pick_folder
        )
        btn_pick_dir.pack(side="left", padx=(0, 10))

        btn_clear_files = ctk.CTkButton(
            top_action_bar,
            text="✕ Xóa chọn lại",
            width=110,
            height=34,
            fg_color="#8c2a2a",
            hover_color="#a83232",
            command=self._on_clear_selected_files
        )
        btn_clear_files.pack(side="right")

        # Label trạng thái file đã chọn
        self.lbl_file_count = ctk.CTkLabel(
            upload_box,
            text="Chưa có file nào được chọn",
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w"
        )
        self.lbl_file_count.pack(fill="x", padx=12, pady=(2, 4))

        # Khung danh sách file đã chọn (Textbox hiển thị ngắn gọn)
        self.txt_selected_files = ctk.CTkTextbox(upload_box, height=85, font=ctk.CTkFont(family="Courier", size=11))
        self.txt_selected_files.pack(fill="x", padx=10, pady=(0, 10))
        self._update_selected_files_display()

        # Box 2: Chọn Template mẫu chuẩn
        tmpl_box = ctk.CTkFrame(container, corner_radius=8)
        tmpl_box.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(tmpl_box, text="File mẫu chuẩn (Template.xlsx):", width=220, anchor="w").grid(row=0, column=0, padx=10, pady=8, sticky="w")
        self.entry_template = ctk.CTkEntry(tmpl_box)
        self.entry_template.insert(0, str(self.template_file))
        self.entry_template.grid(row=0, column=1, padx=5, pady=8, sticky="ew")
        btn_browse_tmpl = ctk.CTkButton(tmpl_box, text="Chọn file...", width=110, command=self._browse_template)
        btn_browse_tmpl.grid(row=0, column=2, padx=10, pady=8)
        tmpl_box.grid_columnconfigure(1, weight=1)

        # Nút to: Bắt đầu xử lý và xuất file ZIP
        btn_bar = ctk.CTkFrame(container, fg_color="transparent")
        btn_bar.pack(fill="x", padx=10, pady=6)

        self.btn_run_etl = ctk.CTkButton(
            btn_bar,
            text="⚡ 1. BẮT ĐẦU BÓC TÁCH & CHỌN NƠI LƯU FILE ZIP",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=44,
            fg_color="#1f6aa5",
            hover_color="#144d77",
            command=self._start_etl_and_export_zip
        )
        self.btn_run_etl.pack(fill="x", expand=True)

        # Log Console
        ctk.CTkLabel(container, text="Nhật ký xử lý chi tiết (Logs):", anchor="w", font=ctk.CTkFont(weight="bold")).pack(fill="x", padx=10, pady=(6, 2))
        self.log_text = ctk.CTkTextbox(container, font=ctk.CTkFont(family="Courier", size=11))
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    # =========================================================================
    # TAB 2: ĐIỀN TỜ KHAI ECUS5
    # =========================================================================
    def _build_tab_ecus(self):
        container = self.tab_ecus

        top_bar = ctk.CTkFrame(container, corner_radius=8)
        top_bar.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(top_bar, text="Chọn file Template trong output:", width=220, anchor="w").pack(side="left", padx=10, pady=8)
        self.combo_files = ctk.CTkComboBox(top_bar, values=["Chưa có file nào"], width=360, command=self._on_select_template_file)
        self.combo_files.pack(side="left", padx=5, pady=8, fill="x", expand=True)
        btn_refresh = ctk.CTkButton(top_bar, text="🔄 Làm mới", width=90, command=self._refresh_output_files)
        btn_refresh.pack(side="right", padx=10, pady=8)

        # Bảng hiển thị 12 trường dữ liệu
        table_label = ctk.CTkLabel(container, text="Bảng xem trước dữ liệu tờ khai (12 trường sẽ điền vào tab Thông tin chung):", anchor="w", font=ctk.CTkFont(weight="bold"))
        table_label.pack(fill="x", padx=10, pady=(5, 2))

        # Sử dụng Treeview ttk để hiển thị bảng chuyên nghiệp
        tree_frame = ctk.CTkFrame(container, corner_radius=6)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("field", "value", "source", "status")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=8)
        self.tree.heading("field", text="Trường thông tin")
        self.tree.heading("value", text="Giá trị sẽ điền")
        self.tree.heading("source", text="Ô nguồn")
        self.tree.heading("status", text="Trạng thái")

        self.tree.column("field", width=220, anchor="w")
        self.tree.column("value", width=380, anchor="w")
        self.tree.column("source", width=120, anchor="center")
        self.tree.column("status", width=130, anchor="center")

        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")

        # Khung Ký hiệu & số hiệu chi tiết
        sign_box = ctk.CTkFrame(container, corner_radius=8)
        sign_box.pack(fill="x", padx=10, pady=5)

        sign_lbl = ctk.CTkLabel(sign_box, text="Chuỗi 'Ký hiệu và số hiệu' đầy đủ:", anchor="w", font=ctk.CTkFont(weight="bold"))
        sign_lbl.pack(anchor="w", padx=10, pady=(6, 2))

        self.entry_sign = ctk.CTkEntry(sign_box, font=ctk.CTkFont(size=12))
        self.entry_sign.pack(side="left", fill="x", expand=True, padx=(10, 5), pady=(0, 8))

        btn_copy_sign = ctk.CTkButton(sign_box, text="📋 Copy", width=80, command=self._copy_sign_to_clipboard)
        btn_copy_sign.pack(side="right", padx=(5, 10), pady=(0, 8))

        # Nút Hành động điền
        action_btn_text = "🚀 2. TỰ ĐỘNG ĐIỀN VÀO PHẦN MỀM ECUS5 (TAB THÔNG TIN CHUNG)" if IS_WINDOWS else "▶️ 2. CHẠY MÔ PHỎNG TIẾN TRÌNH ĐIỀN (SIMULATION TRÊN MAC)"
        action_btn_color = "#1f6aa5" if IS_WINDOWS else "#935c18"

        self.btn_run_ecus = ctk.CTkButton(
            container,
            text=action_btn_text,
            font=ctk.CTkFont(size=14, weight="bold"),
            height=44,
            fg_color=action_btn_color,
            command=self._start_ecus_process
        )
        self.btn_run_ecus.pack(fill="x", padx=10, pady=(8, 10))

    # =========================================================================
    # SỰ KIỆN CHỌN FILE VÀ UPLOAD
    # =========================================================================
    def _on_pick_multiple_files(self):
        """Cho phép người dùng chọn 1 hoặc nhiều file cùng lúc."""
        initial_dir = self.selected_files[0].parent if self.selected_files else (BASE_DIR / "input" if (BASE_DIR / "input").exists() else BASE_DIR)
        chosen = filedialog.askopenfilenames(
            title="Chọn một hoặc nhiều file hóa đơn (IV + PL)",
            initialdir=str(initial_dir),
            filetypes=[("Excel Files (*.xlsx)", "*.xlsx"), ("All Files", "*.*")]
        )
        if chosen:
            new_paths = [Path(p) for p in chosen if not Path(p).name.startswith("~$")]
            self.selected_files = new_paths
            self._update_selected_files_display()

    def _on_pick_folder(self):
        """Cho phép người dùng chọn cả thư mục nếu có nhiều file."""
        initial_dir = BASE_DIR / "input" if (BASE_DIR / "input").exists() else BASE_DIR
        chosen_dir = filedialog.askdirectory(title="Chọn thư mục chứa các file hóa đơn", initialdir=str(initial_dir))
        if chosen_dir:
            d = Path(chosen_dir)
            files = [
                f for f in sorted(d.glob("IV + PL G-*.xlsx"))
                if not f.name.startswith("~$")
            ]
            if not files:
                # Quét tất cả file xlsx nếu không khớp tiền tố
                files = [f for f in sorted(d.glob("*.xlsx")) if not f.name.startswith("~$") and f.name != "Template.xlsx"]
            self.selected_files = files
            self._update_selected_files_display()

    def _on_clear_selected_files(self):
        self.selected_files = []
        self._update_selected_files_display()

    def _update_selected_files_display(self):
        self.txt_selected_files.delete("1.0", "end")
        n = len(self.selected_files)
        if n == 0:
            self.lbl_file_count.configure(text="Chưa có file nào được chọn. Hãy bấm nút 'Chọn 1 hoặc nhiều file...' ở trên.", text_color="gray")
        else:
            self.lbl_file_count.configure(text=f"🟢 Đã chọn {n} file hóa đơn để xử lý:", text_color="#2b7a4b" if not IS_WINDOWS else "#2ecc71")
            lines = [f"{i+1:2d}. {p.name}  ({p.parent})" for i, p in enumerate(self.selected_files)]
            self.txt_selected_files.insert("1.0", "\n".join(lines))

    def _browse_template(self):
        f = filedialog.askopenfilename(initialdir=BASE_DIR, filetypes=[("Excel Files", "*.xlsx")])
        if f:
            self.template_file = Path(f)
            self.entry_template.delete(0, "end")
            self.entry_template.insert(0, str(self.template_file))

    def _append_log(self, text):
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")

    # =========================================================================
    # TIẾN TRÌNH XỬ LÝ & NÉN ZIP
    # =========================================================================
    def _start_etl_and_export_zip(self):
        if not self.selected_files:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất 1 file hóa đơn (IV + PL) trước khi bắt đầu!")
            return

        tmpl_path = Path(self.entry_template.get().strip())
        if not tmpl_path.exists():
            messagebox.showerror("Lỗi", f"Không tìm thấy file template mẫu:\n{tmpl_path}")
            return

        # 1. HỎI NGƯỜI DÙNG NƠI LƯU FILE ZIP TRÊN MÁY
        now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_zip_name = f"KetQua_ToKhai_{now_str}.zip"

        save_zip_path = filedialog.asksaveasfilename(
            title="Chọn nơi lưu file ZIP kết quả trên máy của bạn",
            defaultextension=".zip",
            filetypes=[("Zip Archive (*.zip)", "*.zip")],
            initialfile=default_zip_name
        )

        if not save_zip_path:
            # Người dùng bấm Hủy (Cancel)
            return

        self.btn_run_etl.configure(state="disabled", text="⏳ Đang bóc tách và nén file ZIP...")
        self.log_text.delete("1.0", "end")
        self._append_log("=== BẮT ĐẦU TIẾN TRÌNH XỬ LÝ CHỨNG TỪ ===")
        self._append_log(f"Tổng số file cần xử lý: {len(self.selected_files)}")
        self._append_log(f"Vị trí sẽ xuất file ZIP: {save_zip_path}")

        # Chạy trong Thread để giao diện không bị treo
        threading.Thread(target=self._etl_worker, args=(Path(save_zip_path), tmpl_path), daemon=True).start()

    def _etl_worker(self, zip_output_path, tmpl_path):
        try:
            created_files = []
            folder_inside_zip = zip_output_path.stem  # Tên folder con bên trong file zip

            for i, source_path in enumerate(self.selected_files):
                inv = etl_module.get_inv(source_path.name)
                if not inv:
                    inv = source_path.stem

                self._append_log(f"\n[{i+1}/{len(self.selected_files)}] Đang xử lý: {source_path.name}...")

                source_vals = etl_module.load_workbook(source_path, data_only=True)
                source_forms = etl_module.load_workbook(source_path, data_only=False)
                target_wb = etl_module.load_workbook(tmpl_path)

                # Điền labels
                for label in etl_module.LABELS:
                    val, found = etl_module.get_value(source_vals, source_forms, label)
                    if found and val is not None:
                        etl_module.put_value(target_wb, label, val)
                        self._append_log(f"  + {label} {val}")

                # Bóc tách
                pack_info = etl_module.report_packing(source_vals)
                head_info = etl_module.report_header(source_vals, inv)
                totals_info = etl_module.report_totals(source_vals)
                misc_info = etl_module.report_misc(source_vals)

                # Copy bảng hàng
                n_rows = etl_module.copy_table(source_vals, target_wb)
                if n_rows:
                    etl_module.trim_extra_rows(target_wb, n_rows)

                # Ghi vào template
                etl_module.write_packing(target_wb, pack_info)
                etl_module.write_header(target_wb, head_info)
                etl_module.write_totals(target_wb, totals_info)
                etl_module.write_misc(target_wb, misc_info)

                target_wb.calculation.fullCalcOnLoad = True
                out_path = self.output_folder / f"Template {inv}.xlsx"
                target_wb.save(out_path)
                created_files.append(out_path)

                source_vals.close()
                source_forms.close()
                target_wb.close()
                self._append_log(f"  -> Đã tạo kết quả: {out_path.name}")

            # Tính toán ngầm Excel nếu trên Windows
            if IS_WINDOWS and etl_module.RECALC_WITH_EXCEL:
                self._append_log("\nĐang gọi Excel chạy ngầm tính toán lại công thức...")
                etl_module.recalc_with_excel(created_files)

            # 2. ĐÓNG GÓI TẤT CẢ FILE KẾT QUẢ VÀO FILE ZIP THEO NƠI USER CHỌN
            self._append_log(f"\nĐang nén {len(created_files)} file vào file ZIP: {zip_output_path.name}...")
            with zipfile.ZipFile(zip_output_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for fpath in created_files:
                    # Cho vào thư mục con trong zip để giải nén ra rất gọn gàng
                    arcname = f"{folder_inside_zip}/{fpath.name}"
                    zf.write(fpath, arcname=arcname)

            file_size_kb = zip_output_path.stat().st_size / 1024
            self._append_log(f"-> Nén thành công! Dung lượng ZIP: {file_size_kb:.1f} KB")
            self._append_log("\n=== HOÀN TẤT TOÀN BỘ TIẾN TRÌNH! ===")

            # Hộp thoại thông báo thành công có nút mở folder
            msg = (
                f"Đã bóc tách thành công {len(created_files)} file!\n\n"
                f"File ZIP kết quả đã được lưu tại:\n{zip_output_path}\n"
                f"(Dung lượng nén: {file_size_kb:.1f} KB)\n\n"
                f"Dữ liệu cũng đã được nạp sẵn vào Tab 2 để bạn kiểm tra và điền vào ECUS."
            )
            messagebox.showinfo("Thành công", msg)

        except Exception as e:
            self._append_log(f"\n[LỖI]: {e}")
            messagebox.showerror("Lỗi xử lý", str(e))
        finally:
            self.btn_run_etl.configure(state="normal", text="⚡ 1. BẮT ĐẦU BÓC TÁCH & CHỌN NƠI LƯU FILE ZIP")
            self._refresh_output_files()

    # =========================================================================
    # TAB 2: QUẢN LÝ DỮ LIỆU ĐIỀN ECUS
    # =========================================================================
    def _refresh_output_files(self):
        files = sorted(self.output_folder.glob("Template G-*.xlsx"))
        files = [f for f in files if not f.name.startswith("~$")]
        if files:
            names = [f.name for f in files]
            self.combo_files.configure(values=names)
            self.combo_files.set(names[0])
            self._load_template_data(files[0])
        else:
            self.combo_files.configure(values=["(Chưa có file nào trong output)"])
            self.combo_files.set("(Chưa có file nào trong output)")
            self._clear_tree()

    def _on_select_template_file(self, filename):
        path = self.output_folder / filename
        if path.exists():
            self._load_template_data(path)

    def _clear_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.entry_sign.delete(0, "end")

    def _load_template_data(self, path):
        self._clear_tree()
        try:
            values = ecus_module.read_template(path)
            self.current_template_values = values

            for name, key, label, idx, fmt in ecus_module.FIELDS:
                raw = values.get(key)
                if raw is None or str(raw).strip() == "":
                    src = ecus_module.CELLS.get(key, "Công thức")
                    self.tree.insert("", "end", values=(name, "(TRỐNG)", src, "Thiếu dữ liệu"))
                    continue
                try:
                    text = fmt(raw)
                    src = ecus_module.CELLS.get(key, "Công thức")
                    self.tree.insert("", "end", values=(name, text, src, "Hợp lệ"))
                except Exception:
                    self.tree.insert("", "end", values=(name, str(raw), "", "Lỗi định dạng"))

            # Điền chuỗi ký hiệu & số hiệu
            sign = values.get("sign_text", "")
            if sign:
                self.entry_sign.insert(0, sign)

        except Exception as e:
            messagebox.showerror("Lỗi đọc file", f"Không đọc được file {path.name}: {e}")

    def _copy_sign_to_clipboard(self):
        text = self.entry_sign.get().strip()
        if text:
            self.clipboard_clear()
            self.clipboard_append(text)
            messagebox.showinfo("Clipboard", "Đã sao chép chuỗi 'Ký hiệu và số hiệu' vào Clipboard!")

    def _start_ecus_process(self):
        selected_file = self.combo_files.get()
        target_path = self.output_folder / selected_file
        if not target_path.exists():
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn 1 file Template hợp lệ trong danh sách!")
            return

        if not IS_WINDOWS:
            # Chạy mô phỏng trên macOS
            confirm = messagebox.askyesno(
                "Chế độ mô phỏng (macOS)",
                f"Bạn đang chạy trên macOS (không có phần mềm ECUS5 trực tiếp).\n"
                f"Bạn có muốn chạy mô phỏng tiến trình gõ 12 trường của file:\n{selected_file} không?"
            )
            if confirm:
                threading.Thread(target=self._mac_simulation_worker, daemon=True).start()
        else:
            # Chạy thật trên Windows
            confirm = messagebox.askyesno(
                "Xác nhận điền vào ECUS5",
                f"Bạn có chắc muốn tự động điền dữ liệu của file:\n{selected_file}\nvào cửa sổ ECUS5 đang mở không?\n\n"
                f"(Hãy chắc chắn đã mở tờ khai xuất khẩu ở tab Thông tin chung!)"
            )
            if confirm:
                self.btn_run_ecus.configure(state="disabled", text="⏳ Đang điền vào ECUS5...")
                threading.Thread(target=self._windows_ecus_worker, args=(target_path,), daemon=True).start()

    def _mac_simulation_worker(self):
        self.btn_run_ecus.configure(state="disabled", text="⏳ Đang mô phỏng điền...")
        try:
            for item in self.tree.get_children():
                vals = self.tree.item(item)["values"]
                self.tree.item(item, values=(vals[0], vals[1], vals[2], "Đang gõ..."))
                time.sleep(0.15)
                self.tree.item(item, values=(vals[0], vals[1], vals[2], "Đã điền [OK]"))
            messagebox.showinfo("Hoàn tất mô phỏng", "Đã mô phỏng điền thành công toàn bộ 12 trường dữ liệu trên Mac!")
        finally:
            self.btn_run_ecus.configure(state="normal", text="▶️ 2. CHẠY MÔ PHỎNG TIẾN TRÌNH ĐIỀN (SIMULATION TRÊN MAC)")

    def _windows_ecus_worker(self, path):
        try:
            ecus_module.setup_windows()
            main_hwnd = ecus_module.find_ecus_window()
            if main_hwnd is None:
                messagebox.showerror("Lỗi", "Không tìm thấy cửa sổ ECUS5 đang mở! Vui lòng mở ECUS5 trước.")
                return

            ok, why = ecus_module.go_to_tab(main_hwnd, ecus_module.TAB_THONG_TIN_CHUNG)
            if not ok:
                messagebox.showwarning("Cảnh báo", f"Không tự chuyển được sang tab Thông tin chung: {why}")

            values = ecus_module.read_template(path)
            items = ecus_module.collect_items(main_hwnd)

            plan = []
            for name, key, label, idx, fmt in ecus_module.FIELDS:
                raw = values.get(key)
                if raw is None or str(raw).strip() == "":
                    continue
                try:
                    text = fmt(raw)
                except Exception:
                    continue
                ctl, note = ecus_module.locate(items, label, idx)
                if ctl:
                    plan.append((name, ctl, text))

            if not plan:
                messagebox.showwarning("Cảnh báo", "Không tìm thấy ô nhập liệu nào trên ECUS để điền!")
                return

            # Đếm ngược 3s
            time.sleep(3)
            for name, ctl, text in plan:
                ok = ecus_module.write_field(main_hwnd, ctl.hwnd, text)
                time.sleep(0.05)

            ecus_module.go_to_tab(main_hwnd, ecus_module.DEFAULT_TAB)
            messagebox.showinfo("Thành công", "Đã tự động điền xong 12 trường vào ECUS5 và chuyển sang tab Danh sách hàng!")

        except Exception as e:
            messagebox.showerror("Lỗi khi điền ECUS", str(e))
        finally:
            self.btn_run_ecus.configure(state="normal", text="🚀 2. TỰ ĐỘNG ĐIỀN VÀO PHẦN MỀM ECUS5 (TAB THÔNG TIN CHUNG)")


if __name__ == "__main__":
    app = EcusApp()
    app.mainloop()
