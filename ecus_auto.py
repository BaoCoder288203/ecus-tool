"""
ECUS Auto - Bee Logistics
Giao dien chung cho moi khach hang: dua hoa don vao -> tao Template tung lo -> bo sung o thieu -> chay ECUS.

Chay: nhap dup "ECUS Auto.bat" (hoac: python ecus_auto.py) trong thu muc D:\\test.

Cau truc thu muc cua moi khach (tu tao khi chon khach):
    D:\\test\\Don hang\\<Khach> xuat\\
        Template_<khach>.xlsx   Template goc cua khach (vd Template_aldila.xlsx), chep tu khach_hang\\Template_chuan.xlsx
        dau_vao\\          tep vua dua vao, cho tao Template
        da_xu_ly\\         tep da tao Template xong (chia theo lan chay)
        output\\           Template tung lo -> bo sung o thieu, Ctrl+S -> Chay ECUS
        nho_ma.json       ma dia diem da khai (de tu dien lan sau)
        da_khai.json      lo nao da dien vao ECUS, luc nao

Khach hang moi: them 1 file ho so trong khach_hang\\ho_so (xem tai lieu huong dan). Giao dien tu hien.
Keo tha tep: cai them "pip install tkinterdnd2" (khong bat buoc; khong co thi dung nut "Them tep").
"""
import builtins
import importlib
import unicodedata
import json
import os
import queue
import shutil
import subprocess
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk
except ImportError:
    print("Python tren may chua co tkinter (giao dien). Cai lai Python tu python.org va chon 'tcl/tk and IDLE'.")
    raise SystemExit(1)

try:                                    # keo tha tep (khong bat buoc)
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:
    TkinterDnD = None

GOC = Path(__file__).resolve().parent
sys.path.insert(0, str(GOC))
sys.path.insert(0, str(GOC / "khach_hang"))

import chay                     # noqa: E402  (khach_hang\chay.py)
import ecus_fill_demo as ecus   # noqa: E402
import chung as chay_chung      # noqa: E402  (khach_hang\chung.py: THU_MUC_DON)

THU_MUC_HO_SO = GOC / "khach_hang" / "ho_so"
TEMPLATE_CHUAN = GOC / "khach_hang" / "Template_chuan.xlsx"
LOAI_HOA_DON = {                       # lua chon tren hop "Them khach" -> (MAU_FILE, LOAI_TEP)
    "Word (.docx)": ("**/*.doc*", "*.docx *.doc"),
    "Excel (.xlsx)": ("**/*.xlsx", "*.xlsx *.xls"),
    "PDF": ("**/*.pdf", "*.pdf"),
    "Nhiều loại": ("**/*.*", "*.docx *.doc *.xlsx *.xls *.pdf"),
}

TEN_APP = "ECUS Auto - Bee Logistics"
PHIEN_BAN = "0.1"
DAU_VAO, DA_XU_LY, DAU_RA = "dau_vao", "da_xu_ly", "output"
LA_WINDOWS = sys.platform == "win32"

PT_VC = [("1", "Đường không"), ("2", "Đường biển (container)"), ("3", "Đường biển (không cont)"),
         ("4", "Đường bộ (xe tải)"), ("5", "Đường sắt"), ("6", "Đường sông"), ("9", "Loại khác")]

MAU = {
    "nen": "#f4f6f9", "dau": "#1f3a5f", "chu_dau": "#ffffff", "phu_dau": "#c9d6e8",
    "du": "#e3f1df", "thieu": "#fde4e2", "kiem": "#fff6d6", "da_khai": "#eef1f5", "chu_mo": "#6b7785",
}


# ===================== tien ich =====================
def mo_tep(path):
    """Mo tep / thu muc bang chuong trinh mac dinh cua may."""
    path = str(path)
    if LA_WINDOWS:
        os.startfile(path)
    else:
        subprocess.Popen(["xdg-open", path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def dang_mo_trong_excel(path):
    """Excel tao tep khoa '~$ten' khi dang mo tep."""
    p = Path(path)
    return (p.parent / ("~$" + p.name)).exists() or (p.parent / ("~$" + p.name[2:])).exists()


def vao_thung_rac(path):
    """Windows: dua file vao Thung rac (lay lai duoc). May khac: xoa han."""
    path = Path(path)
    if not LA_WINDOWS:
        path.unlink()
        return
    import ctypes
    from ctypes import wintypes

    class SHFILEOPSTRUCTW(ctypes.Structure):
        _fields_ = [("hwnd", wintypes.HWND), ("wFunc", ctypes.c_uint), ("pFrom", wintypes.LPCWSTR),
                    ("pTo", wintypes.LPCWSTR), ("fFlags", ctypes.c_ushort), ("fAnyOperationsAborted", wintypes.BOOL),
                    ("hNameMappings", ctypes.c_void_p), ("lpszProgressTitle", wintypes.LPCWSTR)]
    FO_DELETE, FOF_ALLOWUNDO, FOF_NOCONFIRMATION, FOF_SILENT, FOF_NOERRORUI = 3, 0x40, 0x10, 0x4, 0x400
    op = SHFILEOPSTRUCTW(None, FO_DELETE, str(path.resolve()) + "\0", None,
                         FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI, False, None, None)
    kq = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    if kq != 0 or path.exists():
        raise OSError(f"SHFileOperation loi {kq}")


def khong_dau(chu):
    chu = unicodedata.normalize("NFD", chu).replace("đ", "d").replace("Đ", "D")
    return "".join(c for c in chu if unicodedata.category(c) != "Mn")


class GhiLog:
    """Thay sys.stdout: moi lenh print cua cac script hien vao o Nhat ky."""
    def __init__(self, q):
        self.q = q

    def write(self, s):
        if s:
            self.q.put(("log", s))

    def flush(self):
        pass


# ===================== hop thoai them khach =====================
class HopThemKhach:
    def __init__(self, app):
        self.app = app
        w = self.w = tk.Toplevel(app.root)
        w.title("Thêm khách hàng")
        w.transient(app.root)
        w.resizable(False, False)
        w.configure(bg=MAU["nen"])
        f = ttk.Frame(w, padding=16)
        f.pack(fill="both", expand=True)
        f.columnconfigure(1, weight=1)

        def dong(r, nhan, o, goi_y=""):
            ttk.Label(f, text=nhan).grid(row=r, column=0, sticky="nw", pady=5, padx=(0, 10))
            o.grid(row=r, column=1, sticky="ew", pady=5)
            if goi_y:
                ttk.Label(f, text=goi_y, style="Phu.TLabel").grid(row=r + 1, column=1, sticky="w")

        self.ten = ttk.Entry(f, width=40)
        dong(0, "Tên khách", self.ten, "Ví dụ: Nike, Top Glove. Thư mục sẽ là Don hang\\<Tên> xuat")
        ttk.Label(f, text="Tờ khai").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Label(f, text="Xuất khẩu (tờ khai nhập chưa hỗ trợ)").grid(row=2, column=1, sticky="w")
        self.loai = ttk.Combobox(f, state="readonly", values=list(LOAI_HOA_DON))
        self.loai.current(0)
        dong(3, "Hóa đơn dạng", self.loai)
        self.cach = ttk.Combobox(f, state="readonly", values=["AI đọc (dùng ngay, cần API key)",
                                                                "Quy tắc (phải viết hàm doc())"])
        self.cach.current(0)
        dong(4, "Cách đọc", self.cach, "Bắt đầu bằng AI; khách ổn định nhiều lô thì chuyển sang quy tắc.")
        nguon = ["Template chuẩn"] + [f"Copy từ: {self.app.ds_khach.get(i).strip()}"
                                     for i in range(self.app.ds_khach.size())]
        self.nguon = ttk.Combobox(f, state="readonly", values=nguon)
        self.nguon.current(0)
        dong(6, "Template", self.nguon, "Template chuẩn: Sheet1 + danh mục mã, Sheet3 trống để cài barem.")
        ttk.Label(f, text="Đặc điểm hóa đơn\n(cho AI, không bắt buộc)").grid(row=8, column=0, sticky="nw", pady=5)
        self.goi_y = tk.Text(f, width=46, height=4, font=("Segoe UI", 10), relief="flat", highlightthickness=1)
        self.goi_y.grid(row=8, column=1, sticky="ew", pady=5)
        ttk.Label(f, style="Phu.TLabel", text="Ví dụ: giá ghi theo 1000 cái; người nhận ở ô Consignee; "
                                              "dòng FOC là hàng không tính tiền.").grid(row=9, column=1, sticky="w")
        nut = ttk.Frame(f)
        nut.grid(row=10, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(nut, text="Hủy", command=w.destroy).pack(side="right")
        ttk.Button(nut, text="Tạo khách", style="Chinh.TButton", command=self.tao).pack(side="right", padx=8)
        self.ten.focus_set()
        w.bind("<Return>", lambda e: self.tao() if e.widget is not self.goi_y else None)
        w.bind("<Escape>", lambda e: w.destroy())
        w.update_idletasks()                   # dat hop giua cua so chinh
        x = app.root.winfo_rootx() + (app.root.winfo_width() - w.winfo_reqwidth()) // 2
        y = app.root.winfo_rooty() + (app.root.winfo_height() - w.winfo_reqheight()) // 3
        w.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        w.grab_set()

    def tao(self):
        chon = self.nguon.current()
        if chon == 0:
            tpl = TEMPLATE_CHUAN
        else:
            hs = chay.nap_ho_so(self.app.ds_ten[chon - 1])
            tpl = chay_chung.template_goc(hs)
        if not Path(tpl).exists():
            messagebox.showwarning(TEN_APP, f"Không thấy Template nguồn:\n{tpl}", parent=self.w)
            return
        try:
            ma = self.app.tao_khach(self.ten.get(), self.loai.get(), "ai" if self.cach.current() == 0 else "quy_tac",
                                    self.goi_y.get("1.0", "end"), tpl)
        except Exception as e:
            messagebox.showwarning(TEN_APP, str(e), parent=self.w)
            return
        self.w.destroy()
        self.app._nap_khach(chon_ten=ma)
        if messagebox.askyesno(TEN_APP, "Đã tạo khách mới.\n\nBước tiếp theo:\n"
                                        "1. Mở Template gốc, chỉnh Sheet3 theo barem ECUS của khách.\n"
                                        "2. (Nếu cần) Cài đặt khách: sửa mẫu Ký hiệu và số hiệu, cột hàng.\n"
                                        "3. Thêm hóa đơn mẫu, Tạo Template, so với tờ khai làm tay.\n\n"
                                        "Mở Template gốc ngay bây giờ?"):
            self.app._mo_template_goc()


# ===================== ung dung =====================
class UngDung:
    def __init__(self):
        if LA_WINDOWS:                         # bat DPI truoc khi tao cua so (chu net, toa do click dung)
            try:
                ecus.setup_windows()
            except Exception:
                pass
        self.root =TkinterDnD.Tk() if TkinterDnD else tk.Tk()
        self.root.title(TEN_APP)
        self.root.geometry("1240x800")
        self.root.minsize(1100, 680)
        self.root.configure(bg=MAU["nen"])

        self.q = queue.Queue()                 # luong nen -> giao dien
        self.q_kiem = queue.Queue()            # danh sach file can kiem tra trang thai
        self.ban = False                       # dang chay viec nen
        self.khach = None                      # module ho so dang chon
        self.ten_khach = None
        self.trang_thai = {}                   # path -> (mtime, thieu)
        self.danh_muc = {}                     # ten_khach -> danh muc ma (Sheet4)
        self.ecus_san_sang = False
        self._luot = 0                         # so luot kiem tra (bo ket qua cu khi vua sua)
        self.cho_bao = set()                   # lo vua ap booking / dia diem -> bao ket qua kiem tra vao nhat ky

        sys.stdout = GhiLog(self.q)
        sys.stderr = GhiLog(self.q)
        builtins.input = self._input_qua_hop_thoai

        self._kieu()
        self._dung_giao_dien()
        self._nap_khach()

        threading.Thread(target=self._luong_kiem_tra, daemon=True).start()
        self.root.after(100, self._xu_ly_hang_doi)
        self.root.after(500, self._theo_doi_thu_muc)
        self.root.after(800, self._theo_doi_ecus)

    # ---------- giao dien ----------
    def _kieu(self):
        s = ttk.Style(self.root)
        if "vista" not in s.theme_names():
            s.theme_use("clam")
        s.configure(".", font=("Segoe UI", 10))
        s.configure("TFrame", background=MAU["nen"])
        s.configure("TLabelframe", background=MAU["nen"])
        s.configure("TLabelframe.Label", background=MAU["nen"], font=("Segoe UI", 11, "bold"),
                    foreground=MAU["dau"])
        s.configure("TLabel", background=MAU["nen"])
        s.configure("Phu.TLabel", background=MAU["nen"], foreground=MAU["chu_mo"], font=("Segoe UI", 9))
        s.configure("Chinh.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 7))
        s.configure("TButton", padding=(10, 5))
        s.configure("Treeview", rowheight=26, font=("Segoe UI", 10))
        s.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

    def _dung_giao_dien(self):
        r = self.root
        r.columnconfigure(0, weight=1)
        r.rowconfigure(1, weight=1)

        # --- dau trang ---
        dau = tk.Frame(r, bg=MAU["dau"], height=64)
        dau.grid(row=0, column=0, sticky="ew")
        dau.columnconfigure(1, weight=1)
        tk.Label(dau, text="ECUS Auto", bg=MAU["dau"], fg=MAU["chu_dau"],
                 font=("Segoe UI", 18, "bold")).grid(row=0, column=0, padx=(18, 8), pady=(10, 0), sticky="w")
        tk.Label(dau, text="Bee Logistics · Tự động điền tờ khai xuất khẩu ECUS", bg=MAU["dau"],
                 fg=MAU["phu_dau"], font=("Segoe UI", 10)).grid(row=1, column=0, columnspan=2, padx=18,
                                                                 pady=(0, 10), sticky="w")
        self.nhan_ecus = tk.Label(dau, text="ECUS: đang kiểm tra...", bg=MAU["dau"], fg=MAU["phu_dau"],
                                  font=("Segoe UI", 10, "bold"))
        self.nhan_ecus.grid(row=0, column=2, rowspan=2, padx=18)

        than = ttk.Frame(r, padding=(12, 10, 12, 4))
        than.grid(row=1, column=0, sticky="nsew")
        than.columnconfigure(1, weight=1)
        than.rowconfigure(0, weight=1)

        # --- cot trai: khach hang ---
        trai = ttk.LabelFrame(than, text="Khách hàng", padding=8)
        trai.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        self.ds_khach = tk.Listbox(trai, width=24, height=14, activestyle="none", exportselection=False,
                                   font=("Segoe UI", 11), relief="flat", highlightthickness=1,
                                   selectbackground=MAU["dau"], selectforeground="white")
        self.ds_khach.pack(fill="both", expand=True)
        self.ds_khach.bind("<<ListboxSelect>>", lambda e: self._chon_khach())
        ttk.Button(trai, text="+  Thêm khách hàng", style="Chinh.TButton",
                   command=self._them_khach).pack(fill="x", pady=(8, 0))
        ttk.Button(trai, text="Mở thư mục khách", command=self._mo_thu_muc_khach).pack(fill="x", pady=(6, 0))
        ttk.Button(trai, text="Mở Template gốc", command=self._mo_template_goc).pack(fill="x", pady=(6, 0))
        ttk.Button(trai, text="Cài đặt khách (hồ sơ)", command=self._mo_ho_so).pack(fill="x", pady=(6, 0))

        # --- cot phai ---
        phai = ttk.Frame(than)
        phai.grid(row=0, column=1, sticky="nsew")
        phai.columnconfigure(0, weight=1)
        phai.rowconfigure(1, weight=1)

        # Buoc 1: dau vao
        b1 = ttk.LabelFrame(phai, text="Bước 1 · Đầu vào: hóa đơn cần xử lý", padding=8)
        b1.grid(row=0, column=0, sticky="ew")
        b1.columnconfigure(0, weight=1)
        self.ds_vao = tk.Listbox(b1, height=5, selectmode="extended", activestyle="none",
                                 font=("Segoe UI", 10), relief="flat", highlightthickness=1)
        self.ds_vao.grid(row=0, column=0, rowspan=3, sticky="nsew")
        cot = ttk.Frame(b1)
        cot.grid(row=0, column=1, rowspan=3, sticky="n", padx=(8, 0))
        ttk.Button(cot, text="Thêm tệp…", command=self._them_tep).pack(fill="x")
        ttk.Button(cot, text="Bỏ tệp đã chọn", command=self._bo_tep).pack(fill="x", pady=4)
        ttk.Button(cot, text="Mở thư mục", command=lambda: self._mo_con(DAU_VAO)).pack(fill="x")
        duoi1 = ttk.Frame(b1)
        duoi1.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        duoi1.columnconfigure(0, weight=1)
        self.goi_y_vao = ttk.Label(duoi1, style="Phu.TLabel",
                                   text=("Kéo thả tệp vào ô trên, hoặc bấm Thêm tệp." if TkinterDnD
                                         else "Bấm Thêm tệp (cài tkinterdnd2 để kéo thả)."))
        self.goi_y_vao.grid(row=0, column=0, sticky="w")
        self.nut_tao = ttk.Button(duoi1, text="▶  Tạo Template", style="Chinh.TButton", command=self._tao)
        self.nut_tao.grid(row=0, column=1, sticky="e")
        if TkinterDnD:
            self.ds_vao.drop_target_register(DND_FILES)
            self.ds_vao.dnd_bind("<<Drop>>", self._tha_tep)

        # Buoc 2: dau ra
        b2 = ttk.LabelFrame(phai, text="Bước 2 · Đầu ra: Template từng lô (bổ sung ô thiếu, Ctrl+S, chạy ECUS)",
                            padding=8)
        b2.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        b2.columnconfigure(0, weight=1)
        b2.rowconfigure(0, weight=1)
        cot_bang = ("file", "trang_thai", "thieu", "dia_diem", "booking", "da_khai", "sua")
        self.bang = ttk.Treeview(b2, columns=cot_bang, show="headings", selectmode="extended", height=8)
        for c, t, w, st in (("file", "Template", 170, False), ("trang_thai", "Trạng thái", 90, False),
                            ("thieu", "Còn thiếu", 240, True), ("dia_diem", "Địa điểm (cuối · xếp · BT)", 220, False),
                            ("booking", "Booking", 210, False),
                            ("da_khai", "Đã khai ECUS", 120, False), ("sua", "Sửa lúc", 90, False)):
            self.bang.heading(c, text=t, anchor="w")
            self.bang.column(c, width=w, stretch=st, anchor="w")
        self.bang.tag_configure("du", background=MAU["du"])
        self.bang.tag_configure("thieu", background=MAU["thieu"])
        self.bang.tag_configure("kiem", background=MAU["kiem"])
        self.bang.tag_configure("da_khai", background=MAU["da_khai"], foreground=MAU["chu_mo"])
        thanh = ttk.Scrollbar(b2, orient="vertical", command=self.bang.yview)
        self.bang.configure(yscrollcommand=thanh.set)
        self.bang.grid(row=0, column=0, sticky="nsew")
        thanh.grid(row=0, column=1, sticky="ns")
        self.bang.bind("<Double-1>", lambda e: self._mo_excel())
        self.bang.bind("<<TreeviewSelect>>", lambda e: self._khi_chon_lo())

        # Booking cho lo dang chon (thay cho go C3 / D12 trong Excel)
        bk = ttk.Frame(b2, padding=(0, 8, 0, 0))
        bk.grid(row=1, column=0, columnspan=2, sticky="ew")
        ttk.Label(bk, text="Booking lô đã chọn:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, padx=(0, 8))
        ttk.Label(bk, text="Phương thức").grid(row=0, column=1)
        self.bk_pt = ttk.Combobox(bk, width=24, state="readonly", values=[f"{m} · {t}" for m, t in PT_VC])
        self.bk_pt.grid(row=0, column=2, padx=(4, 10))
        ttk.Label(bk, text="Tên tàu / chuyến").grid(row=0, column=3)
        self.bk_ten = ttk.Entry(bk, width=20)
        self.bk_ten.grid(row=0, column=4, padx=(4, 10))
        ttk.Label(bk, text="Ngày đi").grid(row=0, column=5)
        self.bk_ngay = ttk.Entry(bk, width=12)
        self.bk_ngay.grid(row=0, column=6, padx=(4, 10))
        ttk.Button(bk, text="Áp cho lô đã chọn", command=self._ap_booking).grid(row=0, column=7)
        ttk.Button(bk, text="Bỏ booking", command=self._bo_booking).grid(row=0, column=8, padx=(6, 0))

        # Ma dia diem (thay cho go C10 / C11 / C32 trong Excel)
        dd = ttk.Frame(b2, padding=(0, 6, 0, 0))
        dd.grid(row=2, column=0, columnspan=2, sticky="ew")
        ttk.Label(dd, text="Địa điểm lô đã chọn:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, padx=(0, 8))
        self.o_dd = {}
        for i, (khoa, nhan) in enumerate((("dd_cuoi", "Nhận cuối cùng"), ("dd_xep", "Xếp hàng"),
                                          ("dd_bt", "Bảo thuế"))):
            ttk.Label(dd, text=nhan).grid(row=0, column=1 + i * 3, padx=(0 if i == 0 else 10, 4))
            cb = ttk.Combobox(dd, width=21)
            cb.grid(row=0, column=2 + i * 3)
            cb.bind("<KeyRelease>", lambda e, k=khoa: self._loc_ma(k, e))
            cb.bind("<<ComboboxSelected>>", lambda e, k=khoa: self._doi_ma(k))
            cb.bind("<FocusOut>", lambda e, k=khoa: self._doi_ma(k))
            ten = ttk.Entry(dd, width=12)
            ten.grid(row=0, column=3 + i * 3, padx=(3, 0))
            self.o_dd[khoa] = (cb, ten)
        ttk.Button(dd, text="Áp", command=self._ap_dia_diem).grid(row=0, column=10, padx=(10, 0))
        ttk.Button(dd, text="Bỏ", command=self._bo_dia_diem).grid(row=0, column=11, padx=(6, 0))
        self.giu = tk.BooleanVar(value=True)   # giu booking + dia diem tren o nhap khi chon lo khac
        ttk.Checkbutton(b2, variable=self.giu, text="Giữ booking + địa điểm cho lô sau (chọn lô khác không xóa ô "
                        "nhập; nhớ cả khi mở lại chương trình)").grid(row=5, column=0, columnspan=2, sticky="w",
                                                                       pady=(4, 0))
        ttk.Label(b2, style="Phu.TLabel",
                  text="Chọn nhiều lô (Ctrl / Shift + click) rồi áp một lần. Gõ vài chữ của mã hoặc tên rồi mở "
                       "danh sách để lọc. Ô nhỏ cạnh mã = tên địa điểm, chỉ cần khi mã ZZZ (OTHER/UNKNOWN). "
                       "Nhập ở đây thì bỏ qua ô tương ứng trong Excel.").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(4, 0))

        duoi2 = ttk.Frame(b2)
        duoi2.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        duoi2.columnconfigure(4, weight=1)
        ttk.Button(duoi2, text="Mở Excel", command=self._mo_excel).grid(row=0, column=0)
        ttk.Button(duoi2, text="Mở thư mục", command=lambda: self._mo_con(DAU_RA)).grid(row=0, column=1, padx=6)
        ttk.Button(duoi2, text="Kiểm tra lại", command=self._kiem_lai).grid(row=0, column=2)
        ttk.Button(duoi2, text="Xóa (Delete)", command=self._xoa_lo).grid(row=0, column=3, padx=6)
        self.bang.bind("<Delete>", lambda e: self._xoa_lo())
        self.nhap_hang = tk.BooleanVar(value=ecus.IMPORT_GOODS)
        ttk.Checkbutton(duoi2, text="Nhập dòng hàng (Sheet3)", variable=self.nhap_hang,
                        command=lambda: setattr(ecus, "IMPORT_GOODS", self.nhap_hang.get())).grid(
            row=0, column=4, sticky="w", padx=12)
        self.nut_tat_ca = ttk.Button(duoi2, text="▶▶  Chạy tất cả lô đủ", command=self._chay_tat_ca)
        self.nut_tat_ca.grid(row=0, column=5, sticky="e", padx=(0, 6))
        self.nut_ecus = ttk.Button(duoi2, text="▶  Chạy ECUS", style="Chinh.TButton", command=self._chay_ecus)
        self.nut_ecus.grid(row=0, column=6, sticky="e")

        # Nhat ky
        b3 = ttk.LabelFrame(phai, text="Nhật ký", padding=(8, 4))
        b3.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        b3.columnconfigure(0, weight=1)
        self.nhat_ky = scrolledtext.ScrolledText(b3, height=6, font=("Consolas", 9), relief="flat",
                                                 state="disabled", wrap="word", background="#ffffff")
        self.nhat_ky.grid(row=0, column=0, sticky="ew")

        # Thanh trang thai
        cuoi = ttk.Frame(r, padding=(12, 2, 12, 6))
        cuoi.grid(row=2, column=0, sticky="ew")
        cuoi.columnconfigure(0, weight=1)
        self.nhan_viec = ttk.Label(cuoi, style="Phu.TLabel", text="Sẵn sàng.")
        self.nhan_viec.grid(row=0, column=0, sticky="w")
        self.tien_trinh = ttk.Progressbar(cuoi, mode="indeterminate", length=160)
        self.tien_trinh.grid(row=0, column=1, padx=8)
        ttk.Label(cuoi, style="Phu.TLabel", text=f"Bee Logistics · v{PHIEN_BAN}").grid(row=0, column=2)
        self._cap_nhat_nut()

    # ---------- khach hang ----------
    def _nap_khach(self, chon_ten=None):
        self.ds_ten = []
        self.ds_khach.delete(0, "end")
        for ten in chay.danh_sach_ho_so():
            try:
                hs = chay.nap_ho_so(ten)
            except Exception as e:
                print(f"Khong nap duoc ho so {ten}: {e}")
                continue
            if not hasattr(hs, "THU_MUC"):
                continue
            self.ds_ten.append(ten)
            self.ds_khach.insert("end", "  " + getattr(hs, "TEN", ten).replace(" (dung main.py)", ""))
        if self.ds_ten:
            i = self.ds_ten.index(chon_ten) if chon_ten in self.ds_ten else 0
            self.ds_khach.selection_clear(0, "end")
            self.ds_khach.selection_set(i)
            self.ds_khach.see(i)
            self._chon_khach()

    def _chon_khach(self):
        chon = self.ds_khach.curselection()
        if not chon:
            return
        ten = self.ds_ten[chon[0]]
        if ten == self.ten_khach:
            return
        self.ten_khach = ten
        self.khach = chay.nap_ho_so(ten)
        for con in (DAU_VAO, DAU_RA):
            (self._thu_muc() / con).mkdir(parents=True, exist_ok=True)
        tpl = self._tpl()
        if not tpl.exists():
            print(f"[CHU Y] Chua co {tpl.name} trong {self._thu_muc()}")
            if TEMPLATE_CHUAN.exists() and messagebox.askyesno(
                    TEN_APP, f"Khách {getattr(self.khach, 'TEN', ten)} chưa có {tpl.name}.\n\nChép Template chuẩn vào thư mục khách?"):
                shutil.copy2(TEMPLATE_CHUAN, tpl)
                print(f"Da chep Template chuan vao {self._thu_muc()}")
        self.trang_thai.clear()
        self.bang.delete(*self.bang.get_children())
        self._lam_moi_vao()
        self._lam_moi_ra()
        self._nap_ds_ma()
        self._nap_nho_o()
        if ten not in self.danh_muc:
            khach = self.khach

            def nap():
                tpl = chay_chung.template_goc(khach)
                try:
                    self.danh_muc[ten] = ecus.doc_danh_muc(tpl) if tpl.exists() else None
                except Exception as e:
                    print(f"[CANH BAO] Khong doc duoc danh muc ma (Sheet4): {e}")
                self.q.put(("danh_muc", ten))
            threading.Thread(target=nap, daemon=True).start()

    # ---------- them / cai dat khach ----------
    def _mo_template_goc(self):
        if not self.khach:
            return
        f = self._tpl()
        if f.exists():
            mo_tep(f)
        else:
            messagebox.showinfo(TEN_APP, f"Chưa có {f}")

    def _mo_ho_so(self):
        if not self.khach:
            return
        f = Path(self.khach.__file__)
        if LA_WINDOWS:
            subprocess.Popen(["notepad.exe", str(f)])
        else:
            mo_tep(f)
        messagebox.showinfo(TEN_APP, f"Đã mở hồ sơ {f.name}.\n\nSửa phần CẤU HÌNH, lưu lại, rồi tắt và mở lại "
                                     "ECUS Auto để áp dụng.\nHướng dẫn chi tiết: tài liệu 'Thêm khách hàng mới'.")

    def _them_khach(self):
        HopThemKhach(self)

    def tao_khach(self, ten, loai_tep, cach_doc, goi_y, nguon_template):
        """Tao khach moi: ho so .py + thu muc + Template. Tra ve (ten module) hoac raise ValueError."""
        ten = " ".join(ten.split())
        if not ten:
            raise ValueError("Chưa nhập tên khách.")
        ma = khong_dau(ten).lower()
        ma = "".join(c if c.isalnum() else "_" for c in ma).strip("_")
        while "__" in ma:
            ma = ma.replace("__", "_")
        if not ma or ma[0].isdigit():
            ma = "kh_" + ma
        thu_muc_ten = f"{ten} xuat"
        file_ho_so = THU_MUC_HO_SO / f"{ma}.py"
        if file_ho_so.exists() or ma in chay.danh_sach_ho_so():
            raise ValueError(f"Đã có hồ sơ '{ma}'. Chọn tên khác.")
        mau = (THU_MUC_HO_SO / "_mau_khach_moi.py").read_text(encoding="utf-8")
        mau_file, loai = LOAI_HOA_DON[loai_tep]
        thay = [
            ('HO SO KHACH: <TEN KHACH>   <-- FILE MAU. Copy thanh ho_so/<ten_khach>.py (khong co dau gach duoi o dau) roi sua.',
             f'HO SO KHACH: {khong_dau(ten).upper()}  (tao tu ECUS Auto ngay {datetime.now():%d/%m/%Y}). '
             'Sua phan CAU HINH ben duoi.'),
            ('TEN = "<Ten khach> xuat"', f'TEN = "{thu_muc_ten}"'),
            ('THU_MUC = THU_MUC_DON / "<Ten khach> xuat"', f'THU_MUC = THU_MUC_DON / "{thu_muc_ten}"'),
            ('MAU_FILE = "**/*.docx"', f'MAU_FILE = "{mau_file}"'),
            ('LOAI_TEP = [("Hoa don", "*.docx *.xlsx *.pdf")]', f'LOAI_TEP = [("Hoa don", "{loai}")]'),
            ('TEMPLATE = THU_MUC / "Template_<ma>.xlsx"', f'TEMPLATE = THU_MUC / "Template_{ma}.xlsx"'),
            ('CACH_DOC = "ai"               #', f'CACH_DOC = "{cach_doc}"' + " " * (15 - len(cach_doc)) + "#"),
        ]
        for cu, moi in thay:
            if cu not in mau:
                raise ValueError(f"File mẫu _mau_khach_moi.py đã bị sửa, không thấy dòng: {cu[:40]}")
            mau = mau.replace(cu, moi, 1)
        if goi_y.strip():
            dau = mau.index("GOI_Y_AI = (")
            cuoi = mau.index(")\n", dau) + 2
            mau = mau[:dau] + f"GOI_Y_AI = {goi_y.strip()!r}\n" + mau[cuoi:]
        compile(mau, str(file_ho_so), "exec")          # bao loi som neu ten / goi y lam hong cu phap
        thu_muc = chay_chung.THU_MUC_DON / thu_muc_ten
        for con in (DAU_VAO, DAU_RA):
            (thu_muc / con).mkdir(parents=True, exist_ok=True)
        tpl = thu_muc / f"Template_{ma}.xlsx"
        if not tpl.exists():
            shutil.copy2(nguon_template, tpl)
        file_ho_so.write_text(mau, encoding="utf-8")
        importlib.invalidate_caches()
        print(f"Da tao khach moi '{ten}': ho so {file_ho_so.name}, thu muc {thu_muc}, Template tu {Path(nguon_template).name}")
        return ma

    def _tpl(self):
        """Template goc cua khach dang chon: Template_<khach>.xlsx (tu doi ten file cu Template.xlsx)."""
        return Path(chay_chung.template_goc(self.khach))

    def _thu_muc(self):
        return Path(self.khach.THU_MUC)

    def _mo_thu_muc_khach(self):
        if self.khach:
            mo_tep(self._thu_muc())

    def _mo_con(self, con):
        if self.khach:
            mo_tep(self._thu_muc() / con)

    # ---------- buoc 1: dau vao ----------
    def _tep_vao(self):
        d = self._thu_muc() / DAU_VAO
        return sorted(f for f in d.iterdir() if f.is_file() and not f.name.startswith("~$")) if d.exists() else []

    def _lam_moi_vao(self):
        ds = [f.name for f in self._tep_vao()]
        if list(self.ds_vao.get(0, "end")) != ds:
            self.ds_vao.delete(0, "end")
            for n in ds:
                self.ds_vao.insert("end", n)
        self.nut_tao.configure(text=f"▶  Tạo Template ({len(ds)} tệp)" if ds else "▶  Tạo Template")
        self._cap_nhat_nut()

    def _them_tep(self):
        if not self.khach:
            return
        loai = list(getattr(self.khach, "LOAI_TEP", [])) + [("Tất cả", "*.*")]
        cac = filedialog.askopenfilenames(title="Chọn hóa đơn đưa vào", filetypes=loai)
        self._chep_vao(cac)

    def _tha_tep(self, event):
        self._chep_vao(self.root.tk.splitlist(event.data))

    def _chep_vao(self, cac):
        if not cac or not self.khach:
            return
        dich = self._thu_muc() / DAU_VAO
        n = 0
        for c in cac:
            p = Path(c)
            if p.is_file():
                shutil.copy2(p, dich / p.name)
                n += 1
        print(f"Da dua {n} tep vao {dich}")
        self._lam_moi_vao()

    def _bo_tep(self):
        chon = [self.ds_vao.get(i) for i in self.ds_vao.curselection()]
        if not chon:
            return
        if not messagebox.askyesno(TEN_APP, f"Bỏ {len(chon)} tệp khỏi đầu vào?\n(Tệp gốc của bạn không bị xóa.)"):
            return
        for n in chon:
            (self._thu_muc() / DAU_VAO / n).unlink(missing_ok=True)
        self._lam_moi_vao()

    def _tao(self):
        tep = self._tep_vao()
        if not tep or self.ban:
            return
        if hasattr(self.khach, "du_kien_dau_ra"):
            se_ghi_de = [p for p in self.khach.du_kien_dau_ra(tep) if Path(p).exists()]
            if se_ghi_de:
                ds = "\n".join("• " + Path(p).name for p in se_ghi_de[:12])
                if not messagebox.askyesno(TEN_APP, f"Các Template sau đã có, sẽ bị tạo lại (mất phần đã sửa tay):\n\n"
                                                    f"{ds}\n\nVẫn tạo lại?"):
                    return
        khach, thu_muc = self.khach, self._thu_muc()

        def viec():
            return chay.tao_template(khach, tep)

        def xong(kq):
            if not kq:
                return
            da_tao, thieu, so_lo = kq
            if da_tao and len(da_tao) >= so_lo:
                dich = thu_muc / DA_XU_LY / datetime.now().strftime("%Y-%m-%d %H%M")
                dich.mkdir(parents=True, exist_ok=True)
                for f in tep:
                    if f.exists():
                        shutil.move(str(f), dich / f.name)
                print(f"Da chuyen {len(tep)} tep dau vao sang {dich}")
            elif so_lo:
                print(f"[CHU Y] Chi tao duoc {len(da_tao)}/{so_lo} lo. Tep dau vao giu nguyen, xem nhat ky o tren.")
            self._lam_moi_vao()
            self._lam_moi_ra(chon=[Path(p) for p in da_tao])
            n_thieu = len(thieu or [])
            self._bao(f"Đã tạo {len(da_tao)} Template · {len(da_tao) - n_thieu} đủ · {n_thieu} cần bổ sung.")

        self._chay_nen("Đang tạo Template...", viec, xong)

    # ---------- buoc 2: dau ra ----------
    def _tep_ra(self):
        d = self._thu_muc() / DAU_RA
        if not d.exists():
            return []
        ds = [f for f in d.glob("Template *.xlsx") if not f.name.startswith("~$")]
        return sorted(ds, key=lambda f: f.stat().st_mtime, reverse=True)

    def _lam_moi_ra(self, chon=None):
        if not self.khach:
            return
        lich_su = (ecus.doc_lich_su(self._thu_muc()), ecus.doc_booking(self._thu_muc()))
        tep = self._tep_ra()
        hien = set(self.bang.get_children())
        con = set()
        for f in tep:
            iid = str(f)
            con.add(iid)
            mtime = f.stat().st_mtime
            cu = self.trang_thai.get(iid)
            if cu is None or cu[0] != mtime:
                self._luot += 1
                self.trang_thai[iid] = (mtime, None, {}, self._luot)
                self.q_kiem.put((self.ten_khach, self.khach, f, mtime, self._luot))
            gia_tri = self._dong(f, mtime, lich_su)
            if iid in hien:
                self.bang.item(iid, values=gia_tri[0], tags=gia_tri[1])
            else:
                self.bang.insert("", "end", iid=iid, values=gia_tri[0], tags=gia_tri[1])
        for iid in hien - con:
            self.bang.delete(iid)
        # giu thu tu moi nhat len dau
        for i, f in enumerate(tep):
            self.bang.move(str(f), "", i)
        if chon:
            for p in chon:
                if str(p) in con:
                    self.bang.selection_set(str(p))
                    self.bang.see(str(p))
                    break
        self._cap_nhat_nut()

    def _dong(self, f, mtime, lich_su):
        lich_su, booking = lich_su
        thieu = self.trang_thai.get(str(f), (None, None))[1]
        da_khai = lich_su.get(f.name, "")
        b = booking.get(f.name) or {}
        bk = " · ".join(x for x in (str(b.get("pt_ma", "")), b.get("pt_ten", ""), b.get("ngay_di", "")) if x) \
            if b.get("pt_ma") else "theo Excel"
        gt = self._gia_tri(str(f))
        dd = " · ".join(gt.get(k, "—") for k in ("dd_cuoi", "dd_xep", "dd_bt")) if gt else ""
        if any(b.get(k) for k in ("dd_cuoi", "dd_xep", "dd_bt")):
            dd += "  (giao diện)"
        sua = datetime.fromtimestamp(mtime).strftime("%d/%m %H:%M")
        mo = dang_mo_trong_excel(f)
        if thieu is None:
            tt, chi_tiet, tag = "Đang kiểm tra…", "", "kiem"
        elif thieu:
            tt, tag = f"THIẾU ({len(thieu)})", "thieu"
            chi_tiet = "; ".join(t.split(":")[0] for t in thieu)
        else:
            tt, chi_tiet, tag = "ĐỦ", "", "du"
        if da_khai and tag == "du":
            tag = "da_khai"
        if mo:
            chi_tiet = ("Đang mở trong Excel, nhớ Ctrl+S. " + chi_tiet).strip()
        return (f.name.replace(".xlsx", ""), tt, chi_tiet, dd, bk, da_khai, sua), (tag,)

    def _luong_kiem_tra(self):
        """Luong nen: kiem tra tung Template (doc Sheet4 cua khach 1 lan)."""
        while True:
            ten, khach, f, mtime, luot = self.q_kiem.get()
            try:
                if ten not in self.danh_muc:
                    tpl = chay_chung.template_goc(khach)
                    self.danh_muc[ten] = ecus.doc_danh_muc(tpl) if tpl.exists() else None
                thieu, gia_tri = ecus.kiem_tra_template(f, self.danh_muc[ten])
                self.q.put(("trang_thai", str(f), mtime, thieu, gia_tri, luot))
            except Exception as e:
                self.q.put(("trang_thai", str(f), mtime, [f"Khong doc duoc file: {e}"], {}, luot))

    def _kiem_lai(self):
        self.trang_thai.clear()
        self.danh_muc.pop(self.ten_khach, None)
        self._lam_moi_ra()

    def _dang_chon(self):
        chon = self.bang.selection()
        return Path(chon[0]) if chon else None

    def _mo_excel(self):
        p = self._dang_chon()
        if p and p.exists():
            mo_tep(p)

    def _chay_ecus(self):
        p = self._dang_chon()
        if not p or self.ban:
            return
        thieu = self.trang_thai.get(str(p), (None, None))[1]
        if thieu:
            messagebox.showwarning(TEN_APP, f"{p.name} còn thiếu:\n\n" + "\n".join("• " + t for t in thieu)
                                   + "\n\nMở Excel, bổ sung, Ctrl+S rồi chạy lại.")
            return
        if not LA_WINDOWS:
            messagebox.showinfo(TEN_APP, "Chạy ECUS chỉ làm được trên Windows (máy có ECUS).")
            return
        if dang_mo_trong_excel(p) and not messagebox.askyesno(
                TEN_APP, f"{p.name} đang mở trong Excel.\nĐã lưu (Ctrl+S) chưa? Bấm Có để tiếp tục."):
            return
        if not messagebox.askyesno(TEN_APP, f"Điền {p.name} vào ECUS?\n\n"
                                            "• Mở sẵn ECUS ở tờ khai cần điền.\n"
                                            "• Trong lúc chạy, đừng chạm chuột / bàn phím.\n"
                                            "• Script không tự bấm Ghi / Gửi tờ khai."):
            return
        self._dien_mot_lo(p, lambda kq: self._bao("Đã điền xong vào ECUS. Kiểm tra lại trên ECUS rồi bấm Ghi."
                                                  if kq else "Chưa điền xong vào ECUS, xem nhật ký."))

    def _dien_mot_lo(self, p, sau_khi_xong):
        """Thu nho cua so, dien 1 lo vao ECUS o luong nen, xong thi hien lai cua so va goi sau_khi_xong(kq)."""
        def viec():
            return ecus.chay_ecus(p, xac_nhan=lambda cau_hoi: True)

        def xong(kq):
            self.root.deiconify()
            self.root.lift()
            self._lam_moi_ra(chon=[p])
            sau_khi_xong(kq)

        self.root.iconify()
        self._chay_nen(f"Đang điền {p.name} vào ECUS... đừng chạm chuột / bàn phím", viec, xong)

    # ---------- chay nhieu lo ----------
    def _lo_cho_khai(self):
        """Cac lo DU va chua khai, theo thu tu tren bang."""
        lich_su = ecus.doc_lich_su(self._thu_muc())
        return [Path(i) for i in self.bang.get_children()
                if self.trang_thai.get(i, (None, None))[1] == [] and Path(i).name not in lich_su]

    def _chay_tat_ca(self):
        if self.ban or not self.khach:
            return
        if not LA_WINDOWS:
            messagebox.showinfo(TEN_APP, "Chạy ECUS chỉ làm được trên Windows (máy có ECUS).")
            return
        ds = self._lo_cho_khai()
        if not ds:
            messagebox.showinfo(TEN_APP, "Không có lô nào ĐỦ mà chưa khai.")
            return
        dang_mo = [p.name for p in ds if dang_mo_trong_excel(p)]
        ten = "\n".join(f"  {i + 1}. {p.stem}" for i, p in enumerate(ds))
        canh_bao = ("\n\nĐang mở trong Excel (nhớ Ctrl+S): " + ", ".join(dang_mo)) if dang_mo else ""
        if not messagebox.askyesno(TEN_APP, f"Chạy lần lượt {len(ds)} lô:\n\n{ten}{canh_bao}\n\n"
                                            "Trước mỗi lô, chương trình dừng lại để bạn mở / copy tờ khai "
                                            "cho lô đó trên ECUS rồi bấm OK. Bấm Hủy bất cứ lúc nào để dừng."):
            return
        self._lo_tiep(ds, 0, [], [])

    def _lo_tiep(self, ds, i, xong, loi):
        if i >= len(ds):
            tom_tat = f"Đã điền {len(xong)}/{len(ds)} lô."
            if loi:
                tom_tat += "\nChưa xong: " + ", ".join(p.stem for p in loi) + " (xem nhật ký)."
            messagebox.showinfo(TEN_APP, tom_tat + "\n\nNhớ kiểm tra và bấm Ghi lô cuối trên ECUS.")
            self._bao(tom_tat.split("\n")[0])
            print(f"\n=== Chay nhieu lo: {tom_tat} ===")
            return
        p = ds[i]
        truoc = (f"Lô trước ({ds[i - 1].stem}) đã điền xong: kiểm tra và bấm GHI trên ECUS trước.\n\n"
                 if i > 0 and ds[i - 1] in xong else "")
        if not messagebox.askokcancel(TEN_APP, f"{truoc}Lô {i + 1}/{len(ds)}: {p.stem}\n\n"
                                               "Trên ECUS: mở (hoặc copy từ lô cũ) tờ khai cho lô này.\n"
                                               "Xong bấm OK để điền. Hủy = dừng chạy nhiều lô."):
            print(f"\nDa dung o lo {i + 1}/{len(ds)} ({p.stem}).")
            self._bao(f"Đã dừng. Điền xong {len(xong)}/{len(ds)} lô.")
            return
        thieu = self.trang_thai.get(str(p), (None, None))[1]
        if thieu != []:                        # file vua bi sua, khong con du
            print(f"\n[BO QUA] {p.name}: khong con du thong tin.")
            self._lo_tiep(ds, i + 1, xong, loi + [p])
            return
        print(f"\n=== Lo {i + 1}/{len(ds)}: {p.name} ===")

        def sau(kq):
            self._lo_tiep(ds, i + 1, xong + [p] if kq else xong, loi if kq else loi + [p])

        self._dien_mot_lo(p, sau)

    # ---------- booking ----------
    def _cac_chon(self):
        return [Path(i) for i in self.bang.selection()]

    def _khi_chon_lo(self):
        """Chon 1 lo co booking -> hien booking do len o nhap de xem / sua."""
        chon = self._cac_chon()
        if len(chon) == 1 and self.khach and not self.giu.get():
            b = ecus.doc_booking(self._thu_muc()).get(chon[0].name) or {}
            gt = self._gia_tri(str(chon[0]))
            for khoa, (cb, o_ten) in self.o_dd.items():
                ma = b.get(khoa) or gt.get(khoa, "")
                cb.set(self._hien_ma(khoa, ma) if ma else "")
                o_ten.delete(0, "end")
                o_ten.insert(0, b.get(khoa + "_ten") or gt.get(khoa + "_ten", ""))
            if b.get("pt_ma"):
                ma = str(b.get("pt_ma", ""))
                self.bk_pt.set(next((f"{m} · {t}" for m, t in PT_VC if m == ma), ""))
                for o, k in ((self.bk_ten, "pt_ten"), (self.bk_ngay, "ngay_di")):
                    o.delete(0, "end")
                    o.insert(0, b.get(k, ""))
        self._cap_nhat_nut()

    def _ap_booking(self):
        chon = self._cac_chon()
        if not chon:
            messagebox.showinfo(TEN_APP, "Chọn lô (một hoặc nhiều dòng) ở bảng trên trước.")
            return
        ngay = self.bk_ngay.get().strip()
        if ngay:                               # cho go nhanh: 5/11, 05/11/26, 5-11-2026
            ngay = ngay.replace("-", "/").replace(".", "/")
            phan = ngay.split("/")
            if len(phan) == 2:
                phan.append(str(datetime.now().year))
            if len(phan) == 3 and len(phan[2]) == 2:
                phan[2] = "20" + phan[2]
            try:
                ngay = datetime(int(phan[2]), int(phan[1]), int(phan[0])).strftime("%d/%m/%Y")
            except (ValueError, IndexError):
                messagebox.showwarning(TEN_APP, f"Ngày đi '{self.bk_ngay.get()}' không đúng. Gõ dạng 05/11/2026.")
                return
        b = {"pt_ma": self.bk_pt.get().split(" ")[0], "pt_ten": self.bk_ten.get().strip().upper(), "ngay_di": ngay}
        loi = ecus.kiem_tra_booking(b)
        if loi:
            messagebox.showwarning(TEN_APP, "\n".join(loi))
            return
        ecus.ghi_booking(self._thu_muc(), [p.name for p in chon], b)
        for o, v in ((self.bk_ngay, ngay), (self.bk_ten, b["pt_ten"])):
            o.delete(0, "end")
            o.insert(0, v)
        self._luu_nho_o()
        print(f"Da ap booking {b['pt_ma']} / {b['pt_ten']} / {ngay or '(khong ngay)'} cho {len(chon)} lo: "
              + ", ".join(p.stem for p in chon))
        self._kiem_lai_cac(chon)

    def _bo_booking(self):
        chon = self._cac_chon()
        if not chon or not messagebox.askyesno(TEN_APP, f"Bỏ booking của {len(chon)} lô? "
                                                         "(Lô đó sẽ lấy lại B3/C3/D12 trong Excel.)"):
            return
        ecus.ghi_booking(self._thu_muc(), [p.name for p in chon], xoa_khoa=ecus.KHOA_BOOKING)
        self._kiem_lai_cac(chon)

    # ---------- nho o nhap cho lo sau ----------
    FILE_NHO_O = "nho_giao_dien.json"

    def _o_nhap(self):
        """(ten, o) cua moi o booking + dia diem."""
        ds = [("pt", self.bk_pt), ("pt_ten", self.bk_ten), ("ngay_di", self.bk_ngay)]
        for khoa, (cb, o_ten) in self.o_dd.items():
            ds += [(khoa, cb), (khoa + "_ten", o_ten)]
        return ds

    def _luu_nho_o(self):
        try:
            gt = {k: o.get() for k, o in self._o_nhap()}
            (self._thu_muc() / self.FILE_NHO_O).write_text(json.dumps(gt, ensure_ascii=False, indent=2),
                                                           encoding="utf-8")
        except Exception as e:
            print(f"[CANH BAO] Khong luu duoc o nhap: {e}")

    def _nap_nho_o(self):
        """Doi khach / mo lai chuong trinh: dien lai booking + dia diem lan ap gan nhat cua khach nay."""
        try:
            f = self._thu_muc() / self.FILE_NHO_O
            gt = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
        except Exception:
            gt = {}
        for k, o in self._o_nhap():
            v = gt.get(k, "")
            if isinstance(o, ttk.Combobox):
                o.set(v)
            else:
                o.delete(0, "end")
                o.insert(0, v)

    # ---------- ma dia diem ----------
    def _gia_tri(self, iid):
        t = self.trang_thai.get(iid)
        return t[2] if t and len(t) > 2 and t[2] else {}

    def _dm(self, khoa):
        dm = self.danh_muc.get(self.ten_khach) or {}
        return dm.get(khoa) or {}

    def _hien_ma(self, khoa, ma):
        ten = self._dm(khoa).get(ma, "")
        return f"{ma} · {ten}" if ten else ma

    def _nap_ds_ma(self):
        for khoa, (cb, _) in self.o_dd.items():
            cb["values"] = [f"{m} · {t}" for m, t in list(self._dm(khoa).items())[:400]]

    def _loc_ma(self, khoa, event):
        """Go vai chu -> danh sach chi con ma / ten co chua chu do."""
        if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
            return
        cb = self.o_dd[khoa][0]
        go = cb.get().strip().upper()
        ds = [f"{m} · {t}" for m, t in self._dm(khoa).items() if not go or go in m or go in t.upper()]
        cb["values"] = ds[:400]

    def _ma_tu_o(self, khoa):
        return self.o_dd[khoa][0].get().split(" · ")[0].strip().upper()

    def _la_zzz(self, khoa, ma):
        return self._dm(khoa).get(ma, "").upper() in ecus.TEN_CAN_GO_TAY

    def _doi_ma(self, khoa):
        """Chon xong 1 ma: ma ZZZ thi nhac go ten, ma thuong thi xoa o ten."""
        ma = self._ma_tu_o(khoa)
        o_ten = self.o_dd[khoa][1]
        if ma and not self._la_zzz(khoa, ma) and o_ten.get().strip():
            o_ten.delete(0, "end")

    def _ap_dia_diem(self):
        chon = self._cac_chon()
        if not chon:
            messagebox.showinfo(TEN_APP, "Chọn lô (một hoặc nhiều dòng) ở bảng trên trước.")
            return
        cap_nhat, loi = {}, []
        for khoa, nhan in (("dd_cuoi", "Nhận cuối cùng"), ("dd_xep", "Xếp hàng"), ("dd_bt", "Bảo thuế")):
            ma = self._ma_tu_o(khoa)
            if not ma:
                continue                       # o trong = giu nguyen nhu hien tai
            dm = self._dm(khoa)
            if dm and ma not in dm:
                loi.append(f"{nhan}: mã '{ma}' không có trong danh mục (Sheet4)")
                continue
            ten = self.o_dd[khoa][1].get().strip().upper()
            if self._la_zzz(khoa, ma) and not ten:
                loi.append(f"{nhan}: mã {ma} là ZZZ, gõ tên địa điểm vào ô nhỏ bên cạnh")
                continue
            cap_nhat[khoa] = ma
            cap_nhat[khoa + "_ten"] = ten if self._la_zzz(khoa, ma) else ""
        if loi:
            messagebox.showwarning(TEN_APP, "\n".join(loi))
            return
        if not cap_nhat:
            messagebox.showinfo(TEN_APP, "Chưa chọn mã nào.")
            return
        ecus.ghi_booking(self._thu_muc(), [p.name for p in chon], cap_nhat)
        self._luu_nho_o()
        print(f"Da ap dia diem {', '.join(v for k, v in cap_nhat.items() if not k.endswith('_ten'))} cho "
              f"{len(chon)} lo: " + ", ".join(p.stem for p in chon))
        self._kiem_lai_cac(chon)

    def _bo_dia_diem(self):
        chon = self._cac_chon()
        if not chon or not messagebox.askyesno(TEN_APP, f"Bỏ mã địa điểm đã nhập ở đây cho {len(chon)} lô? "
                                                         "(Lô đó sẽ lấy lại C10/C11/C32 trong Excel.)"):
            return
        ecus.ghi_booking(self._thu_muc(), [p.name for p in chon], xoa_khoa=ecus.KHOA_DIA_DIEM)
        self._kiem_lai_cac(chon)

    # ---------- xoa lo ----------
    def _xoa_lo(self):
        chon = self._cac_chon()
        if not chon or self.ban:
            return
        dang_mo = [p.name for p in chon if dang_mo_trong_excel(p)]
        if dang_mo:
            messagebox.showwarning(TEN_APP, "Đóng file trong Excel trước khi xóa:\n" + "\n".join(dang_mo))
            return
        ds = "\n".join("• " + p.stem for p in chon[:15]) + (f"\n… và {len(chon) - 15} lô khác" if len(chon) > 15 else "")
        if not messagebox.askyesno(TEN_APP, f"Xóa {len(chon)} Template?\n\n{ds}\n\n"
                                            "File được đưa vào Thùng rác (lấy lại được). "
                                            "Booking / địa điểm / lịch sử đã khai của lô đó cũng bỏ."):
            return
        n = 0
        for p in chon:
            try:
                vao_thung_rac(p)
                n += 1
            except Exception as e:
                print(f"[LOI] Khong xoa duoc {p.name}: {e}")
            self.trang_thai.pop(str(p), None)
        ecus.xoa_khoi_so(self._thu_muc(), [p.name for p in chon])
        print(f"Da xoa {n} Template: " + ", ".join(p.stem for p in chon))
        self._lam_moi_ra()
        self._bao(f"Đã xóa {n} Template.")

    def _kiem_lai_cac(self, cac):
        for p in cac:
            self.trang_thai.pop(str(p), None)
            self.cho_bao.add(str(p))
        self._lam_moi_ra()
        for p in cac:
            if self.bang.exists(str(p)):
                self.bang.selection_add(str(p))

    def _cap_nhat_nut(self):
        co_vao = self.ds_vao.size() > 0 if hasattr(self, "ds_vao") else False
        self.nut_tao.state(["!disabled"] if (co_vao and not self.ban) else ["disabled"])
        if not hasattr(self, "nut_ecus"):
            return
        p = self._dang_chon()
        du = p is not None and self.trang_thai.get(str(p), (None, None))[1] == []
        self.nut_ecus.state(["!disabled"] if (p is not None and not self.ban) else ["disabled"])
        self.nut_ecus.configure(text="▶  Chạy ECUS" if (du or p is None) else "▶  Chạy ECUS (còn thiếu)")
        n = len(self._lo_cho_khai()) if self.khach else 0
        self.nut_tat_ca.configure(text=f"▶▶  Chạy tất cả lô đủ ({n})")
        self.nut_tat_ca.state(["!disabled"] if (n and not self.ban) else ["disabled"])

    # ---------- chay nen ----------
    def _chay_nen(self, mo_ta, viec, xong=None):
        self.ban = True
        self.nhan_viec.configure(text=mo_ta)
        self.tien_trinh.start(12)
        self._cap_nhat_nut()

        def chay():
            kq = None
            try:
                kq = viec()
            except Exception:
                print("\n[LOI]\n" + traceback.format_exc())
            self.q.put(("xong", xong, kq))

        threading.Thread(target=chay, daemon=True).start()

    def _bao(self, chu):
        self.nhan_viec.configure(text=chu)

    def _input_qua_hop_thoai(self, cau_hoi=""):
        """Thay input(): script hoi gi thi hien hop thoai (an toan khi goi tu luong nen)."""
        if threading.current_thread() is threading.main_thread():
            return simpledialog.askstring(TEN_APP, cau_hoi, parent=self.root) or ""
        ket_qua, xong = [""], threading.Event()
        self.q.put(("hoi", cau_hoi, ket_qua, xong))
        xong.wait()
        return ket_qua[0]

    def _xu_ly_hang_doi(self):
        try:
            while True:
                m = self.q.get_nowait()
                if m[0] == "log":
                    self.nhat_ky.configure(state="normal")
                    self.nhat_ky.insert("end", m[1])
                    self.nhat_ky.see("end")
                    self.nhat_ky.configure(state="disabled")
                elif m[0] == "trang_thai":
                    _, iid, mtime, thieu, gia_tri, luot = m
                    cu = self.trang_thai.get(iid)
                    if cu and cu[0] == mtime and cu[-1] == luot:     # bo qua ket qua cua luot kiem tra cu
                        self.trang_thai[iid] = (mtime, thieu, gia_tri, luot)
                        if iid in self.cho_bao:
                            self.cho_bao.discard(iid)
                            print(f"  {Path(iid).stem}: " + ("DU, chay ECUS duoc" if not thieu
                                                              else "con thieu " + "; ".join(thieu)))
                        if self.bang.exists(iid):
                            f = Path(iid)
                            vals, tags = self._dong(f, mtime, (ecus.doc_lich_su(self._thu_muc()),
                                                               ecus.doc_booking(self._thu_muc())))
                            self.bang.item(iid, values=vals, tags=tags)
                        self._cap_nhat_nut()
                elif m[0] == "danh_muc":
                    if m[1] == self.ten_khach:
                        self._nap_ds_ma()
                elif m[0] == "hoi":
                    _, cau_hoi, ket_qua, xong = m
                    ket_qua[0] = simpledialog.askstring(TEN_APP, cau_hoi, parent=self.root) or ""
                    xong.set()
                elif m[0] == "xong":
                    _, ham, kq = m
                    self.ban = False
                    self.tien_trinh.stop()
                    self._bao("Sẵn sàng.")
                    if ham:
                        ham(kq)
                    self._cap_nhat_nut()
        except queue.Empty:
            pass
        self.root.after(100, self._xu_ly_hang_doi)

    def _theo_doi_thu_muc(self):
        """Moi 2,5 giay: cap nhat dau vao / dau ra (luu Excel la trang thai tu doi)."""
        if self.khach and not self.ban:
            try:
                self._lam_moi_vao()
                self._lam_moi_ra()
            except Exception as e:
                print(f"[CANH BAO] {e}")
        self.root.after(2500, self._theo_doi_thu_muc)

    def _theo_doi_ecus(self):
        if LA_WINDOWS:
            try:
                if ecus.user32 is None:
                    ecus.setup_windows()
                self.ecus_san_sang = ecus.find_ecus_window() is not None
            except Exception:
                self.ecus_san_sang = False
            self.nhan_ecus.configure(text="ECUS: đang mở ✓" if self.ecus_san_sang else "ECUS: chưa mở",
                                     fg="#9be29b" if self.ecus_san_sang else "#ffb4a8")
        else:
            self.nhan_ecus.configure(text="ECUS: chỉ chạy trên Windows")
        self.root.after(5000, self._theo_doi_ecus)

    def chay(self):
        print(f"{TEN_APP} v{PHIEN_BAN} · thu muc don hang: {ecus.DON_HANG_DIR}")
        self.root.mainloop()


if __name__ == "__main__":
    UngDung().chay()
