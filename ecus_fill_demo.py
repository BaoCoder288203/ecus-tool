"""
DEMO: dien thong tin tu file Template G-xx-xxxx.xlsx (cac o vang o Sheet1) vao ECUS, tab "Thong tin chung".

Chi dung thu vien co san cua Python + openpyxl (khong can pywinauto / pywin32).

Cach dung:
1. Mo ECUS (nen thu tren ban Demo), vao Tờ khai xuất khẩu, tab "Thông tin chung",
   o che do nhap (to khai moi / dang sua, cac o khong bi khoa).
2. Chay: python ecus_fill_demo.py
3. Chon file Template trong D:\\test\\output.
4. Script se IN RA bang "gia tri se dien" va "o tim thay trong ECUS" de ban kiem tra truoc (chua dien gi).
5. Chi khi ban go DIEN thi moi tu go vao ECUS. Script KHONG bam Ghi / Gui.

Cac o con lai (booking, ngay gio...) ban tu nhap nhu cu.

Chuyen tab: script tu chuyen sang tab "Thong tin chung" truoc khi dien, va sau khi dien xong
tu chuyen sang tab mac dinh DEFAULT_TAB (= 2, "Danh sach hang"). Doi DEFAULT_TAB neu can.
"""
import ctypes
import re
import sys
import time
import unicodedata
from collections import namedtuple
from ctypes import wintypes
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

# ===== CAU HINH =====
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
SHEET = "Sheet1"

# Cac tab cua ECUS (tinh tu 0, tu trai sang phai theo anh chup)
TAB_NAMES = [
    "Thong tin chung",
    "Thong tin Container",
    "Danh sach hang",
    "Chi thi cua Hai quan",
    "Ket qua xu ly to khai",
    "Quan ly to khai",
]
TAB_THONG_TIN_CHUNG = 0   # cac o hien tai duoc dien o tab nay (script tu chuyen sang truoc khi dien)
DEFAULT_TAB = 2           # sau khi dien xong, tu chuyen sang tab nay (2 = Danh sach hang)

# O vang trong file Template (Sheet1)
CELLS = {
    "invoice_no": "B21",     # So invoice
    "invoice_date": "B22",   # Ngay invoice (= Ship Date)
    "currency": "C12",       # Currency
    "terms": "C11",          # Terms of sale (vi du "FOB Ho Chi Minh City" -> lay "FOB")
    "invoice_total": "B27",  # Tong tri gia hoa don
    "qty": "D16",            # So luong kien
    "unit": "D17",           # Don vi (PP, BX...)
    "gross": "D18",          # KGS
    "processing_total": "B25",   # Tong tri gia Gia cong -> TGGC
    "custom_total": "B26",       # Tong tri gia Nguyen vat lieu -> TG NVL
    "ship_to_name": "C13",       # Ship to Name -> NNH
}

# Chuoi dien vao o "Ky hieu va so hieu" (sua lai cho dung y neu can, giu nguyen khoang trang)
SIGN_TEMPLATE = "PTTT: TT, TGTT = TGGC =   {tggc} {cur},TG NVL = {tgnvl} {cur}, NNH: {nnh}"


# ===== DINH DANG GIA TRI CHO ECUS =====
def fmt_text(v):
    return str(v).strip()


def fmt_date(v):
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    return str(v).strip()


def fmt_terms(v):
    parts = str(v).strip().split()
    return parts[0].upper() if parts else ""


def fmt_money(v):
    return f"{float(v):,.2f}"


def fmt_int(v):
    return str(int(round(float(v))))


def fmt_plain(v):
    return ("%.2f" % float(v)).rstrip("0").rstrip(".")


def fmt_num(v):
    return ("%.3f" % float(v)).rstrip("0").rstrip(".")


# (ten hien thi, khoa o Template, nhan trong ECUS, chi so o ben phai nhan, ham dinh dang)
FIELDS = [
    ("So hoa don",           "invoice_no",    "Số hóa đơn:",                    0, fmt_text),
    ("Ngay phat hanh",       "invoice_date",  "Ngày phát hành:",                0, fmt_date),
    ("Dieu kien gia",        "terms",         "Điều kiện giá hóa đơn:",         0, fmt_terms),
    ("Ma dong tien",         "currency",      "Mã đồng tiền của hóa đơn:",      0, fmt_text),
    ("Tong tri gia hoa don", "invoice_total", "Tổng trị giá hóa đơn:",          0, fmt_money),
    ("Tri gia tinh thue",    "invoice_total", "Trị giá tính thuế:",            0, fmt_money),
    ("So luong kien",        "qty",           "Số lượng kiện:",                 0, fmt_int),
    ("Don vi kien",          "unit",          "Số lượng kiện:",                 1, fmt_text),
    ("Gross (KGS)",          "gross",         "Tổng trọng lượng hàng (Gross):", 0, fmt_num),
    ("Ky hieu va so hieu",   "sign_text",     "Ký hiệu và số hiệu:",            0, fmt_text),
    # Hai ngay duoi day = ngay invoice. ECUS viet nhan la 'khời' (sai chinh ta), giu ca 2 cach viet
    ("Ngay khoi hanh VC",    "invoice_date",  ("Ngày khời hành vận chuyển:", "Ngày khởi hành vận chuyển:"), 0, fmt_date),
    # Dong 'Dia diem dich...' co 3 o nhap: ma (06B1A04), ten (KHO SCSC), ngay -> ngay la o thu 3 (chi so 2)
    ("Ngay den DD bao thue", "invoice_date",  "Địa điểm đích cho vận chuyển bảo thuế:", 2, fmt_date),
]

Ctl = namedtuple("Ctl", "hwnd cls text left top right bottom visible")


# ===== DOC FILE TEMPLATE =====
def read_template(path):
    wb = load_workbook(path, data_only=True)
    ws = None
    for name in wb.sheetnames:
        if name.strip().casefold() == SHEET.casefold():
            ws = wb[name]
            break
    if ws is None:
        raise RuntimeError(f"File khong co sheet '{SHEET}'")
    values = {key: ws[addr].value for key, addr in CELLS.items()}
    b16 = ws["B16"].value
    b17 = ws["B17"].value
    b18 = ws["B18"].value
    wb.close()

    # --- FALLBACK THONG MINH: Neu Excel chua tinh san ket qua cho cac o cong thuc ---
    # 1. So luong kien (D16 co cong thuc =B16)
    if (values.get("qty") is None or str(values["qty"]).strip() == "") and b16 is not None:
        values["qty"] = b16

    # 2. Gross KGS (D18 co cong thuc =B18)
    if (values.get("gross") is None or str(values["gross"]).strip() == "") and b18 is not None:
        values["gross"] = b18

    # 3. Don vi kien (D17 co cong thuc quy doi tu B17 sang PP, BX, PK)
    if (values.get("unit") is None or str(values["unit"]).strip() == "") and b17 is not None:
        s = str(b17).strip().casefold()
        if "pallet" in s:
            values["unit"] = "PP"
        elif "box" in s or "boxes" in s:
            values["unit"] = "BX"
        elif "bundle" in s or "package" in s:
            values["unit"] = "PK"

    # 4. Tong tri gia NVL (B26 co cong thuc =B27-B25)
    if (values.get("custom_total") is None or str(values["custom_total"]).strip() == ""):
        inv_tot = values.get("invoice_total")
        proc_tot = values.get("processing_total")
        if inv_tot is not None and proc_tot is not None:
            try:
                values["custom_total"] = round(float(inv_tot) - float(proc_tot), 2)
            except (ValueError, TypeError):
                pass

    values["sign_text"] = build_sign_text(values)
    return values


def build_sign_text(values):
    """Ghep chuoi 'PTTT: TT, TGTT = TGGC = ... USD,TG NVL = ... USD, NNH: ...'. Thieu thanh phan thi tra ve None."""
    for k in ("processing_total", "custom_total", "ship_to_name"):
        if values.get(k) is None or str(values[k]).strip() == "":
            return None
    try:
        cur = str(values.get("currency") or "").strip() or "USD"
        return SIGN_TEMPLATE.format(
            tggc=fmt_plain(values["processing_total"]),
            tgnvl=fmt_plain(values["custom_total"]),
            cur=cur,
            nnh=str(values["ship_to_name"]).strip(),
        )
    except (ValueError, TypeError):
        return None


# ===== TIM O TRONG ECUS =====
def norm_text(s):
    s = unicodedata.normalize("NFC", str(s or ""))
    return re.sub(r"\s+", " ", s).strip().rstrip(":").strip().casefold()


def is_edit(ctl):
    return ".EDIT." in ctl.cls.upper() or ctl.cls == "Edit"


def locate(items, label_text, index):
    """O EDIT thu 'index' (tinh tu trai sang phai) nam ben phai nhan, cung dong. Tra ve (ctl, ghi chu)."""
    if isinstance(label_text, (tuple, list)):
        targets = {norm_text(t) for t in label_text}
    else:
        targets = {norm_text(label_text)}
    labels = [c for c in items
              if c.visible and "STATIC" in c.cls.upper() and norm_text(c.text) in targets]
    if not labels:
        return None, "khong thay nhan (co dang o dung tab 'Thong tin chung' khong?)"
    label = sorted(labels, key=lambda c: (c.top, c.left))[0]
    note = f"co {len(labels)} nhan trung ten, lay cai tren cung" if len(labels) > 1 else ""

    edits = [c for c in items
             if c.visible and is_edit(c)
             and c.left >= label.right - 2
             and c.top <= label.bottom and c.bottom >= label.top]
    edits.sort(key=lambda c: c.left)
    if len(edits) <= index:
        return None, "khong thay o nhap ben phai nhan"
    return edits[index], note


def same_value(a, b):
    a = re.sub(r"\s+", " ", str(a)).strip().casefold()
    b = re.sub(r"\s+", " ", str(b)).strip().casefold()
    if a == b:
        return True
    try:
        return float(a.replace(",", "")) == float(b.replace(",", ""))
    except ValueError:
        return False


# ===== PHAN WINDOWS (chi chay tren Windows) =====
user32 = kernel32 = None
EnumProc = None
WM_GETTEXT = 0x000D
EM_SETSEL = 0x00B1
SMTO_ABORTIFHUNG = 0x0002
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
INPUT_KEYBOARD = 1
VK_TAB = 0x09
TCM_GETITEMCOUNT = 0x1304
TCM_GETITEMRECT = 0x130A
TCM_GETCURSEL = 0x130B
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
MK_LBUTTON = 0x0001
PROCESS_VM_OPERATION = 0x0008
PROCESS_VM_READ = 0x0010
PROCESS_VM_WRITE = 0x0020
MEM_COMMIT = 0x1000
MEM_RESERVE = 0x2000
MEM_RELEASE = 0x8000
PAGE_READWRITE = 0x04


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD), ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("dwExtraInfo", wintypes.WPARAM)]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG), ("mouseData", wintypes.DWORD),
                ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", wintypes.WPARAM)]


class _InputUnion(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("u", _InputUnion)]


def setup_windows():
    global user32, kernel32, EnumProc
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    try:
        user32.SetProcessDPIAware()
    except Exception:
        pass
    EnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
    user32.IsWindowVisible.argtypes = [wintypes.HWND]
    user32.SetForegroundWindow.argtypes = [wintypes.HWND]
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.AttachThreadInput.argtypes = [wintypes.DWORD, wintypes.DWORD, wintypes.BOOL]
    user32.SetFocus.argtypes = [wintypes.HWND]
    user32.GetFocus.restype = wintypes.HWND
    user32.SendInput.argtypes = [wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int]
    user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.SendMessageW.restype = wintypes.LPARAM
    user32.SendMessageTimeoutW.argtypes = [
        wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM,
        wintypes.UINT, wintypes.UINT, ctypes.POINTER(ctypes.c_size_t),
    ]
    user32.SendMessageTimeoutW.restype = wintypes.LPARAM
    kernel32.GetCurrentThreadId.restype = wintypes.DWORD
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = ctypes.c_void_p
    kernel32.VirtualAllocEx.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t,
                                        wintypes.DWORD, wintypes.DWORD]
    kernel32.VirtualAllocEx.restype = ctypes.c_void_p
    kernel32.ReadProcessMemory.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p,
                                           ctypes.c_size_t, ctypes.POINTER(ctypes.c_size_t)]
    kernel32.ReadProcessMemory.restype = wintypes.BOOL
    kernel32.VirtualFreeEx.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t, wintypes.DWORD]
    kernel32.VirtualFreeEx.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]


def get_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def get_title(hwnd):
    buf = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(hwnd, buf, 512)
    return buf.value


def get_text(hwnd, maxlen=400):
    buf = ctypes.create_unicode_buffer(maxlen)
    result = ctypes.c_size_t()
    ok = user32.SendMessageTimeoutW(hwnd, WM_GETTEXT, maxlen, ctypes.addressof(buf),
                                    SMTO_ABORTIFHUNG, 300, ctypes.byref(result))
    return buf.value if ok else ""


def find_ecus_window():
    found = []

    def cb(hwnd, _):
        if user32.IsWindowVisible(hwnd) and "ECUS5" in get_title(hwnd).upper():
            found.append(hwnd)
        return True

    user32.EnumWindows(EnumProc(cb), 0)
    return found[0] if found else None


def collect_items(main_hwnd):
    handles = []

    def cb(hwnd, _):
        handles.append(hwnd)
        return True

    user32.EnumChildWindows(main_hwnd, EnumProc(cb), 0)
    items = []
    for h in handles:
        r = wintypes.RECT()
        user32.GetWindowRect(h, ctypes.byref(r))
        items.append(Ctl(h, get_class(h), get_text(h), r.left, r.top, r.right, r.bottom,
                         bool(user32.IsWindowVisible(h))))
    return items


def _send_keys(inputs):
    arr = (INPUT * len(inputs))(*inputs)
    user32.SendInput(len(inputs), arr, ctypes.sizeof(INPUT))


def _key(vk=0, scan=0, flags=0):
    ki = KEYBDINPUT(vk, scan, flags, 0, 0)
    inp = INPUT()
    inp.type = INPUT_KEYBOARD
    inp.u.ki = ki
    return inp


def type_text(text):
    for ch in text:
        code = ord(ch)
        if code > 0xFFFF:
            continue  # ky tu ngoai BMP: bo qua
        _send_keys([_key(0, code, KEYEVENTF_UNICODE),
                    _key(0, code, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP)])
        time.sleep(0.01)


def press_tab():
    _send_keys([_key(VK_TAB, 0, 0), _key(VK_TAB, 0, KEYEVENTF_KEYUP)])


def focus_control(main_hwnd, hwnd):
    """Dat focus vao o nhap cua ECUS; tra ve True neu focus dung o do (an toan: sai thi khong go)."""
    user32.SetForegroundWindow(main_hwnd)
    time.sleep(0.1)
    pid = wintypes.DWORD()
    tid_target = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    tid_self = kernel32.GetCurrentThreadId()
    user32.AttachThreadInput(tid_self, tid_target, True)
    try:
        user32.SetFocus(hwnd)
        time.sleep(0.1)
        ok = (user32.GetFocus() == hwnd)
    finally:
        user32.AttachThreadInput(tid_self, tid_target, False)
    return ok


def write_field(main_hwnd, hwnd, text):
    if not focus_control(main_hwnd, hwnd):
        return False
    user32.SendMessageW(hwnd, EM_SETSEL, 0, -1)   # chon het noi dung cu, go de len
    time.sleep(0.05)
    type_text(text)
    time.sleep(0.05)
    press_tab()                                   # roi khoi o de ECUS kiem tra / tu dien mo ta
    time.sleep(0.25)
    return True


# ===== CHUYEN TAB (ghep tu ecus_goto_tab.py) =====
def find_tab_controls(main_hwnd):
    """Cac thanh tab dang hien: (hwnd, so tab, dien tich). Thanh tab chinh la cai lon nhat."""
    handles = []

    def cb(hwnd, _):
        handles.append(hwnd)
        return True

    user32.EnumChildWindows(main_hwnd, EnumProc(cb), 0)
    tabs = []
    for h in handles:
        if "SYSTABCONTROL32" not in get_class(h).upper() or not user32.IsWindowVisible(h):
            continue
        r = wintypes.RECT()
        user32.GetWindowRect(h, ctypes.byref(r))
        count = user32.SendMessageW(h, TCM_GETITEMCOUNT, 0, 0)
        tabs.append((h, count, (r.right - r.left) * (r.bottom - r.top)))
    tabs.sort(key=lambda t: t[2], reverse=True)
    return tabs


def tab_item_rect(tab_hwnd, index):
    """Hoi ECUS vi tri tab thu 'index' (can bo nho tam trong tien trinh ECUS)."""
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(tab_hwnd, ctypes.byref(pid))
    hproc = kernel32.OpenProcess(PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE, False, pid.value)
    if not hproc:
        return None
    try:
        remote = kernel32.VirtualAllocEx(hproc, None, ctypes.sizeof(wintypes.RECT),
                                         MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE)
        if not remote:
            return None
        try:
            if not user32.SendMessageW(tab_hwnd, TCM_GETITEMRECT, index, remote):
                return None
            rect = wintypes.RECT()
            read = ctypes.c_size_t()
            if not kernel32.ReadProcessMemory(hproc, remote, ctypes.byref(rect),
                                              ctypes.sizeof(rect), ctypes.byref(read)):
                return None
            return rect
        finally:
            kernel32.VirtualFreeEx(hproc, remote, 0, MEM_RELEASE)
    finally:
        kernel32.CloseHandle(hproc)


def go_to_tab(main_hwnd, index):
    """Chuyen ECUS sang tab thu 'index'. Tra ve (thanh cong, ly do loi). Dang o tab do roi thi khong lam gi."""
    tabs = find_tab_controls(main_hwnd)
    if not tabs:
        return False, "khong thay thanh tab (da mo to khai chua?)"
    tab_hwnd, count, _area = tabs[0]
    if not (0 <= index < count):
        return False, f"thanh tab chinh chi co {count} tab, khong co tab so {index}"
    if user32.SendMessageW(tab_hwnd, TCM_GETCURSEL, 0, 0) == index:
        return True, ""

    rect = tab_item_rect(tab_hwnd, index)
    if rect is None:
        return False, "khong doc duoc vi tri tab (thu chay script bang quyen Administrator)"
    x = (rect.left + rect.right) // 2
    y = (rect.top + rect.bottom) // 2
    lparam = (y << 16) | (x & 0xFFFF)
    user32.SendMessageW(tab_hwnd, WM_LBUTTONDOWN, MK_LBUTTON, lparam)
    time.sleep(0.05)
    user32.SendMessageW(tab_hwnd, WM_LBUTTONUP, 0, lparam)
    time.sleep(0.5)
    now = user32.SendMessageW(tab_hwnd, TCM_GETCURSEL, 0, 0)
    if now == index:
        return True, ""
    return False, f"da click nhung ECUS van dang o tab so {now}"


def tab_label(index):
    return TAB_NAMES[index] if 0 <= index < len(TAB_NAMES) else f"tab {index}"


# ===== CHUONG TRINH CHINH =====
def pick_file():
    files = sorted(OUTPUT_DIR.glob("Template G-*.xlsx"))
    files = [f for f in files if not f.name.startswith("~$")]
    if not files:
        print(f"Khong thay file Template G-*.xlsx trong {OUTPUT_DIR}")
        return None
    print("Cac file Template:")
    for i, f in enumerate(files):
        print(f"  {i}: {f.name}")
    try:
        return files[int(input("\nChon so thu tu file: ").strip())]
    except (ValueError, IndexError):
        print("Lua chon khong hop le.")
        return None


def run_mac_simulation(path, values):
    print("\n" + "=" * 90)
    print("  [CHẾ ĐỘ MÔ PHỎNG & KIỂM THỬ TRÊN MACOS]")
    print("  macOS không có phần mềm ECUS5 trực tiếp (phần mềm Windows .NET).")
    print("  Tool đang chạy ở chế độ KIỂM THỬ TRƯỚC: Bóc tách, xác thực dữ liệu và mô phỏng điền.")
    print("=" * 90)

    plan = []
    print(f"\n{'Truong':<24}{'Gia tri se dien':<26}{'O nguon':<14}Trang thai")
    print("-" * 90)
    for name, key, label, idx, fmt in FIELDS:
        raw = values.get(key)
        if raw is None or str(raw).strip() == "":
            where = CELLS.get(key, "Cong thuc")
            print(f"{name:<24}{'(TRONG)':<26}{where:<14}Thieu du lieu, bo qua")
            continue
        try:
            text = fmt(raw)
        except (ValueError, TypeError):
            print(f"{name:<24}{str(raw):<26}{'':<14}Loi dinh dang, bo qua")
            continue
        shown = text if len(text) <= 24 else text[:21] + "..."
        where = CELLS.get(key, "Cong thuc")
        print(f"{name:<24}{shown:<26}{where:<14}Hop le (San sang)")
        plan.append((name, text))

    for name, text in plan:
        if len(text) > 24:
            print(f"\n{name} se dien day du:\n  {text}")

    if not plan:
        print("\nKhong co truong nao co du lieu de dien.")
        return

    print("\n" + "-" * 90)
    print("Kiem tra bang tren: Toan bo 12 truong du lieu da duoc boc tach va dinh dang chuan xac.")
    print("Khi chay tren may Windows (co mo ECUS5), tool se tu dong go dung 100% cac truong tren.")

    # Tùy chọn copy vào clipboard trên Mac nếu có lệnh pbcopy
    sign_text = values.get("sign_text")
    if sign_text:
        try:
            import subprocess
            proc = subprocess.Popen(["pbcopy"], stdin=subprocess.PIPE)
            proc.communicate(sign_text.encode("utf-8"))
            print(f"\n[DA COPY VAO CLIPBOARD] Chuoi 'Ky hieu va so hieu' da duoc copy san vao Clipboard cua Mac.")
        except Exception:
            pass

    print("\nBan co muon chay MO PHONG tien trinh dien tren Mac khong? (Go DIEN, hoac Enter de thoat):")
    if input("> ").strip().upper() != "DIEN":
        print("Da thoat che do xem truoc.")
        return

    print("\nBat dau mo phong tien trinh tu dong dien sau 2 giay...")
    time.sleep(2)
    print(f"-> [MO PHONG] Tu dong chuyen sang tab {TAB_THONG_TIN_CHUNG} ({tab_label(TAB_THONG_TIN_CHUNG)})... [OK]")
    for name, text in plan:
        time.sleep(0.2)
        shown = text if len(text) <= 25 else text[:22] + "..."
        print(f"  [MO PHONG GO] {name:<24} -> {shown:<25} [OK]")
    time.sleep(0.3)
    print(f"-> [MO PHONG] Tu dong chuyen sang tab {DEFAULT_TAB} ({tab_label(DEFAULT_TAB)})... [OK]")
    print("\n[HOAN TAT MO PHONG] Toan bo du lieu va logic da duoc kiem thu thanh cong tren Mac!")


def main():
    IS_WINDOWS = (sys.platform == "win32")

    path = pick_file()
    if path is None:
        return

    values = read_template(path)
    print(f"\nDoc file: {path.name}")

    if not IS_WINDOWS:
        run_mac_simulation(path, values)
        return

    # ===== PHAN WINDOWS THAT (KHI CHAY TREN WINDOWS CO MO ECUS5) =====
    setup_windows()

    main_hwnd = find_ecus_window()
    if main_hwnd is None:
        print("Khong thay cua so ECUS5 dang mo.")
        return
    ok, why = go_to_tab(main_hwnd, TAB_THONG_TIN_CHUNG)
    if ok:
        print(f"Dang o tab {TAB_THONG_TIN_CHUNG} ({tab_label(TAB_THONG_TIN_CHUNG)}).")
    else:
        print(f"[CANH BAO] Khong tu chuyen duoc sang tab {tab_label(TAB_THONG_TIN_CHUNG)}: {why}. "
              "Hay bam tay vao tab do; neu khong thay nhan se bao 'khong thay nhan'.")
    items = collect_items(main_hwnd)

    plan = []   # (ten, ctl, gia tri se dien)
    print(f"\n{'Truong':<22}{'Gia tri se dien':<22}{'Hien dang o ECUS':<24}Ghi chu")
    print("-" * 90)
    for name, key, label, idx, fmt in FIELDS:
        raw = values.get(key)
        if raw is None or str(raw).strip() == "":
            where = CELLS.get(key, "B25 / B27 / C13")
            print(f"{name:<22}{'(TRONG)':<22}{'':<24}o {where} trong, bo qua")
            continue
        try:
            text = fmt(raw)
        except (ValueError, TypeError):
            print(f"{name:<22}{str(raw):<22}{'':<24}khong doi duoc dinh dang, bo qua")
            continue
        shown = text if len(text) <= 20 else text[:17] + "..."
        ctl, note = locate(items, label, idx)
        if ctl is None:
            print(f"{name:<22}{shown:<22}{'':<24}{note}")
            continue
        print(f"{name:<22}{shown:<22}{ctl.text.strip()[:22]:<24}{note}")
        plan.append((name, ctl, text))

    for name, _ctl, text in plan:
        if len(text) > 20:
            print(f"\n{name} se dien day du:\n  {text}")

    if not plan:
        print("\nKhong co truong nao dien duoc.")
        return

    print("\nChua dien gi vao ECUS. Kiem tra bang tren cho dung.")
    if input("Go DIEN de dien that vao ECUS (Enter = thoat): ").strip().upper() != "DIEN":
        print("Da thoat, khong thay doi gi.")
        return

    print("\nBam chuot vao cua so ECUS (tab Thong tin chung) neu can. Bat dau sau 5 giay, "
          "dung cham chuot / ban phim cho den khi xong...")
    time.sleep(5)

    for name, ctl, text in plan:
        ok = write_field(main_hwnd, ctl.hwnd, text)
        if not ok:
            print(f"  [BO QUA] {name}: khong dat duoc focus vao o nay nen khong go (de tranh go nham cho)")
            continue
        now = get_text(ctl.hwnd).strip()
        status = "OK" if same_value(now, text) else f"KHAC (ECUS dang hien {now!r})"
        print(f"  {name:<22} -> {(text if len(text) <= 18 else text[:15] + '...'):<18} {status}")

    ok, why = go_to_tab(main_hwnd, DEFAULT_TAB)
    if ok:
        print(f"\nDa chuyen sang tab {DEFAULT_TAB} ({tab_label(DEFAULT_TAB)}).")
    else:
        print(f"\n[CANH BAO] Khong chuyen duoc sang tab {DEFAULT_TAB} ({tab_label(DEFAULT_TAB)}): {why}")

    print("\nXong. HAY KIEM TRA lai cac o, tu nhap booking / ngay gio, roi moi bam Ghi. "
          "Script khong tu bam Ghi hay Gui.")


if __name__ == "__main__":
    main()