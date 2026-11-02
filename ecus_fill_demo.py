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

Nhap hang tu Excel (IMPORT_GOODS): o tab "Danh sach hang", script bam F11 (xoa tat ca hang hien co),
roi F6 (nhap tu file Excel) va chon dung file Template cua invoice dang chon,
vi du D:\\test\\output\\Template G-26-1108.xlsx. Buoc xoa tat ca luon hoi xac nhan truoc.
"""
import ctypes
import json
import re
import sys
import time
import unicodedata
from collections import namedtuple
from ctypes import wintypes
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

# ===== CAU HINH =====
DON_HANG_DIR = Path(r"D:\test\Don hang")   # moi khach 1 thu muc: "Aldila xuat", "Top Glove xuat"... ; file o <khach>\output
SO_FILE_HIEN = 30                            # so file moi nhat hien trong danh sach chon
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

# Nhap hang tu Excel o tab Danh sach hang: F11 (xoa tat ca) -> F6 (nhap tu file Excel) -> chon file Template
# TAM DUNG (30/10): khong nhap dong hang Sheet3. Code nhap hang ben duoi GIU NGUYEN (da test on dinh);
# muon bat lai: doi thanh True, hoac tich o "Nhập dòng hàng (Sheet3)" tren giao dien.
IMPORT_GOODS = False
GOODS_TAB = 2                 # tab Danh sach hang
CONFIRM_BEFORE_DELETE = False  # False = F11 xoa luon (mac dinh). True = hoi (go XOA) truoc khi xoa TOAN BO hang hien co
CONFIRM_EMPTY_BEFORE_IMPORT = False   # False = sau F11 di thang sang F6 (mac dinh). True = dung lai hoi ban kiem tra danh sach da trong chua
TRUST_UNREADABLE_SHEET = True         # O chon sheet cua ECUS la o tu ve, khong doc lai duoc gia tri. True = coi nhu da chon dung sau khi go + Tab
GRID_CLICK = (40, 51)         # diem bam trong luoi hang (tu goc tren-trai luoi): cot STT, dong 1.
                              # ECUS chi nhan F11 / F6 khi con tro dang nam trong luoi hang.

# Form "Nhap danh muc hang tu tap tin Excel" (hien ra sau F6)
EXCEL_SHEET = "Sheet3"        # sheet chua du lieu hang trong file Template
EXCEL_FIRST_DATA_ROW = 2      # "Hang dau chua du lieu" (dong 1 la tieu de)
AUTO_CLICK_GHI = True         # True = tu bam Ghi roi bam OK o hop thoai ket qua. False = dung lai de ban tu bam Ghi.

# O trong file Template (Sheet1) - BO CUC MOI 02/11/2026 (xem khach_hang\Template_chuan.xlsx)
CELLS = {
    "invoice_no": "B24",     # So invoice
    "invoice_date": "B25",   # Ngay invoice (= Ship Date)
    "currency": "C16",       # Currency
    "terms": "C15",          # Terms of sale (vi du "FOB Ho Chi Minh City" -> lay "FOB")
    "invoice_total": "B30",  # Tong tri gia hoa don
    "qty": "D19",            # So luong kien
    "unit": "D20",           # Don vi (PP, BX...)
    "kind": "B20",           # Loai kien dang chu tu invoice (BOX, PLASTIC PALLET, BUNDLE...)
    "gross": "D21",          # KGS
    "processing_total": "B28",   # Tong tri gia Gia cong -> TGGC
    "custom_total": "B29",       # Tong tri gia Nguyen vat lieu -> TG NVL
    "ship_to_name": "C17",       # Ship to Name -> NNH
}

# ===== O BAT BUOC: thieu / sai thi KHONG chay ECUS =====
# Ma dia diem: (khoa, o ma, cot ma Sheet4, cot ten Sheet4, o ten (go tay khi ma ZZZ), ten hien thi)
MA_DIA_DIEM = [
    ("dd_cuoi", "C10", "B", "C", "D10", "Ma nhan hang cuoi cung"),
    ("dd_xep",  "C11", "G", "H", "D11", "Ma xep hang"),
    ("dd_bt",   "C32", "L", "M", "D32", "Ma van chuyen bao thue"),
]
TEN_CAN_GO_TAY = {"OTHER", "UNKNOWN"}    # ma ZZZ mang ten nay -> phai go ten dia diem vao cot D
# O main.py tu dien tu hoa don (o nhap, khong phai cong thuc)
O_BAT_BUOC = [("B19", "So luong kien"), ("B20", "Loai kien"), ("B21", "Gross"), ("B24", "So invoice"),
              ("B25", "Ngay invoice"), ("B30", "Tong tri gia hoa don"), ("C16", "Currency"),
              ("C17", "Ship to Name")]
DANH_MUC = "Sheet4"
FILE_NHO = "nho_ma.json"   # luu ma da khai vao thu muc khach; main.py dung de dien san lan sau

# Phuong thuc / phuong tien van chuyen (nhap tay tu booking)
O_PT_CHON = "B3"                                     # "Như cũ" = giu nguyen tren ECUS (copy lo cu) | "Sửa" = dien lai
O_PT_MA, O_PT_HIEU, O_PT_TEN = "C3", "C12", "D12"    # ma phuong thuc | o 1 ho hieu | o 2 ten tau / chuyen
PT_HOP_LE = {"1", "2", "3", "4", "5", "6", "9"}
PT_BIEN_SONG = {"2", "3", "6"}          # o 1 (ho hieu) = 9999; phuong thuc khac: o 1 de trong
PT_CAN_TEN = {"1", "2", "3", "6"}       # bat buoc ten tau / so chuyen
XOA_KHI_TRONG = {"pt_hieu", "pt_ten"}   # o trong trong Template -> xoa trang o do tren ECUS (khong de gia tri cu)

# Booking nhap tren giao dien (ecus_auto.py): luu o <thu muc khach>\booking.json theo ten file Template.
# Lo nao co booking -> dung booking (bo qua B3/C3/D12 trong Excel) + dien them "Ngay hang di du kien".
FILE_BOOKING = "booking.json"


def doc_booking(thu_muc_khach):
    f = Path(thu_muc_khach) / FILE_BOOKING
    try:
        return json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    except Exception:
        return {}


KHOA_BOOKING = ("pt_ma", "pt_ten", "ngay_di")
KHOA_DIA_DIEM = tuple(k for m in MA_DIA_DIEM for k in (m[0], m[0] + "_ten"))   # dd_cuoi, dd_cuoi_ten, ...


def ghi_booking(thu_muc_khach, ten_files, cap_nhat=None, xoa_khoa=()):
    """Gop thong tin nhap tren giao dien cho tung lo.
    cap_nhat: {"pt_ma": "2", "pt_ten": "CV7856/25OCT", "ngay_di": "05/11/2026", "dd_cuoi": "...", ...}
    xoa_khoa: cac khoa can bo (vi du KHOA_BOOKING) -> lo do quay ve lay theo Excel."""
    bk = doc_booking(thu_muc_khach)
    for t in ten_files:
        muc = bk.get(t, {})
        for k in xoa_khoa:
            muc.pop(k, None)
        muc.update(cap_nhat or {})
        if muc:
            bk[t] = muc
        else:
            bk.pop(t, None)
    (Path(thu_muc_khach) / FILE_BOOKING).write_text(json.dumps(bk, ensure_ascii=False, indent=2), encoding="utf-8")


def xoa_khoi_so(thu_muc_khach, ten_files):
    """Xoa Template -> bo luon booking va lich su da khai cua lo do."""
    for ten_file, ham_doc in ((FILE_BOOKING, doc_booking), (FILE_LICH_SU, doc_lich_su)):
        d = ham_doc(thu_muc_khach)
        if any(t in d for t in ten_files):
            for t in ten_files:
                d.pop(t, None)
            (Path(thu_muc_khach) / ten_file).write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")


def kiem_tra_booking(b):
    """Tra ve danh sach loi cua 1 booking (rong = hop le). Chi kiem phan phuong thuc khi co pt_ma."""
    loi = []
    pt = str(b.get("pt_ma", "")).strip()
    if "pt_ma" not in b:
        pass
    elif pt not in PT_HOP_LE:
        loi.append("Booking: phuong thuc van chuyen phai la 1-6 hoac 9")
    elif pt in PT_CAN_TEN and not str(b.get("pt_ten", "")).strip():
        loi.append("Booking: chua nhap ten tau / so chuyen")
    nd = str(b.get("ngay_di", "")).strip()
    if nd:
        try:
            datetime.strptime(nd, "%d/%m/%Y")
        except ValueError:
            loi.append(f"Booking: ngay di '{nd}' sai dang dd/mm/yyyy")
    return loi

# Chuoi dien vao o "Ky hieu va so hieu" (sua lai cho dung y neu can, giu nguyen khoang trang)
NOTE_LABEL = "Ký hiệu và số hiệu"   # nhan o cot A cua Sheet1; gia tri o cot B (neu Template co dong nay)
SIGN_TEMPLATE = "PTTT: TT, TGTT = TGGC = {tggc} {cur}, TG NVL = {tgnvl} {cur}, NNH: {nnh}"
# Ghi them vao cuoi chuoi tren theo loai kien (o B20). Chu so sanh khong phan biet hoa thuong.
GHI_CHU_THEO_LOAI_KIEN = {
    "bundle": ", 1 PK = 1 BUNDLE",
}


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
    # Phuong thuc / phuong tien van chuyen (C3, C12, D12)
    ("Ma phuong thuc VC",    "pt_ma",         "Mã hiệu phương thức vận chuyển:", 0, fmt_text),
    ("Ho hieu (o 1)",        "pt_hieu",       "Phương tiện vận chuyển:",        0, fmt_text),
    ("Ten tau/chuyen (o 2)", "pt_ten",        "Phương tiện vận chuyển:",        1, fmt_text),
    ("Ngay hang di du kien", "ngay_di",       "Ngày hàng đi dự kiến:",          0, fmt_date),   # chi khi co booking
    # Ma dia diem (C10 / C11 / C32). Ten (o thu 2) chi go khi ma ZZZ, lay tu cot D cua Template
    ("DD nhan cuoi cung",    "dd_cuoi",       "Địa điểm nhận hàng cuối cùng:",  0, fmt_text),
    ("Ten DD nhan cuoi",     "dd_cuoi_ten",   "Địa điểm nhận hàng cuối cùng:",  1, fmt_text),
    ("DD xep hang",          "dd_xep",        "Địa điểm xếp hàng:",             0, fmt_text),
    ("Ten DD xep hang",      "dd_xep_ten",    "Địa điểm xếp hàng:",             1, fmt_text),
    ("DD dich bao thue",     "dd_bt",         "Địa điểm đích cho vận chuyển bảo thuế:", 0, fmt_text),
    ("Ten DD bao thue",      "dd_bt_ten",     "Địa điểm đích cho vận chuyển bảo thuế:", 1, fmt_text),
    # Hai ngay duoi day = ngay invoice. ECUS viet nhan la 'khời' (sai chinh ta), giu ca 2 cach viet
    ("Ngay khoi hanh VC",    "invoice_date",  ("Ngày khời hành vận chuyển:", "Ngày khởi hành vận chuyển:"), 0, fmt_date),
    # Dong 'Dia diem dich...' co 3 o nhap: ma (06B1A04), ten (KHO SCSC), ngay -> ngay la o thu 3 (chi so 2)
    ("Ngay den DD bao thue", "invoice_date",  "Địa điểm đích cho vận chuyển bảo thuế:", 2, fmt_date),
]

Ctl = namedtuple("Ctl", "hwnd cls text left top right bottom visible")


# ===== KIEM TRA O BAT BUOC =====
def _ma_don_vi(loai):
    """Giong cong thuc D20 cua Template: chu loai kien -> ma ECUS."""
    t = str(loai or "").casefold()
    if "pallet" in t:
        return "PP"
    if "box" in t:
        return "BX"
    if "bundle" in t or "package" in t:
        return "PK"
    return ""


def _doc_danh_muc_tu(wb):
    ten_sheet = {n.strip().casefold(): n for n in wb.sheetnames}
    danh_muc = {m[0]: {} for m in MA_DIA_DIEM}
    if DANH_MUC.casefold() in ten_sheet:
        vt = {m[0]: (column_index_from_string(m[2]) - 1, column_index_from_string(m[3]) - 1) for m in MA_DIA_DIEM}
        for hang in wb[ten_sheet[DANH_MUC.casefold()]].iter_rows(values_only=True):
            for khoa, (cm, cn) in vt.items():
                if cm < len(hang) and hang[cm] not in (None, ""):
                    danh_muc[khoa][str(hang[cm]).strip().upper()] = str(hang[cn] if cn < len(hang) else "").strip()
    return danh_muc


def doc_danh_muc(path):
    """Doc danh muc ma (Sheet4) 1 lan, vi du tu Template_<khach>.xlsx goc, de kiem tra nhieu file cho nhanh."""
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        return _doc_danh_muc_tu(wb)
    finally:
        wb.close()


def kiem_tra_template(path, danh_muc=None):
    """Tra ve (danh sach thieu, gia tri dia diem). Danh sach rong = du, duoc chay ECUS.
    danh_muc: truyen san (tu doc_danh_muc) de khoi doc lai Sheet4 cua tung file."""
    wb = load_workbook(path, data_only=True, read_only=True)
    ten_sheet = {n.strip().casefold(): n for n in wb.sheetnames}
    ws = wb[ten_sheet[SHEET.casefold()]]
    o = {a: ws[a].value for a in [x for x, _ in O_BAT_BUOC] + [m[1] for m in MA_DIA_DIEM] + [m[4] for m in MA_DIA_DIEM]
         + [O_PT_CHON, O_PT_MA, O_PT_HIEU, O_PT_TEN, "A3"]}

    # danh muc ma o Sheet4
    if danh_muc is None:
        danh_muc = _doc_danh_muc_tu(wb)
    wb.close()

    thieu, gia_tri = [], {}
    if norm_text(o["A3"]) != norm_text("Phương thức vận chuyển"):     # Template tao tu bo cuc cu
        return ["Bo cuc cu: file nay tao tu Template cu (truoc 02/11), hay Tao Template lai"], {}
    booking = doc_booking(Path(path).parent.parent).get(Path(path).name) or {}
    for khoa, o_ma, _cm, _cn, o_ten, ten in MA_DIA_DIEM:
        ma_tho, ten_tho = o[o_ma], o[o_ten]
        if booking.get(khoa):          # ma dia diem nhap tren giao dien -> uu tien hon Excel
            ma_tho, ten_tho = booking[khoa], booking.get(khoa + "_ten", "")
            o_ma, o_ten = f"{o_ma} (giao dien)", "o ten (giao dien)"
        ma = str(ma_tho or "").strip().upper()
        if not ma:
            thieu.append(f"{o_ma} {ten}: chua dien")
            continue
        if danh_muc[khoa] and ma not in danh_muc[khoa]:
            thieu.append(f"{o_ma} {ten}: ma '{ma}' khong co trong {DANH_MUC}")
            continue
        gia_tri[khoa] = ma
        if danh_muc[khoa].get(ma, "").upper() in TEN_CAN_GO_TAY:
            go = str(ten_tho or "").strip()
            if not go or norm_text(go) in ("điền vô", "other", "unknown"):
                thieu.append(f"{o_ten} {ten}: ma {ma} la ZZZ, phai go ten dia diem vao {o_ten}")
            else:
                gia_tri[khoa + "_ten"] = go
    for a, ten in O_BAT_BUOC:
        if o[a] is None or str(o[a]).strip() == "":
            thieu.append(f"{a} {ten}: trong")
    if o.get("B20") not in (None, "") and not _ma_don_vi(o["B20"]):
        thieu.append(f"B20 Loai kien '{o['B20']}': khong doi duoc ra ma ECUS (PP / BX / PK)")

    # --- Phuong thuc / phuong tien van chuyen ---
    def chu(v):
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        return "" if v is None else str(v).strip()
    pt, hieu, ten_pt = chu(o[O_PT_MA]), chu(o[O_PT_HIEU]), chu(o[O_PT_TEN])
    if booking.get("pt_ma"):           # booking nhap tren giao dien -> uu tien hon Excel
        thieu += kiem_tra_booking(booking)
        pt = str(booking.get("pt_ma", "")).strip()
        gia_tri["pt_ma"] = pt
        gia_tri["pt_hieu"] = "9999" if pt in PT_BIEN_SONG else ""
        gia_tri["pt_ten"] = str(booking.get("pt_ten", "")).strip()
        if str(booking.get("ngay_di", "")).strip():
            gia_tri["ngay_di"] = str(booking["ngay_di"]).strip()
    elif norm_text(o[O_PT_CHON]) != norm_text("Sửa"):
        gia_tri["pt_giu"] = True       # "Nhu cu": khong kiem tra, khong dung vao ECUS
    elif pt not in PT_HOP_LE:
        thieu.append(f"{O_PT_MA} Ma phuong thuc van chuyen: " + (f"'{pt}' khong hop le" if pt else "chua dien")
                     + " (1-6 hoac 9)")
    else:
        gia_tri["pt_ma"] = pt
        if pt in PT_CAN_TEN and not ten_pt:
            thieu.append(f"{O_PT_TEN} Ten tau / so chuyen: chua dien")
        gia_tri["pt_hieu"] = "9999" if pt in PT_BIEN_SONG else ""   # tu suy ra tu C3; "" -> xoa o 1 tren ECUS
        gia_tri["pt_ten"] = ten_pt     # "" -> xoa o 2 tren ECUS
    return thieu, gia_tri


FILE_LICH_SU = "da_khai.json"   # trong thu muc khach: file Template nao da dien vao ECUS, luc nao


def doc_lich_su(thu_muc_khach):
    f = Path(thu_muc_khach) / FILE_LICH_SU
    try:
        return json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    except Exception:
        return {}


def ghi_lich_su(path):
    path = Path(path)
    thu_muc = path.parent.parent
    ls = doc_lich_su(thu_muc)
    ls[path.name] = datetime.now().strftime("%d/%m/%Y %H:%M")
    try:
        (thu_muc / FILE_LICH_SU).write_text(json.dumps(ls, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        print(f"[CANH BAO] Khong ghi duoc lich su: {e}")


def luu_nho(path, values):
    """Ghi lai ma dia diem vua khai vao <thu muc khach>\\nho_ma.json de main.py dien san lan sau."""
    f = Path(path).parent.parent / FILE_NHO
    try:
        nho = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    except Exception:
        nho = {}
    nho.setdefault("gan_nhat", {})
    nho.setdefault("theo_ship_to", {})
    ship = " ".join(str(values.get("ship_to_name") or "").split()).casefold()
    for khoa, o_ma, _cm, _cn, _o_ten, _ten in MA_DIA_DIEM:
        if not values.get(khoa):
            continue
        muc = {"ma": values[khoa], "ten": values.get(khoa + "_ten")}
        nho["gan_nhat"][o_ma] = muc
        if o_ma == "C10" and ship:
            nho["theo_ship_to"][ship] = muc
    try:
        f.write_text(json.dumps(nho, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\nDa nho ma dia diem vao {f.name} (lan sau main.py dien san).")
    except Exception as e:
        print(f"\n[CANH BAO] Khong luu duoc {f.name}: {e}")


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
    # Uu tien chuoi 'Ky hieu va so hieu' da ghi san trong Template (moi khach co mau rieng, do
    # khach_hang\chay.py ghi vao dong co nhan nay o cot A). Khong co thi ghep theo mau Aldila nhu cu.
    san = None
    for r in range(1, ws.max_row + 1):
        if norm_text(ws.cell(r, 1).value) == norm_text(NOTE_LABEL):
            san = ws.cell(r, 2).value
            break
    wb.close()
    values["sign_text"] = str(san).strip() if san not in (None, "") else build_sign_text(values)
    return values


def build_sign_text(values):
    """Ghep chuoi 'PTTT: TT, TGTT = TGGC = ... USD,TG NVL = ... USD, NNH: ...'. Thieu thanh phan thi tra ve None."""
    for k in ("processing_total", "custom_total", "ship_to_name"):
        if values.get(k) is None or str(values[k]).strip() == "":
            return None
    try:
        cur = str(values.get("currency") or "").strip() or "USD"
        text = SIGN_TEMPLATE.format(
            tggc=fmt_money(values["processing_total"]),     # luon 2 so le: 749.00, 2,594.82
            tgnvl=fmt_money(values["custom_total"]),
            cur=cur,
            nnh=str(values["ship_to_name"]).strip(),
        )
        loai = norm_text(values.get("kind"))
        for chu, them in GHI_CHU_THEO_LOAI_KIEN.items():
            if chu in loai:
                text += them
        return text
    except (ValueError, TypeError):
        return None


# ===== TIM O TRONG ECUS =====
def norm_text(s):
    s = unicodedata.normalize("NFC", str(s or ""))
    return re.sub(r"\s+", " ", s).strip().rstrip(":").strip().casefold()


def is_edit(ctl):
    return ".EDIT." in ctl.cls.upper() or ctl.cls == "Edit"


def is_kind(ctl, kind):
    if kind == "combo":
        return "COMBOBOX" in ctl.cls.upper()
    if kind == "any":
        cls = ctl.cls.upper()
        return "STATIC" not in cls and "BUTTON" not in cls
    return is_edit(ctl)


def locate(items, label_text, index, kind="edit"):
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
             if c.visible and is_kind(c, kind)
             and c.left >= label.right - 2
             and c.top <= label.bottom and c.bottom >= label.top]
    edits.sort(key=lambda c: (c.left, -(c.right - c.left) * (c.bottom - c.top)))
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
VK_DELETE = 0x2E
VK_RETURN = 0x0D
VK_F6 = 0x75
VK_F11 = 0x7A
BM_CLICK = 0x00F5
CB_SETEDITSEL = 0x0142
CB_SETCURSEL = 0x014E
WM_COMMAND = 0x0111
CBN_SELCHANGE = 1
CBN_SELENDOK = 9
KEYEVENTF_EXTENDEDKEY = 0x0001
VK_UP = 0x26
VK_DOWN = 0x28
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
    user32.GetForegroundWindow.restype = ctypes.c_void_p
    user32.GetParent.argtypes = [wintypes.HWND]
    user32.GetParent.restype = ctypes.c_void_p
    user32.GetDlgCtrlID.argtypes = [wintypes.HWND]
    user32.GetDlgCtrlID.restype = ctypes.c_int
    user32.IsChild.argtypes = [wintypes.HWND, wintypes.HWND]
    user32.IsChild.restype = wintypes.BOOL
    user32.WindowFromPoint.argtypes = [wintypes.POINT]
    user32.WindowFromPoint.restype = ctypes.c_void_p
    user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.PostMessageW.restype = wintypes.BOOL
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
        f = user32.GetFocus()
        ok = (f == hwnd) or bool(f and user32.IsChild(hwnd, f))
    finally:
        user32.AttachThreadInput(tid_self, tid_target, False)
    return ok


def write_field(main_hwnd, hwnd, text):
    if not focus_control(main_hwnd, hwnd):
        return False
    user32.SendMessageW(hwnd, EM_SETSEL, 0, -1)   # chon het noi dung cu, go de len
    time.sleep(0.05)
    if text == "":
        press_vk(VK_DELETE, extended=True)          # chuoi rong = xoa trang o
    else:
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


# ===== F11 (XOA TAT CA) + F6 (NHAP TU EXCEL) =====
POSITIVE_BUTTONS = {"yes", "có", "co", "ok", "đồng ý", "dong y", "xóa", "xoa", "xác nhận", "xac nhan"}


def pick_positive(buttons):
    """Chon nut dong y trong hop thoai xac nhan. buttons: [(hwnd, text)]. Khong co thi tra ve None."""
    for hwnd, text in buttons:
        if norm_text(str(text).replace("&", "")) in POSITIVE_BUTTONS:
            return hwnd
    return None


def press_vk(vk, extended=False):
    flags = KEYEVENTF_EXTENDEDKEY if extended else 0
    _send_keys([_key(vk, 0, flags), _key(vk, 0, flags | KEYEVENTF_KEYUP)])


def process_id(hwnd):
    pid = wintypes.DWORD()
    if hwnd:
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


def process_windows(pid):
    """Tap cac cua so cap cao dang hien cua tien trinh ECUS."""
    found = set()

    def cb(hwnd, _):
        if user32.IsWindowVisible(hwnd) and process_id(hwnd) == pid:
            found.add(hwnd)
        return True

    user32.EnumWindows(EnumProc(cb), 0)
    return found


def wait_new_window(pid, before, timeout=6.0):
    """Cho cua so moi (hop thoai) cua ECUS xuat hien. Tra ve hwnd hoac None."""
    end = time.time() + timeout
    while time.time() < end:
        new = process_windows(pid) - before
        if new:
            time.sleep(0.3)
            # uu tien hop thoai chuan (#32770)
            return sorted(new, key=lambda h: get_class(h) != "#32770")[0]
        time.sleep(0.2)
    return None


def dialog_buttons(dlg):
    handles = []

    def cb(hwnd, _):
        handles.append(hwnd)
        return True

    user32.EnumChildWindows(dlg, EnumProc(cb), 0)
    return [(h, get_text(h)) for h in handles if "BUTTON" in get_class(h).upper()]


def window_rect(hwnd):
    r = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(r))
    return r.left, r.top, r.right, r.bottom


def find_goods_grid(tab_hwnd):
    """Tim luoi hang trong tab hien tai: do cac diem phia duoi thanh cong cu, lay cua so lon, nho nhat trong so do.
    Tra ve (hwnd, rect) hoac None."""
    tl, tt, tr, _tb = window_rect(tab_hwnd)
    tab_w = tr - tl
    cands = {}
    for dy in range(90, 500, 25):
        for dx in (60, 300, tab_w // 2):
            h = user32.WindowFromPoint(wintypes.POINT(tl + dx, tt + dy))
            if not h or h in cands:
                continue
            if "WINDOW.8" not in get_class(h).upper():
                continue
            l, t, r, b = window_rect(h)
            w, ht = r - l, b - t
            if w >= 0.7 * tab_w and ht >= 300 and t >= tt + 60:   # du lon, nam duoi thanh cong cu
                cands[h] = (w * ht, (l, t, r, b))
    if not cands:
        return None
    h = min(cands, key=lambda k: cands[k][0])
    return h, cands[h][1]


def focus_goods_grid(main_hwnd, tab_hwnd, pid):
    """Dua ECUS len truoc va bam vao 1 o trong luoi hang (ECUS chi nhan F11 / F6 khi con tro o trong luoi)."""
    user32.SetForegroundWindow(main_hwnd)
    time.sleep(0.2)
    found = find_goods_grid(tab_hwnd)
    if found:
        grid, rect = found
        print(f"  Luoi hang: hwnd={grid} rect={rect}")
        x, y = GRID_CLICK
        lparam = (y << 16) | (x & 0xFFFF)
        user32.SendMessageW(grid, WM_LBUTTONDOWN, MK_LBUTTON, lparam)
        time.sleep(0.05)
        user32.SendMessageW(grid, WM_LBUTTONUP, 0, lparam)
        time.sleep(0.2)
        focus_control(main_hwnd, grid)
    else:
        print("  [CANH BAO] Khong tim thay luoi hang tu dong, chi dat focus vao thanh tab.")
        focus_control(main_hwnd, tab_hwnd)
    time.sleep(0.2)
    return process_id(user32.GetForegroundWindow()) == pid


def press_and_wait(vk, pid, timeout):
    before = process_windows(pid)
    press_vk(vk)
    return wait_new_window(pid, before, timeout)


def handle_confirm_dialog(dlg):
    buttons = dialog_buttons(dlg)
    print(f"  Hop thoai: {get_title(dlg)!r}  cac nut: {[t for _h, t in buttons]}")
    yes = pick_positive(buttons)
    if yes is not None:
        user32.PostMessageW(yes, BM_CLICK, 0, 0)
        time.sleep(1.0)
        print("  Da bam nut dong y.")
    else:
        input("  Khong nhan ra nut dong y. Hay bam tay nut dong y tren ECUS roi bam Enter o day: ")


def sheet_names(excel_path):
    """Ten cac sheet cua file Excel, theo thu tu (de biet Sheet3 la muc thu may trong o chon sheet)."""
    try:
        wb = load_workbook(excel_path, read_only=True)
        names = list(wb.sheetnames)
        wb.close()
        return names
    except Exception:
        return []


def combo_text(form, combo):
    """Gia tri dang hien o ComboBox. Neu la o tu ve (khong phai COMBOBOX chuan) va khong doc truc tiep duoc
    thi doc tu cac cua so con nam ben trong no. Khong doc duoc thi tra ve chuoi rong."""
    txt = get_text(combo.hwnd).strip()
    if txt:
        return txt
    l, t, r, b = window_rect(combo.hwnd)
    for c in collect_items(form):
        if c.hwnd == combo.hwnd or "STATIC" in c.cls.upper() or not c.visible:
            continue
        if c.text.strip() and c.left >= l - 2 and c.right <= r + 2 and c.top >= t - 2 and c.bottom <= b + 2:
            return c.text.strip()
    return ""


def select_combo_item(form, combo, want, names):
    """Chon muc 'want' trong o chon (combo: Ctl). Tra ve (cach, da_xac_nhan).
    - ComboBox chuan: thu phim mui ten -> CB_SETCURSEL -> go chu.
    - O tu ve (nhu o Sheet cua ECUS): go chu + Tab; neu khong doc lai duoc gia tri thi tra ve (cach, False).
    cach = None nghia la that bai."""
    hwnd = combo.hwnd
    native = "COMBOBOX" in combo.cls.upper()
    folded = [norm_text(n) for n in names]
    want_idx = folded.index(norm_text(want)) if norm_text(want) in folded else None

    def matches():
        return norm_text(combo_text(form, combo)) == norm_text(want)

    def by_keys():
        if want_idx is None or not focus_control(form, hwnd):
            return
        last = None
        for _ in range(8):
            cur = norm_text(combo_text(form, combo))
            if cur == norm_text(want) or cur == last:
                break   # da toi noi, hoac bam phim ma o khong doi (de khong bam lung tung)
            last = cur
            if cur in folded:
                delta = want_idx - folded.index(cur)
                key = VK_DOWN if delta > 0 else VK_UP
                for _i in range(abs(delta)):
                    press_vk(key, extended=True)
                    time.sleep(0.15)
            else:
                press_vk(VK_DOWN, extended=True)   # chua chon muc nao: bam xuong de chon muc dau
                time.sleep(0.15)
            time.sleep(0.25)

    def by_message():
        if want_idx is None:
            return
        user32.SendMessageW(hwnd, CB_SETCURSEL, want_idx, 0)
        parent = user32.GetParent(hwnd)
        cid = user32.GetDlgCtrlID(hwnd)
        if parent:
            for note in (CBN_SELENDOK, CBN_SELCHANGE):
                user32.SendMessageW(parent, WM_COMMAND, (note << 16) | (cid & 0xFFFF), hwnd)
        time.sleep(0.3)

    def by_typing():
        if focus_control(form, hwnd):
            user32.SendMessageW(hwnd, CB_SETEDITSEL, 0, 0xFFFF0000)
            time.sleep(0.05)
            type_text(want)
            time.sleep(0.1)
            press_tab()
            time.sleep(0.4)

    if matches():
        return "da dung san", True

    if native:
        steps = [("phim mui ten", by_keys), ("CB_SETCURSEL", by_message), ("go chu", by_typing)]
    else:
        steps = [("go chu", by_typing), ("phim mui ten", by_keys)]

    for label, fn in steps:
        fn()
        if matches():
            return label, True
        if not native and combo_text(form, combo) == "":
            return f"{label} (khong doc lai duoc)", False   # o tu ve: dung lai, khong thu tiep de khoi doi gia tri
    return None, False


def base_name(path_text):
    """Ten file (khong co thu muc) tu duong dan dang chu, dung ca khi la \\ hay /."""
    return re.split(r"[\\/]", str(path_text))[-1]


OK_BUTTONS = {"ok", "đồng ý", "dong y"}


def dialog_messages(dlg):
    """Cac dong chu (noi dung thong bao) trong mot hop thoai."""
    handles = []

    def cb(hwnd, _):
        handles.append(hwnd)
        return True

    user32.EnumChildWindows(dlg, EnumProc(cb), 0)
    msgs = []
    for h in handles:
        if "STATIC" in get_class(h).upper():
            t = get_text(h).strip()
            if t:
                msgs.append(t)
    return msgs


def click_ok_dialogs(pid, before, first_timeout=6.0, next_timeout=2.5, max_dialogs=4):
    """Sau khi bam Ghi: doi cac hop thoai ket qua cua ECUS, in noi dung roi bam OK. Tra ve so hop thoai da xu ly."""
    handled = 0
    timeout = first_timeout
    while handled < max_dialogs:
        dlg = wait_new_window(pid, before, timeout)
        if dlg is None:
            break
        buttons = dialog_buttons(dlg)
        print(f"  Hop thoai: {get_title(dlg)!r}  noi dung: {dialog_messages(dlg)}  "
              f"cac nut: {[t for _h, t in buttons]}")
        ok = None
        for h, t in buttons:
            if norm_text(str(t).replace("&", "")) in OK_BUTTONS:
                ok = h
        if ok is not None:
            user32.PostMessageW(ok, BM_CLICK, 0, 0)
            time.sleep(1.0)
            print("  Da bam OK.")
        else:
            input("  Hop thoai nay khong co nut OK nen script khong tu bam. Hay xu ly bang tay roi bam Enter: ")
        handled += 1
        before = process_windows(pid)
        timeout = next_timeout
    return handled


def import_excel_form(form, pid, excel_path):
    """Form 'Nhap danh muc hang tu tap tin Excel': chon file -> chon sheet -> hang dau. Khong bam Ghi (tru khi AUTO_CLICK_GHI)."""
    title = norm_text(get_title(form))
    if "nhập danh mục hàng từ tập tin excel" not in title:
        print(f"  [DUNG] Cua so mo ra khong phai form nhap Excel (title={get_title(form)!r}), khong thao tac tiep.")
        return False
    print(f"  Form nhap Excel: {get_title(form)!r}")
    user32.SetForegroundWindow(form)
    time.sleep(0.4)

    # --- 1. Chon tap tin ---
    buttons = dialog_buttons(form)
    pick = None
    for h, t in buttons:
        if norm_text(str(t).replace("&", "")).startswith("chọn tập tin"):
            pick = h
    if pick is None:
        print(f"  [DUNG] Khong thay nut 'Chon tap tin...'. Cac nut: {[t for _h, t in buttons]}")
        return False
    before = process_windows(pid)
    user32.PostMessageW(pick, BM_CLICK, 0, 0)
    fdlg = wait_new_window(pid, before, 5.0)
    if fdlg is None:
        print("  [DUNG] Bam 'Chon tap tin...' ma khong thay hop thoai chon file.")
        return False
    print(f"  Hop thoai chon file: title={get_title(fdlg)!r}  class={get_class(fdlg)}")
    if get_class(fdlg) != "#32770":
        print("  [DUNG] Khong phai hop thoai chon file chuan nen khong go duong dan, de tranh go nham.")
        return False
    user32.SetForegroundWindow(fdlg)
    time.sleep(0.5)
    if user32.GetForegroundWindow() != fdlg:
        print("  [DUNG] Khong dua duoc hop thoai len truoc, khong go duong dan de tranh go nham.")
        return False
    type_text(str(excel_path))
    time.sleep(0.3)
    press_vk(VK_RETURN)
    end = time.time() + 8
    while time.time() < end and user32.IsWindowVisible(fdlg):
        time.sleep(0.2)
    time.sleep(0.8)

    items = collect_items(form)
    verified = True

    # Kiem tra ECUS da nhan dung file chua
    file_ctl, _ = locate(items, "Tập tin Excel:", 0)
    shown = get_text(file_ctl.hwnd).strip() if file_ctl else ""
    want = base_name(excel_path)
    got = base_name(shown) if shown else ""
    if got.casefold() == want.casefold():
        print(f"  Tap tin: {shown}  -> OK")
    else:
        print(f"  [CANH BAO] O 'Tap tin Excel' dang hien {shown!r}, khac file mong muon {want!r}")
        verified = False

    # --- 2. Chon sheet ---
    combo, _ = locate(items, "Sheet chứa dữ liệu:", 0, kind="combo")
    if combo is None:
        combo, _ = locate(items, "Sheet chứa dữ liệu:", 0, kind="any")
    if combo is None:
        print("  [CANH BAO] Khong thay o 'Sheet chua du lieu'. Hay chon sheet bang tay.")
        verified = False
    else:
        names = sheet_names(excel_path)
        print(f"  O sheet: class={combo.cls}  dang hien={combo_text(form, combo)!r}  "
              f"cac sheet trong file: {names}")
        how, confirmed = select_combo_item(form, combo, EXCEL_SHEET, names)
        if how is None:
            input(f"  [CANH BAO] Chua chon duoc {EXCEL_SHEET}. Hay chon bang tay tren ECUS roi bam Enter: ")
            now = combo_text(form, combo)
            if norm_text(now) == norm_text(EXCEL_SHEET):
                print(f"  Sheet: {now}  -> OK")
            elif now == "" and TRUST_UNREADABLE_SHEET:
                print("  Khong doc lai duoc o Sheet, tin vao viec ban da chon.")
            else:
                print(f"  [CANH BAO] O 'Sheet chua du lieu' dang hien {now!r}, mong muon {EXCEL_SHEET!r}.")
                verified = False
        elif confirmed:
            print(f"  Sheet: {EXCEL_SHEET}  -> OK (bang {how})")
        else:
            print(f"  Sheet: da chon {EXCEL_SHEET} ({how}). ECUS khong cho doc lai gia tri o nay, "
                  f"hay nhin man hinh xem o Sheet co hien {EXCEL_SHEET} khong.")
            if not TRUST_UNREADABLE_SHEET:
                verified = False

    # --- 3. Hang dau chua du lieu ---
    row_ctl, _ = locate(items, "Hàng đầu chứa dữ liệu:", 0)
    if row_ctl is None:
        print("  [CANH BAO] Khong thay o 'Hang dau chua du lieu'. Hay nhap bang tay.")
        verified = False
    else:
        write_field(form, row_ctl.hwnd, str(EXCEL_FIRST_DATA_ROW))
        now = get_text(row_ctl.hwnd).strip()
        if now == str(EXCEL_FIRST_DATA_ROW):
            print(f"  Hang dau chua du lieu: {now}  -> OK")
        else:
            print(f"  [CANH BAO] O 'Hang dau chua du lieu' dang hien {now!r}, mong muon {EXCEL_FIRST_DATA_ROW}.")
            verified = False

    # --- 4. Ghi ---
    if AUTO_CLICK_GHI and verified:
        for h, t in dialog_buttons(form):
            if norm_text(str(t).replace("&", "")) == "ghi":
                before = process_windows(pid)
                user32.PostMessageW(h, BM_CLICK, 0, 0)
                print("  Da bam Ghi.")
                n = click_ok_dialogs(pid, before)
                if n == 0:
                    print("  Khong thay hop thoai nao sau Ghi (co the ECUS da nhap xong).")
                if user32.IsWindowVisible(form):
                    print("  Form nhap Excel van con mo. Neu da nhap xong thi bam Dong.")
                else:
                    print("  Form nhap Excel da dong. Hoan tat nhap hang.")
                return True
        print("  [CANH BAO] Khong thay nut Ghi, hay bam tay.")
    elif AUTO_CLICK_GHI:
        print("  Khong tu bam Ghi vi co muc chua dung (xem canh bao o tren).")
    else:
        print("  Chua bam Ghi. Hay kiem tra form tren ECUS (file, sheet, hang dau, cot Excel) roi tu bam Ghi.")
    return verified


def import_goods(main_hwnd, excel_path):
    pid = process_id(main_hwnd)
    tabs = find_tab_controls(main_hwnd)
    if not tabs:
        print("  [HANG] Khong thay thanh tab, bo qua buoc nhap hang")
        return False
    tab_hwnd = tabs[0][0]

    # ---------- F11: xoa tat ca hang ----------
    print("\n[HANG] Buoc 1/2: F11 = XOA TOAN BO hang dang co trong ECUS")
    if CONFIRM_BEFORE_DELETE:
        if input("  Go XOA de dong y xoa het hang hien co (Enter = bo qua ca buoc xoa va nhap): "
                 ).strip().upper() != "XOA":
            print("  Da bo qua xoa / nhap hang.")
            return False

    if not focus_goods_grid(main_hwnd, tab_hwnd, pid):
        print("  [DUNG] Khong dua duoc ECUS len truoc de gui phim, dung lai de tranh bam nham cho khac.")
        return False
    dlg = press_and_wait(VK_F11, pid, 3.0)
    if dlg is None:
        print("  Khong thay hop thoai xac nhan sau F11 (co the danh sach da trong).")
        if CONFIRM_EMPTY_BEFORE_IMPORT:
            ans = input("  Neu danh sach CON hang: hay bam tay vao 1 o trong luoi hang tren ECUS roi bam Enter "
                        "de script bam lai F11. Neu danh sach da trong: go B: ")
            if ans.strip().upper() != "B":
                user32.SetForegroundWindow(main_hwnd)
                time.sleep(0.3)
                dlg = press_and_wait(VK_F11, pid, 3.0)
    if dlg is not None:
        handle_confirm_dialog(dlg)

    if CONFIRM_EMPTY_BEFORE_IMPORT:
        input("  KIEM TRA danh sach hang tren ECUS da TRONG chua. Bam Enter de tiep tuc nhap tu Excel "
              "(Ctrl+C de dung): ")
    else:
        time.sleep(1.0)   # cho ECUS xoa xong roi chuyen thang sang nhap Excel

    # ---------- F6: nhap tu file Excel ----------
    print(f"\n[HANG] Buoc 2/2: F6 = nhap tu file Excel: {excel_path}")
    if not focus_goods_grid(main_hwnd, tab_hwnd, pid):
        print("  [DUNG] Khong dua duoc ECUS len truoc de gui phim, dung lai.")
        return False
    dlg = press_and_wait(VK_F6, pid, 6.0)
    if dlg is None:
        print("  F6 khong mo hop thoai nao.")
        ans = input("  Hay bam tay vao 1 o trong luoi hang tren ECUS roi bam Enter de script bam lai F6 "
                    "(go B de bo qua): ")
        if ans.strip().upper() != "B":
            user32.SetForegroundWindow(main_hwnd)
            time.sleep(0.3)
            dlg = press_and_wait(VK_F6, pid, 6.0)
    if dlg is None:
        print("  Van khong mo duoc hop thoai chon file. Hay bam tay F6 va chon file tren ECUS.")
        return False

    print(f"  Cua so mo ra sau F6: title={get_title(dlg)!r}  class={get_class(dlg)}")
    return import_excel_form(dlg, pid, excel_path)


# ===== CHUONG TRINH CHINH =====
def pick_file():
    """Liet ke Template *.xlsx trong output cua moi thu muc khach, moi nhat len dau."""
    files = [f for f in DON_HANG_DIR.glob("*/output/Template *.xlsx") if not f.name.startswith("~$")]
    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    files = files[:SO_FILE_HIEN]
    if not files:
        print(f"Khong thay file Template *.xlsx nao trong {DON_HANG_DIR}\\<khach>\\output")
        return None
    print("Cac file Template (moi nhat len dau):")
    rong = max(len(f.parent.parent.name) for f in files)
    for i, f in enumerate(files):
        print(f"  {i:>2}: {f.parent.parent.name:<{rong}}  {f.name}")
    try:
        go = input("\nChon so thu tu file (Enter = so 0, file moi nhat): ").strip()
        chon = files[int(go) if go else 0]
    except (ValueError, IndexError):
        print("Lua chon khong hop le.")
        return None
    if "nhap" in chon.parent.parent.name.casefold():
        print("To khai NHAP dung man hinh ECUS khac, script nay moi lam cho to khai XUAT. Dung lai.")
        return None
    return chon


def main():
    path = pick_file()
    if path is None:
        return
    chay_ecus(path)


def _xac_nhan_man_hinh_den(cau_hoi):
    return input(cau_hoi + " (Enter = co | go N roi Enter = khong): ").strip().upper() != "N"


def chay_ecus(path, xac_nhan=None):
    """Dien 1 file Template vao ECUS. Tra ve True neu da dien xong.
    xac_nhan(cau_hoi) -> True/False: hoi truoc khi go vao ECUS (giao dien truyen vao; mac dinh hoi o man hinh den)."""
    if sys.platform != "win32":
        print("Script nay chi chay tren Windows.")
        return False
    if user32 is None:
        setup_windows()
    xac_nhan = xac_nhan or _xac_nhan_man_hinh_den
    path = Path(path)

    thieu, dia_diem = kiem_tra_template(path)
    if thieu:
        print(f"\n[DUNG] {path.name} CHUA DU THONG TIN, khong chay ECUS:")
        for x in thieu:
            print(f"   - {x}")
        print("Mo file nay bang Excel, dien cac o tren (dong 2 Sheet1 cung liet ke), Ctrl+S, roi chay lai.")
        return False

    values = read_template(path)
    values.update(dia_diem)
    print(f"\nDoc file: {path.name} (du thong tin)")

    main_hwnd = find_ecus_window()
    if main_hwnd is None:
        print("Khong thay cua so ECUS5 dang mo.")
        return False
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
        if key.startswith("pt_") and values.get("pt_giu"):
            if key == "pt_ma":
                print(f"{'Phuong thuc + phuong tien VC':<44}giu nhu cu (B3 = Nhu cu), khong dien")
            continue
        if key.endswith("_ten") and key not in XOA_KHI_TRONG and raw is None:
            continue   # chi go ten dia diem khi ma ZZZ
        if key == "ngay_di" and not raw:
            continue   # khong co booking -> khong dung vao ngay hang di
        if key in XOA_KHI_TRONG and raw == "":
            ctl, note = locate(items, label, idx)
            if ctl is None:
                print(f"{name:<22}{'(xoa trang)':<22}{'':<24}{note}")
            elif ctl.text.strip():
                print(f"{name:<22}{'(xoa trang)':<22}{ctl.text.strip()[:22]:<24}")
                plan.append((name, ctl, ""))
            continue   # o tren ECUS da trong thi khong can lam gi
        if raw is None or str(raw).strip() == "":
            where = CELLS.get(key, "B25 / B27 / C13")
            print(f"{name:<22}{'(TRONG)':<22}{'':<24}o {where} trong, bo qua "
                  "(mo file Template bang Excel, Ctrl+S neu la cong thuc chua tinh)")
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
        return False

    print("\nChua dien gi vao ECUS. Kiem tra bang tren cho dung.")
    if not xac_nhan(f"Bat dau dien {path.name} vao ECUS?"):
        print("Da thoat, khong thay doi gi.")
        return False

    print("\nBat dau sau 2 giay, dung cham chuot / ban phim cho den khi xong...")
    time.sleep(2)

    for name, ctl, text in plan:
        ok = write_field(main_hwnd, ctl.hwnd, text)
        if not ok:
            print(f"  [BO QUA] {name}: khong dat duoc focus vao o nay nen khong go (de tranh go nham cho)")
            continue
        now = get_text(ctl.hwnd).strip()
        status = "OK" if same_value(now, text) else f"KHAC (ECUS dang hien {now!r})"
        print(f"  {name:<22} -> {(text if len(text) <= 18 else text[:15] + '...'):<18} {status}")

    luu_nho(path, values)

    ok, why = go_to_tab(main_hwnd, DEFAULT_TAB)
    if ok:
        print(f"\nDa chuyen sang tab {DEFAULT_TAB} ({tab_label(DEFAULT_TAB)}).")
    else:
        print(f"\n[CANH BAO] Khong chuyen duoc sang tab {DEFAULT_TAB} ({tab_label(DEFAULT_TAB)}): {why}")

    if IMPORT_GOODS:
        if DEFAULT_TAB != GOODS_TAB:
            ok, why = go_to_tab(main_hwnd, GOODS_TAB)
        if ok:
            import_goods(main_hwnd, path.resolve())
        else:
            print(f"\n[BO QUA] Khong vao duoc tab {tab_label(GOODS_TAB)} nen khong xoa / nhap hang: {why}")

    ghi_lich_su(path)
    print("\nXong. HAY KIEM TRA lai cac o, tu nhap booking / ngay gio, roi moi bam Ghi. "
          "Script khong tu bam Ghi hay Gui.")
    return True


if __name__ == "__main__":
    main()