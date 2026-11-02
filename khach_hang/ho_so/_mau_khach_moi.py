"""
HO SO KHACH: <TEN KHACH>   <-- FILE MAU. Copy thanh ho_so/<ten_khach>.py (khong co dau gach duoi o dau) roi sua.

Hai cach doc hoa don:
  A. CACH_DOC = "ai"      -> KHONG can viet ham doc(). Claude doc hoa don va tra ve phieu chuan.
                             Chi can sua phan CAU HINH + GOI_Y_AI + ghi_chu(). Lam duoc ngay trong 15 phut.
  B. CACH_DOC = "quy_tac" -> viet ham doc() ben duoi. Mien phi, chay offline, ket qua co dinh.
                             Dung cho khach on dinh, nhieu lo. Xem ho_so/topglove.py lam vi du day du.

Thu tung buoc (chi doc va in ra, khong ghi file):
    python chay.py <ten_khach> --thu "duong\\dan\\hoa_don"
"""
import re

from chung import THU_MUC_DON, LoiDoc, chuan, dong_hang, ngay, phieu_moi, so

# ===== CAU HINH (sua phan nay) =====
TEN = "<Ten khach> xuat"
LOAI = "xuat"                                    # "xuat" / "nhap" (to khai nhap: man hinh ECUS khac, chua ho tro)
THU_MUC = THU_MUC_DON / "<Ten khach> xuat"      # thu muc cua khach: Template_<ma>.xlsx, hoa don, output\
THU_MUC_NGUON = THU_MUC
MAU_FILE = "**/*.docx"                           # ** = ca thu muc con. Vi du "**/INV*.docx", "*.xlsx", "**/*.pdf"
TEMPLATE = THU_MUC / "Template_<ma>.xlsx"       # copy tu Template_chuan, Sheet1 giu nguyen chu nhan

LOAI_TEP = [("Hoa don", "*.docx *.xlsx *.pdf")]   # loai tep cho nut "Them tep" tren giao dien

CACH_DOC = "ai"               # bat dau bang "ai"; khi da quen mau hoa don thi viet doc() va doi "quy_tac"
DUNG_AI_KHI_LOI = True        # quy tac bi loi (khach doi mau) -> tu chuyen sang AI

SHEET_HANG = "Sheet3"
DONG_BAT_DAU = 2
COT_HANG = {"ma_hang": "A", "ma_hs": "C", "mo_ta": "D", "so_luong": "E", "don_vi": "F",
            "don_gia": "H", "tri_gia": "I", "xuat_xu": "L"}
KHAI_HANG_KHONG_TINH_TIEN = False

# Nhung dieu dac biet cua hoa don khach nay, viet cho AI doc (cang cu the cang tot):
GOI_Y_AI = ("Vi du: gia ghi theo 1000 cai thi chia cho 1000. Nguoi nhan hang la o 'Consignee'. "
            "Dong 'FOC' la hang khong tinh tien.")


def ghi_chu(p):
    """Chuoi 'Ky hieu va so hieu' tren to khai cua khach nay. p la phieu chuan."""
    return f"PO: {p['po']}; PTTT: {p['payment']}; CNEE: {p['consignee']}"


# ===== DOC HOA DON BANG QUY TAC (chi can khi CACH_DOC = "quy_tac") =====
def doc(path):
    """Doc 1 hoa don, tra ve phieu chuan. Gap mau la thi raise LoiDoc("ly do").

    Goi y theo loai file:
      - Word : from doc_word import doc_word ; tl = doc_word(path) ; b = tl.tim_bang("INVOICE")
               b.tim_o("PO No.") ; b.o_duoi(o).chu ; b.o_tai(dong, cot).chu
               Xem cau truc bang: python doc_word.py "hoa_don.docx"
      - Excel: from openpyxl import load_workbook ; ws = load_workbook(path, data_only=True).active
               ws["B5"].value  hoac tim theo nhan nhu main.py
    """
    raise LoiDoc("Chua viet ham doc() cho khach nay")

    # ---- Khung de dien (xoa dong raise o tren khi bat dau viet) ----
    p = phieu_moi()
    # p["invoice_no"] = ...
    # p["invoice_date"] = ngay("21.10.2026")
    # p["currency"] = "USD"
    # p["incoterm"] = "FOB HO CHI MINH"
    # p["consignee"] = ...
    # p["packages_qty"], p["packages_unit"], p["gross_kg"] = 10, "CARTON", 123.4
    # p["items"].append(dong_hang(ma_hang=..., mo_ta=..., ma_hs=..., so_luong=..., don_vi="PCE",
    #                             don_gia=..., tri_gia=..., xuat_xu="VN"))
    # p["total_value"] = ...
    # if <dieu bat thuong>: p["canh_bao"].append("noi dung")
    p["note"] = ghi_chu(p)
    return p
