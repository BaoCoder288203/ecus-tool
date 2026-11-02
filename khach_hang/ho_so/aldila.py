"""
HO SO KHACH: ALDILA

Aldila da co bo xu ly rieng chay on dinh: D:\\test\\main.py (doc file "IV + PL G-xx-xxxx.xlsx",
copy bang hang vao Sheet2/Sheet3, dien Sheet1). Ho so nay chi goi lai main.py, khong viet lai.

Neu sau nay muon dua Aldila ve cung kieu voi cac khach khac thi viet ham doc() tra ve phieu chuan
giong ho_so/topglove.py va xoa ham chay_rieng() o duoi.
"""
import sys

from chung import GOC, THU_MUC_DON

TEN = "Aldila xuat (dung main.py)"
LOAI = "xuat"
THU_MUC = THU_MUC_DON / "Aldila xuat"   # IV + PL G-*.xlsx, Template_aldila.xlsx, output\
TEMPLATE = THU_MUC / "Template_aldila.xlsx"


LOAI_TEP = [("Excel IV + PL", "*.xlsx")]   # loai tep cho nut "Them tep" tren giao dien


def _nap_main():
    if str(GOC) not in sys.path:
        sys.path.insert(0, str(GOC))
    import main          # D:\test\main.py
    main.FOLDER = THU_MUC
    main.TEMPLATE = TEMPLATE
    main.OUTPUT = THU_MUC / "output"
    return main


def du_kien_dau_ra(files):
    """Ten file Template se tao tu cac file dau vao (de giao dien hoi truoc khi ghi de)."""
    m = _nap_main()
    return [m.OUTPUT / f"Template {m.get_inv(f.name)}.xlsx" for f in files if m.get_inv(f.name)]


def tao_template(files):
    m = _nap_main()
    lo = [f for f in files if m.get_inv(f.name)]
    da_tao, thieu = m.main(files=lo, mo_file=False)
    return da_tao, thieu, len(lo)


def chay_rieng():
    """chay.py thay ham nay thi giao toan bo viec cho main.py."""
    if str(GOC) not in sys.path:
        sys.path.insert(0, str(GOC))
    import main          # D:\test\main.py
    main.FOLDER = THU_MUC
    main.TEMPLATE = TEMPLATE
    main.OUTPUT = THU_MUC / "output"
    main.main()
