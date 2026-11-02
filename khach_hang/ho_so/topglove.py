"""
HO SO KHACH: TOP GLOVE  (hoa don + packing list dang Word, bang nam trong khung chu)

Moi lo gom file "INV <so>.docx" va "PKL <so>.docx" (cung so). Co the de chung nhieu lo trong 1 thu muc
(lo cu) hoac moi lo 1 thu muc con (lo moi). Cac file khac (PDF, PI, F41, booking, MAT SAU...) bo qua.
  - Co ca "INV 301" va "INV 301 rev" -> lay ban rev.
  - INV cho: so invoice, ngay, PI, PO, CNEE, dong hang, tong tien.
  - PKL cho: so kien (khai theo BOX), Gross, Nett, so khoi. So invoice / PO / PI trong PKL phai khop INV.

Mau hoa don: bang co cac o "PI No.", "Date", "PO No.", "Term/Due Date", "INVOICE NUMBER",
cot Quantity (Carton, PCS/PRS) | Particulars | Unit Price | Total, gia ghi theo "Price / N Pieces".
Phia duoi co the co dong "NO COMMERCIAL VALUE, VALUE FOR CUSTOM DECLARATION PURPOSE"
(Empty Cartons, Empty Inners...) -> danh dau khong_tinh_tien.

Cach thu khi sua ho so nay:
    python chay.py topglove --thu "D:\\test\\topglove\\INV 301 rev.docx"
"""
import re
from pathlib import Path

from chung import THU_MUC_DON, LoiDoc, chuan, dong_hang, ngay, phieu_moi, so, so_gon
from doc_word import doc_word

# ===== CAU HINH =====
TEN = "Top Glove xuat"
LOAI = "xuat"                                   # "xuat" / "nhap"
THU_MUC = THU_MUC_DON / "Top Glove xuat"       # Template_topglove.xlsx, moi lo 1 thu muc con (INV + PKL), output\
THU_MUC_NGUON = THU_MUC
MAU_FILE = "**/*.doc*"                         # moi file Word (ca thu muc con); loc_file() chi giu INV
TEMPLATE = THU_MUC / "Template_topglove.xlsx"  # Template rieng cua khach nay (Sheet1 giong cac khach khac)

LOAI_TEP = [("Word INV + PKL", "*.docx *.doc")]   # loai tep cho nut "Them tep" tren giao dien

CACH_DOC = "quy_tac"          # "quy_tac" = doc bang code duoi day ; "ai" = gui cho Claude doc
DUNG_AI_KHI_LOI = False       # True = khi doc bang quy tac bi loi thi tu chuyen sang AI

# Dong hang ghi vao sheet nao, cot nao (khop voi cot Excel da cai trong form "Nhap tu Excel" cua ECUS)
SHEET_HANG = "Sheet3"
DONG_BAT_DAU = 2
COT_HANG = {"ma_hang": "A", "ma_hs": "C", "mo_ta": "D", "so_luong": "E", "don_vi": "F",
            "don_gia": "H", "tri_gia": "I", "xuat_xu": "L"}

KHAI_HANG_KHONG_TINH_TIEN = False   # True = khai ca Empty Cartons / Empty Inners (cong vao tong tri gia)
DON_VI_KIEN = "BOX"                 # khai theo box; cong thuc o Template doi chu nay thanh ma ECUS (BX)
DON_VI_HANG = "PCE"                 # ma don vi tinh hang cua ECUS cho cot PCS/PRS
MA_NUOC = {"VIETNAM": "VN", "VIET NAM": "VN"}
GOI_Y_AI = ("Hoa don gang tay Top Glove. Gia ghi theo 'Price / N Pieces': don_gia phai chia cho N. "
            "Dong duoi 'NO COMMERCIAL VALUE' la khong_tinh_tien. Consignee la ten sau 'Ship-to Address' "
            "neu co, neu khong thi la ten nguoi mua.")


def ghi_chu(p):
    """Chuoi 'Ky hieu va so hieu' tren to khai cua khach nay."""
    return f"PI: {p['pi']}; PO: {p['po']}; PTTT: {p['payment']}; CNEE: {p['consignee']}"


# ===== CHON FILE: INV <so> / PKL <so>, uu tien ban rev =====
def _so_trong_ten(path, loai):
    """'INV 301 rev.docx' + 'inv' -> ('301', True). Khong dung loai -> None."""
    m = re.match(rf"^\s*{loai}[\s_-]*0*(\d+)(.*)$", path.stem, re.I)
    if not m:
        return None
    return m.group(1), "rev" in m.group(2).casefold()


def _chon_ban(cac_file, loai):
    """Gom file theo (thu muc, so). Moi nhom giu 1 file: ban rev truoc, roi .docx truoc .doc."""
    nhom = {}
    for f in cac_file:
        k = _so_trong_ten(f, loai)
        if k is None:
            continue
        nhom.setdefault((f.parent, k[0]), []).append((k[1], f))
    chon = {}
    for khoa, ds in nhom.items():
        ds.sort(key=lambda x: (not x[0], x[1].suffix.lower() != ".docx", x[1].name))
        chon[khoa] = ds[0][1]
        bo = [f.name for _, f in ds[1:]]
        if bo:
            print(f"  [{loai.upper()} {khoa[1]}] dung {ds[0][1].name}, bo qua: {', '.join(bo)}")
    return chon


def loc_file(cac_file):
    """chay.py goi ham nay: chi giu file INV, moi so 1 file (uu tien ban rev)."""
    return sorted(_chon_ban(cac_file, "inv").values())


def _tim_pkl(inv_path):
    k = _so_trong_ten(inv_path, "inv")
    if k is None:
        return None
    cung_thu_muc = [f for f in inv_path.parent.iterdir()
                    if f.suffix.lower() in (".docx", ".doc") and not f.name.startswith("~$")]
    return _chon_ban(cung_thu_muc, "pkl").get((inv_path.parent, k[0]))


def doc_pkl(path):
    """Doc packing list: so invoice / PO / PI ghi ben trong + dong Grand Total (thung, cai, m3, gross, nett)."""
    tl = doc_word(path)
    b = tl.tim_bang("Quantity")
    if b is None or b.tim_o("Gross") is None:
        raise LoiDoc(f"{path.name}: khong thay bang co cot Gross")
    o_sl, o_gross, o_nett = b.tim_o("Quantity"), b.tim_o("Gross"), b.tim_o("Nett")
    o_m3 = b.tim_o("Measurement", bat_dau=True)
    o_mt = b.tim_o("Particulars")

    def chu(r, c):
        x = b.o_tai(r, c)
        return x.chu if x is not None and not x.noi_tiep else ""

    kq = {"file": path.name, "invoice_no": None, "po": None, "pi": None}
    for x in b.tat_ca_o():
        t = x.chu
        for khoa, nhan in (("invoice_no", r"^NO\.?\s*:"), ("po", r"^PO\s*NO\.?\s*:"), ("pi", r"^PI\s*NO\.?\s*:")):
            if kq[khoa] is None and re.match(nhan, t, re.I):
                kq[khoa] = re.sub(r"\s+", "", t.split(":", 1)[1])

    tong = None
    for r in range(len(b.dong)):
        ct = chuan(chu(r, o_mt.cot)) if o_mt else ""
        if ct in ("grand total", "sub total", "total"):
            dong = {"thung": so(chu(r, o_sl.cot)), "cai": so(chu(r, o_sl.cot + 1)),
                    "m3": so(chu(r, o_m3.cot)) if o_m3 else None,
                    "gross": so(chu(r, o_gross.cot)), "nett": so(chu(r, o_nett.cot)) if o_nett else None}
            if ct == "grand total" or tong is None:
                tong = dong
    if tong is None or tong["gross"] is None:
        raise LoiDoc(f"{path.name}: khong thay dong Grand Total / Sub Total co Gross")
    kq.update(tong)
    return kq


def _ghep_pkl(p, inv_path):
    """Lay so kien + Gross tu PKL cung so, doi chieu voi INV."""
    pkl_path = _tim_pkl(inv_path)
    if pkl_path is None:
        p["canh_bao"].append(f"Khong thay PKL cung so voi {inv_path.name}: so kien lay tu INV, Gross nhap tay")
        return
    try:
        k = doc_pkl(pkl_path)
    except LoiDoc as e:
        p["loi"].append(f"Khong doc duoc PKL: {e}")
        return
    p["gross_kg"] = so_gon(k["gross"], 3)
    p["net_kg"], p["cbm"], p["pkl_file"] = k["nett"], k["m3"], k["file"]

    # PKL phai la cua dung lo nay
    def khac(a, b):
        return a and b and re.sub(r"\s+", "", str(a)).casefold() != re.sub(r"\s+", "", str(b)).casefold()
    if khac(k["invoice_no"], p["invoice_no"]):
        p["loi"].append(f"{k['file']} ghi so invoice {k['invoice_no']}, khac INV {p['invoice_no']}")
    if khac(k["po"], p["po"]):
        p["loi"].append(f"PO trong PKL ({k['po']}) khac INV ({p['po']})")
    if khac(k["pi"], p["pi"]):
        p["loi"].append(f"PI trong PKL ({k['pi']}) khac INV ({p['pi']})")

    # So thung / so cai: PKL va INV (chi tinh hang tinh tien) phai bang nhau
    hang = [h for h in p["items"] if not h["khong_tinh_tien"]]
    cai_inv = sum(h["so_luong"] or 0 for h in hang)
    if k["cai"] is not None and abs(k["cai"] - cai_inv) > 0.5:
        p["loi"].append(f"So cai: PKL {k['cai']:,.0f} khac INV {cai_inv:,.0f}")
    if k["thung"] is not None and p["packages_qty"] is not None and abs(k["thung"] - p["packages_qty"]) > 0.5:
        p["loi"].append(f"So thung: PKL {k['thung']:,.0f} khac INV {p['packages_qty']:,.0f}")
    if k["thung"] is not None:
        p["packages_qty"] = so_gon(k["thung"])


# ===== DOC HOA DON =====
INCOTERMS = {"EXW", "FCA", "FAS", "FOB", "CFR", "CIF", "CPT", "CIP", "DAP", "DPU", "DDP"}
MA_SP = re.compile(r"[A-Z]{2,}\d{3}[A-Z0-9]*-\d+")      # vi du FVL043CSCL023CCV0N-24


def _ten_nguoi_nhan(dong_chu):
    """Khoi nguoi mua: dong dau la ten nguoi mua; neu co 'Ship-to Address' thi lay ten sau do
    (co the dai 2 dong) cho den dong bat dau bang so (dia chi)."""
    nguoi_mua = dong_chu[0] if dong_chu else None
    for i, d in enumerate(dong_chu):
        if chuan(d).startswith("ship-to address") or chuan(d).startswith("ship to address"):
            phan = []
            sau = d.split(":", 1)[1].strip() if ":" in d else ""
            if sau:
                phan.append(sau)
            for tiep in dong_chu[i + 1:]:
                if re.match(r"^\d", tiep.strip()):
                    break
                phan.append(tiep.strip())
            if phan:
                return nguoi_mua, " ".join(phan)
    return nguoi_mua, nguoi_mua


def doc(path):
    tl = doc_word(path)
    b = tl.tim_bang("INVOICE NUMBER")
    if b is None:
        raise LoiDoc("Khong thay bang co o 'INVOICE NUMBER'")
    p = phieu_moi()
    cb = p["canh_bao"]

    # --- So invoice (cung o voi nhan, hoac o ben duoi) ---
    o = b.tim_o("INVOICE NUMBER", bat_dau=True)
    sau_nhan = [d for d in o.dong_chu if chuan(d) != "invoice number"]
    p["invoice_no"] = sau_nhan[0].strip() if sau_nhan else (b.o_duoi(o).chu if b.o_duoi(o) else None)

    # --- Cac o nhan / gia tri nam duoi ---
    def duoi(nhan, lien=False):
        o = b.tim_o(nhan)
        o2 = b.o_duoi(o)
        if o2 is None:
            return None
        return (o2.chu_lien if lien else o2.chu) or None

    p["pi"] = duoi("PI No.")
    p["invoice_date"] = ngay(duoi("Date"))
    p["po"] = duoi("PO No.", lien=True)          # Word hay ngat dong giua so PO -> noi lien lai
    if p["invoice_date"] is None:
        cb.append(f"Khong doc duoc ngay hoa don: {duoi('Date')!r}")

    p["buyer"], p["consignee"] = _ten_nguoi_nhan(b.o_tai(0, 0).dong_chu if b.o_tai(0, 0) else [])

    # --- Vi tri cac cot theo dong tieu de ---
    o_sl = b.tim_o("Quantity")
    o_dg = b.tim_o("Unit Price")
    o_mt = b.tim_o("Particulars")
    if None in (o_sl, o_dg, o_mt):
        raise LoiDoc("Khong thay dong tieu de Quantity / Particulars / Unit Price")
    c_thung, c_cai, c_mt, c_dg = o_sl.cot, o_sl.cot + 1, o_mt.cot, o_dg.cot
    o_tt = next((x for x in b.dong[o_sl.dong] if chuan(x.chu) == "total"), None)
    if o_tt is None:
        raise LoiDoc("Khong thay cot Total")
    c_tt = o_tt.cot

    def chu(r, c):
        x = b.o_tai(r, c)
        return x.chu if x is not None and not x.noi_tiep else ""

    p["currency"] = chu(o_dg.dong + 1, c_dg) or None

    # --- Dieu kien giao hang: cac o dau tien co chu o cot Total ---
    phan_ic = []
    for r in range(o_dg.dong + 2, len(b.dong)):
        t = chu(r, c_tt)
        if so(t) is not None:
            break
        if t:
            phan_ic.append(t)
        if len(phan_ic) == 2:
            break
    if phan_ic and phan_ic[0].split()[0].upper() in INCOTERMS:
        p["incoterm"] = " ".join(phan_ic)
    else:
        cb.append(f"Khong nhan ra dieu kien giao hang: {phan_ic}")

    # --- Gia tinh theo bao nhieu cai ---
    gia_theo = 1
    for x in b.tat_ca_o():
        m = re.search(r"price\s*/\s*([\d,]+)\s*piece", x.chu, re.I)
        if m:
            gia_theo = int(so(m.group(1)))
            break

    # --- Thong tin san pham chung (ten, brand, HS, quy cach) ---
    ten_hang, ma_hs, cai_moi_thung = None, None, None
    for r in range(o_sl.dong + 1, len(b.dong)):
        t = chu(r, c_mt)
        ct = chuan(t)
        if ct.startswith("size"):
            break
        if not t or ct.startswith(("brand", "product code")):
            continue
        if ct.startswith("hs code"):
            ma_hs = re.sub(r"\D", "", t.split(":", 1)[-1])
        elif ct.startswith("packing style"):
            m = re.search(r"(\d+)\s*pcs\s*/\s*inner\s*;\s*(\d+)\s*inners?\s*/\s*carton", t, re.I)
            if m:
                cai_moi_thung = int(m.group(1)) * int(m.group(2))
        elif ten_hang is None:
            ten_hang = t
    if ma_hs and len(ma_hs) < 8:
        cb.append(f"Ma HS tren hoa don chi co {len(ma_hs)} so ({ma_hs}), ECUS can 8 so: bo sung tay")

    # --- Thong tin cuoi hoa don ---
    cuoi = "\n".join(x.chu for x in b.tat_ca_o())
    m = re.search(r"Payment Method:\s*([^\n/]+)", cuoi, re.I)
    if m:
        pt = m.group(1).strip()
        p["payment"] = "T/T" if re.search(r"\bTT\b|telegraphic", pt, re.I) else pt
    m = re.search(r"Country of Origin:\s*([A-Z ]+)", cuoi, re.I)
    p["origin"] = MA_NUOC.get(m.group(1).strip().upper(), m.group(1).strip()) if m else None
    m = re.search(r"B/L Date on/about:\s*([\d./-]+)", cuoi, re.I)
    p["ship_date"] = ngay(m.group(1)) if m else None

    # --- Dong hang + cac dong tong ---
    tong_hang = tong_kct = tong_ca = net = None
    vung_kct = False
    for r in range(o_sl.dong + 1, len(b.dong)):
        t = chu(r, c_mt)
        ct = chuan(t)
        thung, cai, dg, tt = so(chu(r, c_thung)), so(chu(r, c_cai)), so(chu(r, c_dg)), so(chu(r, c_tt))
        if "no commercial value" in ct:
            vung_kct = True
            continue
        if ct == "sub total":
            if vung_kct:
                tong_kct = (thung, cai, tt)
            else:
                tong_hang = (thung, cai, tt)
            continue
        if ct == "total":
            tong_ca = (thung, cai, tt)
            continue
        # cac dong 'Net Total' nam o cot khac -> doc theo chu cua ca dong
        dong_chu = b.chu_dong(r)
        if chuan(dong_chu).startswith("net total (incl"):
            net = so(dong_chu.split("|")[-1])
            continue
        if ct.startswith("size") and cai is not None:
            co = ct.split(":", 1)[-1].strip().upper()
            ma, lot = None, None
            for rr in (r, r + 1):
                m = MA_SP.search(b.chu_dong(rr)) if rr < len(b.dong) else None
                if m and ma is None:
                    ma = m.group(0)
            lot_chu = chu(r + 1, c_mt)
            if chuan(lot_chu).startswith("lot no"):
                lot = re.sub(r"\s+", " ", lot_chu.split(":", 1)[-1].replace(ma or "", "")).strip()
            h = dong_hang(ma_hang=ma, mo_ta=f"{ten_hang} - SIZE {co}" + (f" - LOT {lot}" if lot else ""),
                          ma_hs=ma_hs, so_luong=cai, don_vi=DON_VI_HANG,
                          don_gia=dg / gia_theo if dg is not None else None, tri_gia=tt, xuat_xu=p["origin"])
            h["_thung"] = thung
            p["items"].append(h)
            if cai_moi_thung and thung and abs(thung * cai_moi_thung - cai) > 0.5:
                cb.append(f"Size {co}: {thung:,.0f} thung x {cai_moi_thung} cai = "
                          f"{thung * cai_moi_thung:,.0f} khac {cai:,.0f} cai")
            continue
        if vung_kct and t and cai is not None and tt is not None:
            p["items"].append(dong_hang(mo_ta=t, so_luong=cai, don_vi="PKG", don_gia=dg, tri_gia=tt,
                                        xuat_xu=p["origin"], khong_tinh_tien=True))

    hang = [h for h in p["items"] if not h["khong_tinh_tien"]]
    kct = [h for h in p["items"] if h["khong_tinh_tien"]]
    if not hang:
        raise LoiDoc("Khong doc duoc dong hang nao (dong bat dau bang 'SIZE:')")

    # --- So kien, tong tri gia, doi chieu voi cac dong tong cua hoa don ---
    so_thung = sum(h.pop("_thung") or 0 for h in hang)
    p["packages_qty"] = so_gon(tong_hang[0] if tong_hang and tong_hang[0] else so_thung)
    p["packages_unit"] = DON_VI_KIEN
    if tong_hang:
        if tong_hang[0] is not None and abs(tong_hang[0] - so_thung) > 0.5:
            cb.append(f"Cong so thung cac dong = {so_thung:,.0f} khac Sub Total {tong_hang[0]:,.0f}")
        if tong_hang[1] is not None and abs(tong_hang[1] - sum(h["so_luong"] for h in hang)) > 0.5:
            cb.append("Cong so cai cac dong khac Sub Total")
    gia_tri_hang = net if net is not None else (tong_hang[2] if tong_hang else None)
    if gia_tri_hang is None:
        cb.append("Khong thay Net Total / Sub Total de lay tong tri gia")
    elif tong_ca and tong_ca[2] is not None and abs(tong_ca[2] - gia_tri_hang) > 0.01:
        cb.append(f"Total {tong_ca[2]:,.2f} khac Net Total {gia_tri_hang:,.2f}")
    them = sum(h["tri_gia"] for h in kct) if KHAI_HANG_KHONG_TINH_TIEN else 0
    p["total_value"] = so_gon(gia_tri_hang + them, 2) if gia_tri_hang is not None else None
    if kct and not KHAI_HANG_KHONG_TINH_TIEN:
        cb.append(f"Co {len(kct)} dong khong tinh tien (" + ", ".join(h["mo_ta"] for h in kct)
                  + ") chua dua vao to khai. Doi KHAI_HANG_KHONG_TINH_TIEN = True neu can khai.")

    _ghep_pkl(p, Path(path))
    p["note"] = ghi_chu(p)
    return p
