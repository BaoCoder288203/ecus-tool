"""
CHAY CHUNG CHO MOI KHACH HANG

    python chay.py                       -> hoi chon khach, xu ly tat ca hoa don trong thu muc cua khach do
    python chay.py topglove              -> chay luon khach topglove
    python chay.py topglove --thu FILE   -> CHI DOC 1 file va in ra phieu + kiem tra (khong ghi gi).
                                            Dung khi dang viet / sua ho so khach.
    python chay.py topglove --ai FILE    -> giong --thu nhung bat buoc doc bang AI (de so sanh voi quy tac)

Ket qua: D:\\test\\Don hang\\<khach>\\output\\Template <so invoice>.xlsx  -> sau do chay ecus_fill_demo.py.
File co LOI se co duoi " - CAN KIEM TRA" va 1 sheet "CANH BAO" ben trong.
"""
import importlib
import sys
from pathlib import Path

THU_MUC = Path(__file__).resolve().parent
sys.path.insert(0, str(THU_MUC))

from chung import LoiDoc, ghi_template, template_goc, in_phieu, kiem_tra, thu_muc_ra, tinh_lai_bang_excel  # noqa: E402


def danh_sach_ho_so():
    return sorted(p.stem for p in (THU_MUC / "ho_so").glob("*.py")
                  if not p.stem.startswith("_"))


def nap_ho_so(ten):
    return importlib.import_module(f"ho_so.{ten}")


def doc_phieu(ho_so, path, bat_buoc_ai=False):
    """Doc 1 hoa don theo cach ho so quy dinh. Tra ve (phieu, cach_da_dung)."""
    cach = "ai" if bat_buoc_ai else getattr(ho_so, "CACH_DOC", "quy_tac")
    if cach == "quy_tac":
        try:
            phieu = ho_so.doc(path)
            return phieu, "quy tac"
        except LoiDoc as e:
            if not getattr(ho_so, "DUNG_AI_KHI_LOI", False):
                raise
            print(f"  Doc bang quy tac khong duoc ({e}) -> chuyen sang AI")
    from doc_ai import doc_bang_ai
    phieu = doc_bang_ai(path, getattr(ho_so, "GOI_Y_AI", ""))
    if hasattr(ho_so, "ghi_chu"):
        phieu["note"] = ho_so.ghi_chu(phieu)
    return phieu, "AI"


def in_kiem_tra(kq):
    if not kq:
        print("  Kiem tra: OK, khong co loi")
    for muc, nd in kq:
        print(f"  [{muc}] {nd}")


def xu_ly_mot_file(ho_so, path):
    phieu, cach = doc_phieu(ho_so, path)
    kq = kiem_tra(phieu, ho_so)
    phieu["_ket_qua_kiem_tra"] = kq
    co_loi = any(m == "LOI" for m, _ in kq)
    so = str(phieu.get("invoice_no") or Path(path).stem).replace("/", "-").replace("\\", "-")
    ra = thu_muc_ra(ho_so) / (f"Template {so}" + (" - CAN KIEM TRA" if co_loi else "") + ".xlsx")
    # Xoa ban cu cua chinh lo nay (dang con lai tu lan chay truoc) de ECUS khong chon nham
    for cu in (thu_muc_ra(ho_so) / f"Template {so}.xlsx", thu_muc_ra(ho_so) / f"Template {so} - CAN KIEM TRA.xlsx"):
        if cu != ra and cu.exists():
            cu.unlink()
            print(f"  (da xoa ban cu: {cu.name})")
    loi_ghi = ghi_template(phieu, ho_so, ra)
    print(f"  Doc bang {cach} | {len(phieu['items'])} dong hang | tong {phieu.get('total_value')}")
    in_kiem_tra(kq + [("CANH BAO", x) for x in loi_ghi])
    print(f"  -> {ra.name}")
    return ra


def tao_template(ho_so, files):
    """Giao dien goi ham nay: tao Template cho danh sach file dau vao cua 1 khach.
    Tra ve (file da tao, file con thieu, so lo du kien)."""
    files = [Path(f) for f in files if not Path(f).name.startswith("~$")]
    ho_so.TEMPLATE = template_goc(ho_so)          # Template_<khach>.xlsx (doi ten file cu Template.xlsx neu con)
    if hasattr(ho_so, "tao_template"):            # khach co bo xu ly rieng (Aldila -> main.py)
        return ho_so.tao_template(files)
    if not Path(ho_so.TEMPLATE).exists():
        print(f"Khong thay Template cua khach: {ho_so.TEMPLATE}")
        return [], [], 0
    lo = ho_so.loc_file(files) if hasattr(ho_so, "loc_file") else files
    da_tao = []
    for f in lo:
        print(f"\nDang xu ly: {f.name}")
        try:
            da_tao.append(xu_ly_mot_file(ho_so, f))
        except Exception as e:
            print(f"  LOI: {e}")
    thieu = tinh_lai_bang_excel(da_tao)
    if thieu is None:
        thieu = list(da_tao)
    print(f"\nHoan tat: {len(da_tao)}/{len(lo)} lo. Ket qua o {thu_muc_ra(ho_so)}")
    return da_tao, thieu, len(lo)


def main():
    args = sys.argv[1:]
    ds = danh_sach_ho_so()
    if args and args[0] in ds:
        ten = args[0]
    else:
        print("Cac khach hang:")
        for i, t in enumerate(ds):
            print(f"  {i}: {t}  ({getattr(nap_ho_so(t), 'TEN', t)})")
        try:
            ten = ds[int(input("Chon so thu tu khach: ").strip())]
        except (ValueError, IndexError):
            print("Lua chon khong hop le.")
            return
    ho_so = nap_ho_so(ten)
    print(f"\n=== {getattr(ho_so, 'TEN', ten)} ===")

    # Che do thu: chi doc va in ra
    if len(args) >= 3 and args[1] in ("--thu", "--ai"):
        phieu, cach = doc_phieu(ho_so, args[2], bat_buoc_ai=(args[1] == "--ai"))
        print(f"Doc bang {cach}: {args[2]}\n")
        in_phieu(phieu, ho_so)
        print()
        in_kiem_tra(kiem_tra(phieu, ho_so))
        return

    # Khach co bo xu ly rieng (vi du Aldila dung main.py)
    if hasattr(ho_so, "chay_rieng"):
        ho_so.chay_rieng()
        return

    thu_muc = Path(ho_so.THU_MUC)
    if not thu_muc.exists():
        (thu_muc / "output").mkdir(parents=True)
        print(f"Da tao thu muc {thu_muc}. Bo Template_<khach>.xlsx va hoa don vao do roi chay lai.")
        return
    ho_so.TEMPLATE = template_goc(ho_so)
    if not Path(ho_so.TEMPLATE).exists():
        print(f"Khong thay Template cua khach: {ho_so.TEMPLATE}")
        return
    cac_file = sorted(f for f in Path(ho_so.THU_MUC_NGUON).glob(ho_so.MAU_FILE)
                      if not f.name.startswith("~$") and thu_muc_ra(ho_so) not in f.parents)
    if hasattr(ho_so, "loc_file"):          # vi du Top Glove: chi lay INV, uu tien ban rev
        cac_file = ho_so.loc_file(cac_file)
    if not cac_file:
        print(f"Khong co hoa don nao ({ho_so.MAU_FILE}) trong {ho_so.THU_MUC_NGUON}")
        return

    da_tao = []
    for f in cac_file:
        print(f"\nDang xu ly: {f.name}")
        try:
            da_tao.append(xu_ly_mot_file(ho_so, f))
        except Exception as e:
            print(f"  LOI: {e}")
    tinh_lai_bang_excel(da_tao)
    print(f"\nHoan tat: {len(da_tao)}/{len(cac_file)} file. Ket qua o {thu_muc_ra(ho_so)}")


if __name__ == "__main__":
    main()
