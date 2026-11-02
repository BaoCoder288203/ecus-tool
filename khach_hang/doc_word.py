"""
Doc file Word (.docx) thanh cac BANG va DOAN VAN, ke ca bang nam trong KHUNG CHU (text box).

Vi sao can file nay: thu vien python-docx bo qua bang nam trong khung chu, ma nhieu hoa don
(vi du Top Glove) dat ca bang hoa don trong khung chu. File nay doc thang XML cua Word.

Dung trong ho so khach hang:
    from doc_word import doc_word
    w = doc_word(duong_dan)
    bang = w.tim_bang("INVOICE NUMBER")    # bang dau tien co chu nay
    o = bang.tim_o("PI No.")               # o co dung chu nay
    gia_tri = bang.o_duoi(o)               # o ngay ben duoi (cung cot)

File .doc (Word cu) se duoc tu chuyen sang .docx bang Word (can Word + pywin32), hoac ban tu
mo bang Word roi "Save As" .docx.
"""
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"w": W, "mc": MC}
TBL, P, T, FALLBACK = f"{{{W}}}tbl", f"{{{W}}}p", f"{{{W}}}t", f"{{{MC}}}Fallback"

_cha = {}   # phan tu con -> phan tu cha (ElementTree khong tu luu)


def _to_tien(el):
    while el in _cha:
        el = _cha[el]
        yield el


def chuan(s):
    """Chu de so sanh: bo khoang trang thua, bo ':' cuoi, khong phan biet hoa thuong."""
    return re.sub(r"\s+", " ", str(s or "")).strip().rstrip(":").strip().casefold()


@dataclass
class O:
    """Mot o trong bang Word."""
    dong: int          # so thu tu dong (tu 0)
    cot: int           # cot bat dau tren luoi (tu 0), da tinh o gop ngang
    rong: int          # so cot luoi o nay chiem (o gop ngang > 1)
    noi_tiep: bool     # True = o nay la phan duoi cua o gop doc (khong co chu rieng)
    dong_chu: list     # cac dong chu (moi doan van trong o la 1 phan tu)

    @property
    def chu(self):
        """Toan bo chu trong o, cac dong noi bang xuong dong."""
        return "\n".join(self.dong_chu).strip()

    @property
    def chu_lien(self):
        """Chu trong o noi lien khong cach (dung cho so PO bi Word ngat dong giua chung)."""
        return re.sub(r"\s+", "", "".join(self.dong_chu))

    def chua(self, cot):
        return self.cot <= cot < self.cot + self.rong


@dataclass
class Bang:
    dong: list = field(default_factory=list)   # list[list[O]]

    def o_tai(self, dong, cot):
        """O phu cot luoi 'cot' tren dong 'dong' (None neu khong co)."""
        if not 0 <= dong < len(self.dong):
            return None
        for o in self.dong[dong]:
            if o.chua(cot):
                return o
        return None

    def tat_ca_o(self):
        for hang in self.dong:
            yield from hang

    def tim_o(self, chu, bat_dau=False):
        """O dau tien co chu dung bang 'chu' (bat_dau=True: chi can bat dau bang 'chu')."""
        muc = chuan(chu)
        for o in self.tat_ca_o():
            c = chuan(o.chu)
            if c == muc or (bat_dau and c.startswith(muc)):
                return o
        return None

    def o_duoi(self, o, buoc=1):
        """O ngay ben duoi 'o' (cung cot luoi), bo qua o noi tiep cua o gop doc."""
        if o is None:
            return None
        d = o.dong + 1
        da_qua = 0
        while d < len(self.dong):
            duoi = self.o_tai(d, o.cot)
            if duoi is not None and not duoi.noi_tiep:
                da_qua += 1
                if da_qua == buoc:
                    return duoi
            d += 1
        return None

    def chu_dong(self, dong):
        """Toan bo chu tren 1 dong (de tim nhanh)."""
        return " | ".join(o.chu for o in self.dong[dong] if o.chu)

    def in_ra(self):
        """In bang ra man hinh de nhin cau truc khi viet ho so khach moi."""
        for i, hang in enumerate(self.dong):
            phan = []
            for o in hang:
                nhan = f"[c{o.cot}" + (f"-{o.cot + o.rong - 1}" if o.rong > 1 else "") + "]"
                phan.append(nhan + (" ^" if o.noi_tiep else repr(o.chu.replace("\n", " / "))))
            print(f"  dong {i:02d}: " + "  ".join(phan))


@dataclass
class TaiLieuWord:
    bang: list            # list[Bang]
    doan_van: list        # cac doan van ngoai bang (list[str])

    def tim_bang(self, chu):
        """Bang dau tien co mot o bat dau bang 'chu'."""
        for b in self.bang:
            if b.tim_o(chu, bat_dau=True) is not None:
                return b
        return None

    def van_ban(self):
        """Toan bo noi dung dang chu (bang ghi theo dong, cac o cach nhau ' | '), dung de gui cho AI."""
        phan = [d for d in self.doan_van if d.strip()]
        for i, b in enumerate(self.bang):
            phan.append(f"\n[BANG {i + 1}]")
            for d in range(len(b.dong)):
                o_chu = [o.chu.replace("\n", " / ") for o in b.dong[d] if not o.noi_tiep]
                if any(o_chu):
                    phan.append(" | ".join(o_chu))
        return "\n".join(phan)


def _chu_doan(p):
    """Chu cua chinh doan van p, KHONG lay chu cua khung chu / doan van long ben trong no."""
    phan = []
    for t in p.iter(T):
        gan_nhat = next((a for a in _to_tien(t) if a.tag == P), None)
        if gan_nhat is p:
            phan.append(t.text or "")
    return "".join(phan)


def _doc_o(tc):
    dong_chu = []
    for p in tc.findall("w:p", NS):
        t = _chu_doan(p)
        if t.strip():
            dong_chu.append(t.strip())
    return dong_chu


def _doc_bang(tbl):
    b = Bang()
    for i, tr in enumerate(tbl.findall("w:tr", NS)):
        cot = 0
        bo_qua = tr.find("w:trPr/w:gridBefore", NS)
        if bo_qua is not None:
            cot = int(bo_qua.get(f"{{{W}}}val", "0"))
        hang = []
        for tc in tr.findall("w:tc", NS):
            span = tc.find("w:tcPr/w:gridSpan", NS)
            rong = int(span.get(f"{{{W}}}val")) if span is not None else 1
            vm = tc.find("w:tcPr/w:vMerge", NS)
            noi_tiep = vm is not None and vm.get(f"{{{W}}}val") in (None, "continue")
            hang.append(O(i, cot, rong, noi_tiep, [] if noi_tiep else _doc_o(tc)))
            cot += rong
        b.dong.append(hang)
    return b


def _chuyen_doc_sang_docx(path):
    """Mo file .doc bang Word va luu thanh .docx canh ben (can Word + pywin32)."""
    import win32com.client as win32   # chi can khi gap file .doc
    moi = path.with_suffix(".docx")
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    try:
        d = word.Documents.Open(str(path.resolve()), ReadOnly=True)
        d.SaveAs2(str(moi.resolve()), FileFormat=16)   # 16 = .docx
        d.Close(False)
    finally:
        word.Quit()
    return moi


def doc_word(path):
    """Doc file Word. Tra ve TaiLieuWord (bang + doan van)."""
    path = Path(path)
    if path.suffix.lower() == ".doc":
        path = _chuyen_doc_sang_docx(path)

    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    _cha.clear()
    _cha.update({con: cha for cha in root.iter() for con in cha})

    def co_to_tien(el, tag):
        return any(a.tag == tag for a in _to_tien(el))

    # Word luu khung chu 2 lan (ban chinh + ban du phong cho Word cu). Chi doc ban chinh.
    bang = [_doc_bang(t) for t in root.iter(TBL)
            if not co_to_tien(t, FALLBACK) and not co_to_tien(t, TBL)]

    doan_van = []
    for p in root.iter(P):
        if co_to_tien(p, FALLBACK) or co_to_tien(p, TBL):
            continue
        t = _chu_doan(p).strip()
        if t:
            doan_van.append(t)
    return TaiLieuWord(bang, doan_van)


if __name__ == "__main__":
    # Chay thu: python doc_word.py "D:\test\topglove\INV 301 rev.docx"
    # -> in ra cau truc bang de ban nhin khi viet ho so khach moi.
    import sys
    if len(sys.argv) < 2:
        print('Cach dung: python doc_word.py "duong\\dan\\hoa_don.docx"')
        raise SystemExit(1)
    tl = doc_word(sys.argv[1])
    print("DOAN VAN NGOAI BANG:")
    for d in tl.doan_van:
        print("  ", d)
    for i, b in enumerate(tl.bang):
        print(f"\nBANG {i}: {len(b.dong)} dong")
        b.in_ra()
