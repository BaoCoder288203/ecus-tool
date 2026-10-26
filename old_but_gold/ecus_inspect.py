"""
Liet ke cac o nhap / nut cua cua so ECUS dang mo, chi dung thu vien co san cua Python (ctypes).
Khong can cai pywinauto / pywin32.

Cach dung:
1. Mo ECUS, vao dung man hinh ban hay nhap tay (vi du: Thong tin chung).
2. Chay: python ecus_inspect2.py
3. Chon cua so ECUS trong danh sach, dat ten cho man hinh (vi du: thong_tin_chung).
4. Gui cho minh file D:\\test\\ecus_win32_<ten>.txt, kem 1 anh chup man hinh ECUS do.

Moi tab / man hinh khac nhau can chay lai 1 lan.
Script chi DOC thong tin, khong click hay go gi vao ECUS.
(Co them phan tuy chon dung thu vien 'uiautomation' neu ban da cai: pip install uiautomation)
"""
import ctypes
import sys
from ctypes import wintypes
from pathlib import Path

OUT_DIR = Path(r"D:\test")

user32 = ctypes.windll.user32
try:
    user32.SetProcessDPIAware()
except Exception:
    pass

WM_GETTEXT = 0x000D
SMTO_ABORTIFHUNG = 0x0002

EnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
user32.IsWindowVisible.argtypes = [wintypes.HWND]
user32.GetDlgCtrlID.argtypes = [wintypes.HWND]
user32.SendMessageTimeoutW.argtypes = [
    wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM,
    wintypes.UINT, wintypes.UINT, ctypes.POINTER(ctypes.c_size_t),
]
user32.SendMessageTimeoutW.restype = wintypes.LPARAM


def get_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def get_title(hwnd):
    buf = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(hwnd, buf, 512)
    return buf.value


def get_text(hwnd, maxlen=300):
    """Doc noi dung o nhap bang WM_GETTEXT (co gioi han thoi gian, khong treo neu ung dung dung hinh)."""
    buf = ctypes.create_unicode_buffer(maxlen)
    result = ctypes.c_size_t()
    ok = user32.SendMessageTimeoutW(
        hwnd, WM_GETTEXT, maxlen, ctypes.addressof(buf),
        SMTO_ABORTIFHUNG, 300, ctypes.byref(result),
    )
    return buf.value if ok else ""


def top_windows():
    found = []

    def cb(hwnd, _lparam):
        if user32.IsWindowVisible(hwnd) and get_title(hwnd).strip():
            found.append(hwnd)
        return True

    user32.EnumWindows(EnumProc(cb), 0)
    return found


def child_windows(parent):
    found = []

    def cb(hwnd, _lparam):
        found.append(hwnd)
        return True

    user32.EnumChildWindows(parent, EnumProc(cb), 0)
    return found


def dump_win32(hwnd, path):
    rows = []
    for ch in child_windows(hwnd):
        rect = wintypes.RECT()
        user32.GetWindowRect(ch, ctypes.byref(rect))
        text = get_text(ch).replace("\r", " ").replace("\n", " ")[:100]
        rows.append((
            rect.top, rect.left, rect.right, rect.bottom,
            get_class(ch), user32.GetDlgCtrlID(ch),
            bool(user32.IsWindowVisible(ch)), text,
        ))
    rows.sort()  # theo vi tri tren man hinh: tren xuong, trai sang phai

    lines = [f"Cua so: {get_title(hwnd)!r}  class={get_class(hwnd)}  (tong {len(rows)} dieu khien con)", ""]
    for i, (top, left, right, bottom, cls, cid, vis, text) in enumerate(rows):
        lines.append(
            f"{i:4d}  class={cls:<26} id={cid:<6} visible={vis!s:<5} "
            f"rect=({left},{top},{right},{bottom})  text={text!r}"
        )
    Path(path).write_text("\n".join(lines), encoding="utf-8")
    return len(rows)


def dump_uia(hwnd, path):
    """Tuy chon: neu da cai uiautomation thi ghi them cay dieu khien UIA."""
    try:
        import uiautomation as auto
    except Exception as e:
        print(f"  [uia] bo qua (khong import duoc uiautomation: {type(e).__name__})")
        return
    try:
        root = auto.ControlFromHandle(hwnd)
        lines, count = [], [0]

        def walk(c, depth):
            if count[0] > 4000 or depth > 12:
                return
            count[0] += 1
            r = c.BoundingRectangle
            lines.append(
                f"{'  ' * depth}{c.ControlTypeName} | class={c.ClassName} | name={c.Name!r} | "
                f"id={c.AutomationId!r} | rect=({r.left},{r.top},{r.right},{r.bottom})"
            )
            for child in c.GetChildren():
                walk(child, depth + 1)

        walk(root, 0)
        Path(path).write_text("\n".join(lines), encoding="utf-8")
        print(f"  [uia] da ghi {path} ({count[0]} dieu khien)")
    except Exception as e:
        print(f"  [uia] LOI: {e}")


def main():
    wins = top_windows()
    if not wins:
        print("Khong thay cua so nao.")
        return

    print("Cac cua so dang mo:")
    for i, h in enumerate(wins):
        print(f"  {i}: {get_title(h)!r}")

    try:
        hwnd = wins[int(input("\nNhap so thu tu cua cua so ECUS: ").strip())]
    except (ValueError, IndexError):
        print("Lua chon khong hop le.")
        return

    label = input("Dat ten cho man hinh nay (vi du thong_tin_chung): ").strip() or "man_hinh"
    OUT_DIR.mkdir(exist_ok=True)

    out = OUT_DIR / f"ecus_win32_{label}.txt"
    n = dump_win32(hwnd, out)
    print(f"\n  [win32] da ghi {out} ({n} dieu khien)")

    dump_uia(hwnd, OUT_DIR / f"ecus_uia_{label}.txt")

    print("\nXong. Gui cho minh file ecus_win32_*.txt va 1 anh chup man hinh ECUS do.")


if __name__ == "__main__":
    if sys.platform != "win32":
        print("Script nay chi chay tren Windows.")
        sys.exit(1)
    main()