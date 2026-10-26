"""
Chuyen sang mot tab cua ECUS (vi du "Danh sach hang") bang script. Chi dung thu vien co san cua Python.

Script CHI chuyen tab, khong go vao o nao, khong bam Ghi / Gui.

Cach dung:
1. Mo ECUS, mo mot to khai (de thay cac tab: Thong tin chung, Thong tin Container, Danh sach hang...).
2. Chay: python ecus_goto_tab.py        (hoac: python ecus_goto_tab.py 2)
3. Chon so thu tu tab (tinh tu 0, tu trai sang phai). Theo anh ban gui, "Danh sach hang" la so 2.

Cach hoat dong: tim thanh tab (SysTabControl32), hoi ECUS vi tri tab can chon, roi gui mot cu "click"
vao dung tab do. Khong di chuyen chuot that cua ban.
Neu bao khong doc duoc vi tri tab: thu chay script bang quyen Administrator (neu ECUS dang chay quyen Admin).
"""
import ctypes
import sys
import time
from ctypes import wintypes

# Thu tu tab theo anh chup (tinh tu 0). Chi dung de hien thi cho de nhin.
TAB_NAMES = [
    "Thong tin chung",
    "Thong tin Container",
    "Danh sach hang",
    "Chi thi cua Hai quan",
    "Ket qua xu ly to khai",
    "Quan ly to khai",
]

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

user32 = kernel32 = None
EnumProc = None


def setup():
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
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.SendMessageW.restype = wintypes.LPARAM

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


def find_ecus_window():
    found = []

    def cb(hwnd, _):
        if user32.IsWindowVisible(hwnd) and "ECUS5" in get_title(hwnd).upper():
            found.append(hwnd)
        return True

    user32.EnumWindows(EnumProc(cb), 0)
    return found[0] if found else None


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
        tabs.append((h, count, (r.right - r.left) * (r.bottom - r.top), (r.left, r.top, r.right, r.bottom)))
    tabs.sort(key=lambda t: t[2], reverse=True)
    return tabs


def tab_item_rect(tab_hwnd, index):
    """Hoi ECUS vi tri (toa do trong thanh tab) cua tab thu 'index'. Can bo nho tam trong tien trinh ECUS."""
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


def select_tab(tab_hwnd, index):
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


def main():
    if sys.platform != "win32":
        print("Script nay chi chay tren Windows.")
        return
    setup()

    main_hwnd = find_ecus_window()
    if main_hwnd is None:
        print("Khong thay cua so ECUS5 dang mo.")
        return

    tabs = find_tab_controls(main_hwnd)
    if not tabs:
        print("Khong thay thanh tab nao (SysTabControl32). Hay mo mot to khai truoc roi chay lai.")
        return

    print("Cac thanh tab tim thay (lay cai lon nhat lam thanh tab chinh):")
    for i, (h, count, _area, rect) in enumerate(tabs):
        print(f"  {'*' if i == 0 else ' '} hwnd={h}  so tab={count}  rect={rect}")
    tab_hwnd, count, _, _ = tabs[0]
    current = user32.SendMessageW(tab_hwnd, TCM_GETCURSEL, 0, 0)
    print(f"\nThanh tab chinh co {count} tab, dang o tab so {current}.")
    print("Thu tu theo anh chup:")
    for i, name in enumerate(TAB_NAMES):
        print(f"  {i}: {name}")

    if len(sys.argv) > 1:
        raw = sys.argv[1]
    else:
        raw = input("\nChon so thu tu tab can chuyen den: ").strip()
    try:
        index = int(raw)
    except ValueError:
        print("Lua chon khong hop le.")
        return
    if not (0 <= index < count):
        print(f"So tab phai tu 0 den {count - 1}.")
        return

    ok, why = select_tab(tab_hwnd, index)
    if ok:
        print(f"Da chuyen sang tab so {index}.")
    else:
        print(f"Khong chuyen duoc: {why}")


if __name__ == "__main__":
    main()
