# -*- coding: utf-8 -*-
"""gdi_print.py —— 直接用 Windows GDI 送印 PDF（不再借道 SumatraPDF）

为什么不用 SumatraPDF 打印：实测同一份 7 页讲义、同一个 D95 驱动，
SumatraPDF 的 `-print-to` 要 53.9 秒，而这 54 秒里它自己的 CPU 只用了 1.7 秒
（基本在干等）；把同样的 600dpi 位图直接 StretchDIBits 给同一个驱动只要 2.3 秒，
输出的打印数据大小还一样（4.06MB vs 4.10MB）。所以送印这一段自己走 GDI：

    取打印机默认 DEVMODE → 改纸张/双面/份数/颜色 → CreateDC
    → StartDoc →（每页 StartPage + 渲染 + StretchDIBits + EndPage）→ EndDoc

好处不只是快：
- 纸张/双面/份数/颜色直接落在 DEVMODE 上，不用再靠 Sumatra 的 paperkind 绕
  （原先 paper=A4 匹配不上驱动里的 "A4(210x297mm)"，会退回自定义纸张再落到 A3）
- 位图按驱动 DPI 1:1 送，不经缩放重采样
- 送印在独立子进程里跑，驱动崩了只崩这个子进程，网页服务不受影响

用法（父进程 print_system.py 就是这么调的）：
    python gdi_print.py <pdf> <打印机名> [--copies N] [--paper A4] [--duplex single]
                        [--scale fit] [--color 1] [--range 1-3,5]
    python gdi_print.py --probe <打印机名>      # 只读设备能力，不出纸

退出码：0 成功；2 还没往打印机送任何数据就失败了（可以安全地换引擎重试）；
3 打到一半失败（别重试，否则重复出纸）；4 参数错误。
"""
import ctypes
import os
import sys
from ctypes import wintypes

try:
    import pymupdf as fitz
except ImportError:                       # 老版本 PyMuPDF
    import fitz

try:
    import numpy as np
except ImportError:                       # 没有 numpy 就只能走 SumatraPDF 兜底
    np = None

BASE = os.path.dirname(os.path.abspath(__file__))

# DMPAPER_* 编号。必须给编号不能给纸名：本机驱动里的纸叫 "A4(210x297mm)"，
# 按名字传会被拒，驱动退回"按文档尺寸自定义纸张"，找不到就落到 A3。
PAPER_ID = {"A4": 9, "A3": 8, "A5": 11, "B4": 12, "B5": 13,
            "Letter": 1, "Legal": 5, "11x17": 3}
# DMDUP_*
DUPLEX_ID = {"single": 1, "long": 2, "short": 3}
# custom（自定义比例）在 prepare 里已经把比例烘进 PDF 了，送印时按 100% 打
SCALE_ID = {"fit": "fit", "shrink": "shrink", "noscale": "noscale", "custom": "noscale"}

DM_ORIENTATION = 0x00000001
DM_PAPERSIZE = 0x00000002
DM_COPIES = 0x00000100
DM_COLOR = 0x00000800
DM_DUPLEX = 0x00001000
DM_COLLATE = 0x00008000
DM_IN_BUFFER = 0x00000008
DM_OUT_BUFFER = 0x00000002

DMORIENT_PORTRAIT = 1
DMCOLOR_MONOCHROME = 1
DMCOLOR_COLOR = 2
DMCOLLATE_TRUE = 1

HORZRES, VERTRES = 8, 10
LOGPIXELSX, LOGPIXELSY = 88, 90
PHYSICALWIDTH, PHYSICALHEIGHT = 110, 111
SRCCOPY = 0x00CC0020
COLORONCOLOR, HALFTONE = 3, 4

# 单页位图上限（像素）。A3 彩打 600dpi 是 70M 像素（约 210MB），A0 就到了 140M，
# 超过上限就按比例降分辨率渲染、再交给 StretchDIBits 拉伸回原尺寸，别把内存吃穿。
MAX_PIXELS = 120_000_000
MM_PER_INCH = 25.4

# 每次 StretchDIBits 送多少行。整页一次性丢给驱动，遇到 A0/A1 或复杂 CAD 图时
# 有些 PCL 驱动会吃不下（表现为图上整条整条空白），分带后每次只有千余行，稳得多。
BAND_ROWS = 1024
# 图纸比纸大多少算"放不下"。打印机边距造成的 3~5% 溢出属正常（100% 打 A4 文档时
# 左右各裁掉 4.5mm），不触发自动缩小，否则「实际大小」档就名存实亡。
OVER_TOL = 1.10


class PrintError(Exception):
    """code 就是进程退出码，父进程靠它决定要不要换引擎重试：
    2 = 设备/驱动层失败，一页都没送出去，换引擎重试是安全的
    3 = 打到一半失败，别重试，否则重复出纸、重复计费
    5 = 参数或文件本身有问题（页码全越界、文件没了），换什么引擎都没用
    """

    def __init__(self, msg, code=2):
        super().__init__(msg)
        self.code = code


class DEVMODEW(ctypes.Structure):
    """DEVMODEW 的固定部分（220 字节）。后面还可能跟 dmDriverExtra 字节的驱动私有数据，
    所以实际缓冲区比这个结构大，改字段时只碰前面这段。"""
    _fields_ = [
        ("dmDeviceName", wintypes.WCHAR * 32),
        ("dmSpecVersion", wintypes.WORD),
        ("dmDriverVersion", wintypes.WORD),
        ("dmSize", wintypes.WORD),
        ("dmDriverExtra", wintypes.WORD),
        ("dmFields", wintypes.DWORD),
        ("dmOrientation", ctypes.c_short),
        ("dmPaperSize", ctypes.c_short),
        ("dmPaperLength", ctypes.c_short),
        ("dmPaperWidth", ctypes.c_short),
        ("dmScale", ctypes.c_short),
        ("dmCopies", ctypes.c_short),
        ("dmDefaultSource", ctypes.c_short),
        ("dmPrintQuality", ctypes.c_short),
        ("dmColor", ctypes.c_short),
        ("dmDuplex", ctypes.c_short),
        ("dmYResolution", ctypes.c_short),
        ("dmTTOption", ctypes.c_short),
        ("dmCollate", ctypes.c_short),
        ("dmFormName", wintypes.WCHAR * 32),
        ("dmLogPixels", wintypes.WORD),
        ("dmBitsPerPel", wintypes.DWORD),
        ("dmPelsWidth", wintypes.DWORD),
        ("dmPelsHeight", wintypes.DWORD),
        ("dmDisplayFlags", wintypes.DWORD),
        ("dmDisplayFrequency", wintypes.DWORD),
        ("dmICMMethod", wintypes.DWORD),
        ("dmICMIntent", wintypes.DWORD),
        ("dmMediaType", wintypes.DWORD),
        ("dmDitherType", wintypes.DWORD),
        ("dmReserved1", wintypes.DWORD),
        ("dmReserved2", wintypes.DWORD),
        ("dmPanningWidth", wintypes.DWORD),
        ("dmPanningHeight", wintypes.DWORD),
    ]


class DOCINFOW(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_int), ("lpszDocName", wintypes.LPCWSTR),
                ("lpszOutput", wintypes.LPCWSTR), ("lpszDatatype", wintypes.LPCWSTR),
                ("fwType", wintypes.DWORD)]


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
                ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
                ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
                ("biClrImportant", wintypes.DWORD)]


class BITMAPINFO(ctypes.Structure):
    # 24 位真彩不需要调色板，留一个颜色槽让结构对齐即可
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 1)]


_api = None


def api():
    """gdi32 / winspool 的函数原型只需要设一次。不设 argtypes 的话 64 位 HDC
    会按 int32 传，直接 OverflowError。"""
    global _api
    if _api is not None:
        return _api
    gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
    spool = ctypes.WinDLL("winspool.drv", use_last_error=True)
    HDC, HANDLE = wintypes.HDC, wintypes.HANDLE

    spool.OpenPrinterW.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(HANDLE), ctypes.c_void_p]
    spool.OpenPrinterW.restype = wintypes.BOOL
    spool.ClosePrinter.argtypes = [HANDLE]
    spool.ClosePrinter.restype = wintypes.BOOL
    spool.DocumentPropertiesW.argtypes = [wintypes.HWND, HANDLE, wintypes.LPCWSTR,
                                         ctypes.POINTER(DEVMODEW), ctypes.POINTER(DEVMODEW),
                                         wintypes.DWORD]
    spool.DocumentPropertiesW.restype = ctypes.c_long

    gdi32.CreateDCW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_void_p]
    gdi32.CreateDCW.restype = HDC
    gdi32.DeleteDC.argtypes = [HDC]
    gdi32.GetDeviceCaps.argtypes = [HDC, ctypes.c_int]
    gdi32.GetDeviceCaps.restype = ctypes.c_int
    gdi32.StartDocW.argtypes = [HDC, ctypes.POINTER(DOCINFOW)]
    gdi32.StartDocW.restype = ctypes.c_int
    gdi32.StartPage.argtypes = [HDC]
    gdi32.StartPage.restype = ctypes.c_int
    gdi32.EndPage.argtypes = [HDC]
    gdi32.EndPage.restype = ctypes.c_int
    gdi32.EndDoc.argtypes = [HDC]
    gdi32.EndDoc.restype = ctypes.c_int
    gdi32.AbortDoc.argtypes = [HDC]
    gdi32.SetMapMode.argtypes = [HDC, ctypes.c_int]
    gdi32.SetStretchBltMode.argtypes = [HDC, ctypes.c_int]
    gdi32.SetBrushOrgEx.argtypes = [HDC, ctypes.c_int, ctypes.c_int, ctypes.c_void_p]
    gdi32.StretchDIBits.argtypes = [HDC] + [ctypes.c_int] * 8 + [
        ctypes.c_void_p, ctypes.POINTER(BITMAPINFO), wintypes.UINT, wintypes.DWORD]
    gdi32.StretchDIBits.restype = ctypes.c_int
    _api = {"gdi32": gdi32, "spool": spool}
    return _api


def available():
    """能不能走 GDI 送印（非 Windows 或缺 numpy 就不行）。"""
    if os.name != "nt" or np is None:
        return False
    try:
        api()
        return True
    except OSError:
        return False


# ---------------- 页码范围 ----------------

def parse_range(spec, total):
    """'1-3,5' -> [1,2,3,5]。空串表示全部。越界的页直接丢掉。"""
    s = str(spec or "").strip()
    if not s:
        return list(range(1, total + 1))
    out = []
    for tok in s.replace(" ", "").replace("，", ",").split(","):
        if not tok:
            continue
        if "-" in tok[1:]:                     # 别把开头的负号当范围
            a, _, b = tok.partition("-")
            lo, hi = int(a), int(b or total)
            out.extend(range(min(lo, hi), max(lo, hi) + 1))
        else:
            out.append(int(tok))
    return [n for n in out if 1 <= n <= total]


# ---------------- DEVMODE ----------------

def _open_devmode(printer, paper, duplex, copies, color):
    """打开打印机，取出默认 DEVMODE，把我们关心的字段改掉，再让驱动校验一遍。"""
    a = api()
    hpr = wintypes.HANDLE()
    if not a["spool"].OpenPrinterW(printer, ctypes.byref(hpr), None):
        raise PrintError(f"打不开打印机「{printer}」（错误码 {ctypes.get_last_error()}）")
    try:
        # 先用 0 问一下需要多大缓冲区（含驱动私有数据），再取默认值
        size = a["spool"].DocumentPropertiesW(None, hpr, printer, None, None, 0)
        if size <= 0:
            raise PrintError(f"读不到打印机「{printer}」的默认设置（错误码 {size}）")
        buf = ctypes.create_string_buffer(size)
        dev = ctypes.cast(buf, ctypes.POINTER(DEVMODEW))
        if a["spool"].DocumentPropertiesW(None, hpr, printer, dev, None, DM_OUT_BUFFER) < 0:
            raise PrintError(f"取打印机默认设置失败「{printer}」")

        d = dev.contents
        d.dmFields |= (DM_ORIENTATION | DM_PAPERSIZE | DM_COPIES | DM_COLOR
                       | DM_DUPLEX | DM_COLLATE)
        d.dmOrientation = DMORIENT_PORTRAIT     # 纸张一律纵向，内容横着由自动旋转处理
        d.dmPaperSize = PAPER_ID.get(paper, 9)
        d.dmCopies = max(1, min(int(copies), 999))
        d.dmCollate = DMCOLLATE_TRUE if d.dmCopies > 1 else 0
        d.dmColor = DMCOLOR_COLOR if color else DMCOLOR_MONOCHROME
        d.dmDuplex = DUPLEX_ID.get(duplex, 1)

        # 交给驱动校验/合并：不认识的纸型、超范围的值会在这里被纠正，
        # 比自己拼 DEVMODE 稳（SumatraPDF 也是这么干的）
        if a["spool"].DocumentPropertiesW(None, hpr, printer, dev, dev,
                                          DM_IN_BUFFER | DM_OUT_BUFFER) < 0:
            raise PrintError(f"打印机「{printer}」不接受这套参数（{paper}/"
                             f"{duplex}/{'彩色' if color else '黑白'}）")
        return hpr, buf, dev
    except Exception:
        a["spool"].ClosePrinter(hpr)
        raise


def printer_caps(printer):
    """读一下可打印区域、分辨率、纸向。probe 和正式打印都用它。"""
    a = api()
    hpr, buf, dev = _open_devmode(printer, "A4", "single", 1, False)
    try:
        hdc = a["gdi32"].CreateDCW("WINSPOOL", printer, None, dev)
        if not hdc:
            raise PrintError(f"CreateDC 失败「{printer}」（错误码 {ctypes.get_last_error()}）")
        try:
            caps = {k: a["gdi32"].GetDeviceCaps(hdc, i) for k, i in
                    (("w", HORZRES), ("h", VERTRES),
                     ("dpi_x", LOGPIXELSX), ("dpi_y", LOGPIXELSY),
                     ("phys_w", PHYSICALWIDTH), ("phys_h", PHYSICALHEIGHT))}
            caps["printer"] = printer
            caps["paper_mm"] = (round(caps["phys_w"] / caps["dpi_x"] * MM_PER_INCH),
                                round(caps["phys_h"] / caps["dpi_y"] * MM_PER_INCH))
            return caps
        finally:
            a["gdi32"].DeleteDC(hdc)
    finally:
        a["spool"].ClosePrinter(hpr)


# ---------------- 位图 ----------------

def _dib(pix, rot):
    """pymupdf 的 pixmap -> (buffer, 宽, 高)。顺便处理两件事：
    1. 需要时顺时针转 90°（横向内容打到纵向纸上）；
    2. DIB 每行必须 4 字节对齐，MuPDF 的 stride 是 width*n，不对齐直接喂 GDI
       会静默返回 0（一张空白页，spool 只有两三 KB），很难查。
    """
    n = pix.n
    a = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.stride)
    a = a[:, :pix.width * n].reshape(pix.height, pix.width, n)
    if rot:
        a = np.rot90(a, k=-1)               # np.rot90 是逆时针，k=-1 才是顺时针
    h, w = a.shape[:2]
    row = w * n
    pad = (-row) % 4
    if pad:
        out = np.zeros((h, row + pad), dtype=np.uint8)
        out[:, :row] = a.reshape(h, row)
        a = out
    else:
        a = a.reshape(h, row)
    return np.ascontiguousarray(a), w, h


def _blit(dib, hdc, dx, dy, dw, dh, sx=0, sy=0, sw=None, sh=None):
    """把位图的一块（源矩形 sx,sy,sw,sh）贴到 dx,dy 处，目标大小 dw×dh。

    DIB 头里写的是整张图（bw×bh），源矩形另给 —— 这样分带送图不用复制内存。
    """
    gdi32 = api()["gdi32"]
    buf, bw, bh = dib
    sw = bw if sw is None else sw
    sh = bh if sh is None else sh
    bi = BITMAPINFO()
    bi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bi.bmiHeader.biWidth = bw
    bi.bmiHeader.biHeight = -bh             # 负数 = 自上而下，和位图内存顺序一致
    bi.bmiHeader.biPlanes = 1
    bi.bmiHeader.biBitCount = 24
    bi.bmiHeader.biCompression = 0          # BI_RGB
    bi.bmiHeader.biSizeImage = buf.nbytes
    ret = gdi32.StretchDIBits(hdc, dx, dy, dw, dh, sx, sy, sw, sh,
                              buf.ctypes.data_as(ctypes.c_void_p), ctypes.byref(bi),
                              0, SRCCOPY)
    if ret == 0:
        raise PrintError("位图送驱动被拒（StretchDIBits 返回 0）", code=3)


def plan_bands(x, y, dw, dh, left, top, right, bottom, bw, bh, band_rows=None):
    """算出分带方案：把位图（bw×bh）按 (x,y,dw,dh) 铺到纸上、只保留纸面内 (left,top,right,bottom)
    的那部分，切成长度不超过 band_rows 的横带。

    产出 (dx, dy, dw, dh, sx, sy, sw, sh)：目标矩形 + 源矩形，逐个交给 StretchDIBits。
    纯计算，不碰 GDI —— 自检可以直接验它有没有留缝、有没有重复。
    """
    band_rows = band_rows or BAND_ROWS
    out = []
    sx0 = max(0, min(bw - 1, int(round((left - x) * bw / dw))))
    sx1 = max(sx0 + 1, min(bw, int(round((right - x) * bw / dw))))
    sy = max(0, min(bh - 1, int(round((top - y) * bh / dh))))
    dy = top
    while dy < bottom:
        band = min(band_rows, bottom - dy)
        sy_next = min(bh, max(sy + 1, int(round((dy + band - y) * bh / dh))))
        out.append((left, dy, right - left, band, sx0, sy, sx1 - sx0, sy_next - sy))
        sy = sy_next
        dy += band
        if sy >= bh:
            break
    return out


def _blit_image(dib, hdc, x, y, dw, dh, paper_w, paper_h):
    """把整张位图贴到纸面的 (x,y) 上，尺寸 dw×dh，只画纸面以内的部分，并按行分带送驱动。

    两个作用：
    * 超出纸面的部分由我们自己先裁掉，而不是把几万像素的大矩形丢给驱动去裁
      （驱动裁切时静默丢数据是"内容不全"的已知来源）；
    * 每带只有 BAND_ROWS 行，A0/A1 这类大图也不会一次压垮驱动。
    返回 (可见宽, 可见高, 是否被裁)。
    """
    buf, w, h = dib
    left, top = max(0, x), max(0, y)
    right, bottom = min(paper_w, x + dw), min(paper_h, y + dh)
    if right <= left or bottom <= top:
        return 0, 0, True
    clipped = x < 0 or y < 0 or x + dw > paper_w or y + dh > paper_h
    for dx, dy, bw, bh, sx, sy, sw, sh in plan_bands(x, y, dw, dh, left, top, right, bottom, w, h):
        _blit(dib, hdc, dx, dy, bw, bh, sx, sy, sw, sh)
    return right - left, bottom - top, clipped



# ---------------- 主流程 ----------------

def print_pdf(path, printer, copies=1, paper="A4", duplex="single",
              scale="fit", color=False, page_spec="", probe=False, oversize="shrink"):
    """把一个 PDF 送到打印机。probe=True 只读设备能力，不建打印作业。"""
    if not available():
        raise PrintError("当前环境不支持 GDI 送印（缺 numpy 或不是 Windows）", code=5)
    a = api()
    mode = SCALE_ID.get(scale, "fit")

    if probe:
        return printer_caps(printer)

    doc = fitz.open(path)
    try:
        pages = parse_range(page_spec, doc.page_count)
        if not pages:
            raise PrintError(f"{os.path.basename(path)}：没有要打的页（页码超出范围或全被排除）", code=5)
        hpr, buf, dev = _open_devmode(printer, paper, duplex, copies, color)
        try:
            gdi32 = a["gdi32"]
            hdc = gdi32.CreateDCW("WINSPOOL", printer, None, dev)
            if not hdc:
                raise PrintError(f"CreateDC 失败「{printer}」（错误码 {ctypes.get_last_error()}）")
            try:
                return _run(doc, pages, hdc, printer, mode, color, path, oversize)
            finally:
                gdi32.DeleteDC(hdc)
        finally:
            a["spool"].ClosePrinter(hpr)
    finally:
        doc.close()


def _run(doc, pages, hdc, printer, mode, color, path, oversize="shrink"):
    gdi32 = api()["gdi32"]
    w = gdi32.GetDeviceCaps(hdc, HORZRES)          # 可打印区域（设备像素）
    h = gdi32.GetDeviceCaps(hdc, VERTRES)
    dpi = gdi32.GetDeviceCaps(hdc, LOGPIXELSX) or 600
    phys_w = gdi32.GetDeviceCaps(hdc, PHYSICALWIDTH)
    phys_h = gdi32.GetDeviceCaps(hdc, PHYSICALHEIGHT)
    if min(w, h, dpi) <= 0:
        raise PrintError(f"打印机「{printer}」没给出可打印区域")
    paper_landscape = phys_w > phys_h

    di = DOCINFOW(ctypes.sizeof(DOCINFOW), os.path.basename(path), None, None, 0)
    if gdi32.StartDocW(hdc, ctypes.byref(di)) <= 0:
        raise PrintError(f"StartDoc 失败「{printer}」（错误码 {ctypes.get_last_error()}）")

    sent = 0
    warns = []
    gdi32.SetMapMode(hdc, 1)                        # MM_TEXT：一个单位 = 一个设备像素
    # 单色机要让 GDI 自己去做抖动（否则灰阶层次全丢成纯黑纯白）；彩机逐色直传
    gdi32.SetStretchBltMode(hdc, COLORONCOLOR if color else HALFTONE)
    if not color:
        gdi32.SetBrushOrgEx(hdc, 0, 0, None)
    try:
        for no in pages:
            page = doc[no - 1]
            pw, ph = page.rect.width / 72.0, page.rect.height / 72.0      # 英寸
            if pw <= 0 or ph <= 0:
                continue
            rot = 90 if (pw > ph) != paper_landscape else 0
            cw, ch = (ph, pw) if rot else (pw, ph)                        # 旋转后的英寸尺寸
            fit = min(w / (cw * dpi), h / (ch * dpi))
            if mode == "fit":
                k = fit
            elif mode == "shrink":
                k = min(1.0, fit)                  # 只缩不放
            else:
                k = 1.0                            # noscale / custom（比例已烘进 PDF）

            # 纸放不下：1:1 打 1261mm 的图纸到 A3，硬打只能出左上角一张 —— 就是
            # "内容不全"。除非明确要求 clip，否则自动缩到能放下，保证整张图出全。
            k_orig = k
            over = max(cw * dpi * k / w, ch * dpi * k / h)
            shrunk = False
            if over > OVER_TOL and oversize == "shrink":
                k = fit if fit < k else k
                shrunk = True
                warns.append(f"图纸 {cw * 25.4:.0f}×{ch * 25.4:.0f}mm 装不进纸张 "
                             f"{w / dpi * 25.4:.0f}×{h / dpi * 25.4:.0f}mm，已按 "
                             f"{k * 100:.0f}% 缩放打全（所选比例 {k_orig * 100:.0f}%）")

            px_w, px_h = cw * dpi * k, ch * dpi * k
            rk = k
            if px_w * px_h > MAX_PIXELS:           # 太大会吃爆内存，降分辨率再拉伸
                rk = k * (MAX_PIXELS / (px_w * px_h)) ** 0.5
            scale = dpi * rk / 72.0
            if scale <= 0:
                continue
            if gdi32.StartPage(hdc) <= 0:
                raise PrintError("StartPage 失败（打印队列可能被暂停）", code=3)
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale),
                                  colorspace=fitz.csRGB, alpha=False)
            dib = _dib(pix, rot)
            pix = None
            dw, dh = max(1, int(round(px_w))), max(1, int(round(px_h)))
            x, y = (w - dw) // 2, (h - dh) // 2            # 居中；超出部分由 _blit_image 裁掉
            _, _, clipped = _blit_image(dib, hdc, x, y, dw, dh, w, h)
            dib = None
            if clipped and not shrunk:
                warns.append(f"图纸 {cw * 25.4:.0f}×{ch * 25.4:.0f}mm 大于纸张"
                             f"{w / dpi * 25.4:.0f}×{h / dpi * 25.4:.0f}mm，"
                             f"按实际大小只打出中间一块；改用「适合」可打全")
            if gdi32.EndPage(hdc) <= 0:
                raise PrintError("EndPage 失败（打印机报错或脱机）", code=3)
            sent += 1
        if sent == 0:
            raise PrintError("没有可打印的页面", code=3)
        if gdi32.EndDoc(hdc) <= 0:
            raise PrintError("EndDoc 失败（驱动拒绝收尾）", code=3)
    except Exception:
        gdi32.AbortDoc(hdc)
        raise
    out = {"pages": len(pages), "printed": sent}
    if warns:
        out["warn"] = "；".join(dict.fromkeys(warns))      # 去重，多页重复的不刷屏
    return out


# ---------------- 命令行 ----------------

def main(argv):
    import argparse
    import json

    ap = argparse.ArgumentParser(description="GDI 直接送印 PDF")
    ap.add_argument("pdf", nargs="?", help="要打的 PDF")
    ap.add_argument("printer", nargs="?", help="打印机名（要和 Windows 里一模一样）")
    ap.add_argument("--copies", type=int, default=1)
    ap.add_argument("--paper", default="A4")
    ap.add_argument("--duplex", default="single", choices=list(DUPLEX_ID))
    ap.add_argument("--scale", default="fit", choices=list(SCALE_ID))
    ap.add_argument("--color", type=int, default=0)
    ap.add_argument("--range", dest="page_spec", default="")
    ap.add_argument("--oversize", default="shrink", choices=["shrink", "clip"],
                    help="图纸比纸大时的处理：shrink=自动缩到放得下（默认），clip=按原大小裁切")
    ap.add_argument("--probe", action="store_true", help="只读设备能力，不出纸")
    args = ap.parse_args(argv)

    try:
        if args.probe:
            # 只传一个名字时它落在 pdf 位上（python gdi_print.py --probe <打印机>）
            caps = printer_caps(args.printer or args.pdf)
            sys.stdout.write(json.dumps(caps, ensure_ascii=False) + "\n")
            return 0
        if not args.pdf or not args.printer:
            raise PrintError("缺少 PDF 路径或打印机名", code=5)
        if not os.path.isfile(args.pdf):
            raise PrintError(f"文件不存在：{args.pdf}", code=5)
        r = print_pdf(args.pdf, args.printer, copies=args.copies, paper=args.paper,
                      duplex=args.duplex, scale=args.scale, color=bool(args.color),
                      page_spec=args.page_spec, oversize=args.oversize)
        sys.stdout.write(json.dumps(r, ensure_ascii=False) + "\n")
        return 0
    except PrintError as e:
        # 父进程按 UTF-8 读，别让中文在 cp936 控制台上炸掉
        sys.stderr.buffer.write(str(e).encode("utf-8", "replace"))
        sys.stderr.flush()
        return e.code
    except Exception as e:
        sys.stderr.buffer.write(f"{type(e).__name__}: {e}".encode("utf-8", "replace"))
        sys.stderr.flush()
        return 3


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
