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

**线稿模式**（`--lineart auto`，默认）：CAD 图纸那种整页 0.06pt 发丝线的 PDF，
600dpi 下一条线不到半个设备点，抗锯齿把它摊成 40% 左右的灰，单色机再一抖动就成了
断断续续的点线——所以这类页渲染时关掉抗锯齿，线直接钉成 1 个设备点的实黑，和
Acrobat 送矢量打出来的一样清楚。判定条件是"描边又多又细 + 页面上没有真位图"，
照片页、普通文档不受影响（见 `is_lineart`）。

用法（父进程 print_system.py 就是这么调的）：
    python gdi_print.py <pdf> <打印机名> [--copies N] [--paper A4] [--duplex single]
                        [--scale fit] [--color 1] [--range 1-3,5] [--nup 4]
                        [--lineart auto|on|off]
    python gdi_print.py --probe <打印机名>      # 只读设备能力，不出纸
    python gdi_print.py x.pdf 打印机 --preview-png out.png   # 只存"要送印的位图"，不出纸

--nup N（每张纸印 N 页）在送印层直接排版：一张纸上 StartPage 一次、画 N 个格子，
不需要先把 N 页拼成一份新 PDF 再打。--range 也照旧有效——先按页号挑页，再按顺序排格子。

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
except ImportError:                       # 缺 numpy 就没法送印，父进程会直接报错退出
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
# 每张纸印几页时的格子排法 (行, 列)。和 print_system.NUP_GRID 必须一致（自检会比对）。
NUP_GRID = {1: (1, 1), 2: (1, 2), 4: (2, 2), 6: (2, 3), 9: (3, 3), 16: (4, 4)}

DM_ORIENTATION = 0x00000001
DM_PAPERSIZE = 0x00000002
DM_COPIES = 0x00000100
DM_COLOR = 0x00000800
DM_DUPLEX = 0x00001000
DM_COLLATE = 0x00008000
DM_IN_BUFFER = 0x00000008
DM_OUT_BUFFER = 0x00000002

DMORIENT_PORTRAIT = 1
DMORIENT_LANDSCAPE = 2
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

# ---- 线稿页判定（见 is_lineart）----
LINEART_MIN_DRAWINGS = 300       # 描边条数下限：CAD 立面/平面动辄上千条，Word 讲义只有几十条
LINEART_HAIRLINE_PT = 0.25       # 描边中位线宽上限：细到 0.25pt 以下才是"发丝线"，才值得关 AA
LINEART_MIN_IMG_PT = 100         # 一边小于 100pt 的图（logo、线型样本）不算照片
LINEART_IMG_AREA = 0.2           # 面积不到整页 20% 的图也不算照片（线稿图上贴张小图很常见）


def _emit(line):
    """给父进程报进度。父进程按行读 stdout，用前缀区分：
        P <已送张数> <总张数>   每结束一张纸报一次
        JOB <id>                spooler 里的作业号（取消时要靠它删作业）
    其余不带前缀的那一行是最终结果 JSON。
    必须 flush：走管道时 stdout 是块缓冲，不 flush 父进程要等到进程结束才看得见。
    """
    try:
        sys.stdout.write(line + "\n")
        sys.stdout.flush()
    except Exception:
        pass


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

def _open_devmode(printer, paper, duplex, copies, color, landscape=False):
    """打开打印机，取出默认 DEVMODE，把我们关心的字段改掉，再让驱动校验一遍。

    landscape 只给多页缩印用：格子列比行多（如 6 页 2x3）时把纸横过来，
    内容才不用转 90°。单页打印一律纵向，横向内容由自动旋转处理。
    """
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
        d.dmOrientation = DMORIENT_LANDSCAPE if landscape else DMORIENT_PORTRAIT
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
    """pymupdf 的 pixmap -> (buffer, 宽, 高)。顺便处理三件事：
    1. R/B 换序：MuPDF 的 samples 是 R,G,B，而 24bpp BI_RGB 的 DIB 内存顺序是
       B,G,R。不换会把红章打成蓝章、蓝字打成红字（灰度/单通道不受影响）；
    2. 需要时顺时针转 90°（横向内容打到纵向纸上）；
    3. DIB 每行必须 4 字节对齐，MuPDF 的 stride 是 width*n，不对齐直接喂 GDI
       会静默返回 0（一张空白页，spool 只有两三 KB），很难查。
    """
    n = pix.n
    a = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.stride)
    a = a[:, :pix.width * n].reshape(pix.height, pix.width, n)
    if n == 3:
        a = np.ascontiguousarray(a[..., ::-1])
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


def _blit(dib, hdc, dw, dh, x, y):
    """把 (buffer, 宽, 高) 这一张位图贴到 x,y，目标大小 dw×dh（1:1 时就是原尺寸）。"""
    gdi32 = api()["gdi32"]
    buf, w, h = dib
    bi = BITMAPINFO()
    bi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bi.bmiHeader.biWidth = w
    bi.bmiHeader.biHeight = -h              # 负数 = 自上而下，和位图内存顺序一致
    bi.bmiHeader.biPlanes = 1
    bi.bmiHeader.biBitCount = 24
    bi.bmiHeader.biCompression = 0          # BI_RGB
    bi.bmiHeader.biSizeImage = buf.nbytes
    ret = gdi32.StretchDIBits(hdc, x, y, dw, dh, 0, 0, w, h,
                              buf.ctypes.data_as(ctypes.c_void_p), ctypes.byref(bi),
                              0, SRCCOPY)
    if ret == 0:
        raise PrintError("位图送驱动被拒（StretchDIBits 返回 0）", code=3)


# ---------------- 线稿页 ----------------

def is_lineart(page):
    """判断这一页是不是纯矢量线稿（CAD 图纸、机加工图、图表）。

    为什么要单独认这一类页：图纸的线常常只有 0.06pt（0.021mm），600dpi 一个设备点是
    0.0423mm，一条线在纸上还不到半个设备点。抗锯齿渲染会把它摊成覆盖率四五十的灰，
    单色机再用 HALFTONE 一抖动，就成了断断续续的点线——看着就是"线条看不清"。
    关掉抗锯齿后 MuPDF 直接把这种线钉成 1 个设备点的实黑，和 Acrobat 送矢量、
    驱动自己做"细线增强"的结果一致。

    判定三条：描边够多（≥300 条）、描边中位线宽 ≤0.25pt、页面上没有成块的位图。
    先用 get_bboxlog() 数描边条数（比 get_drawings 快三五倍），条数不够就直接否掉，
    普通讲义/Word 文档（描边才几条）连解析路径这一步都省了。线型样本那种 1 像素高的
    小图、模板蒙版（fill-imgmask）、标题栏里的小 logo 都不算"照片"，不然图纸一律会被
    当成图片页。真照片/扫描件要"够大"才算：一边 ≥100pt 且面积占整页 ≥20%。
    """
    try:
        page_area = abs(page.rect.width * page.rect.height) or 1.0
        strokes = 0
        for op, r in page.get_bboxlog():
            if op == "fill-image":
                iw, ih = r[2] - r[0], r[3] - r[1]
                if (iw >= LINEART_MIN_IMG_PT and ih >= LINEART_MIN_IMG_PT
                        and iw * ih >= page_area * LINEART_IMG_AREA):
                    return False                # 成块的位图（照片、扫描件、彩色贴图）
            elif op == "stroke-path":
                strokes += 1
        if strokes < LINEART_MIN_DRAWINGS:
            return False
        widths = [d["width"] for d in page.get_drawings() if d.get("width")]
        if not widths or float(np.median(widths)) > LINEART_HAIRLINE_PT:
            return False
        return True
    except Exception:
        return False                            # 分析失败就当普通页，宁可不优化也别打错


def _pixmap(page, scale, lineart):
    """按 scale 渲染一页。lineart=True 时临时关掉抗锯齿（进程级开关，用完还原）。"""
    m = fitz.Matrix(scale, scale)
    if not lineart:
        return page.get_pixmap(matrix=m, colorspace=fitz.csRGB, alpha=False)
    old = fitz.TOOLS.show_aa_level().get("graphics", 8)
    try:
        fitz.TOOLS.set_aa_level(0)
        return page.get_pixmap(matrix=m, colorspace=fitz.csRGB, alpha=False)
    finally:
        fitz.TOOLS.set_aa_level(old)


# ---------------- 主流程 ----------------

def print_pdf(path, printer, copies=1, paper="A4", duplex="single",
              scale="fit", color=False, page_spec="", probe=False, nup=1,
              lineart="auto", preview_png=None):
    """把一个 PDF 送到打印机。probe=True 只读设备能力，不建打印作业。

    nup>1 时每张纸印 nup 页：在送印层按格子排版，不需要先拼一份 PDF 出来。
    lineart：auto（默认，按页判定线稿）/ on / off，决定要不要关抗锯齿。
    preview_png：只把第一页"准备送印的那张位图"存成 PNG，不出纸（调试用）。
    """
    if not available():
        raise PrintError("当前环境不支持 GDI 送印（缺 numpy 或不是 Windows）", code=5)
    a = api()
    mode = SCALE_ID.get(scale, "fit")
    nup = 1 if nup in (None, "") else int(nup)
    lineart = lineart if lineart in ("auto", "on", "off") else "auto"
    if nup not in NUP_GRID:
        raise PrintError(f"每页张数只能是 {'/'.join(map(str, NUP_GRID))}", code=5)

    if probe:
        return printer_caps(printer)

    doc = fitz.open(path)
    try:
        pages = parse_range(page_spec, doc.page_count)
        if not pages:
            raise PrintError(f"{os.path.basename(path)}：没有要打的页（页码超出范围或全被排除）", code=5)
        rows, cols = NUP_GRID[nup]
        hpr, buf, dev = _open_devmode(printer, paper, duplex, copies, color,
                                      landscape=nup > 1 and cols > rows)
        try:
            gdi32 = a["gdi32"]
            hdc = gdi32.CreateDCW("WINSPOOL", printer, None, dev)
            if not hdc:
                raise PrintError(f"CreateDC 失败「{printer}」（错误码 {ctypes.get_last_error()}）")
            try:
                return _run(doc, pages, hdc, printer, mode, color, path, nup,
                            lineart_mode=lineart, preview_png=preview_png)
            finally:
                gdi32.DeleteDC(hdc)
        finally:
            a["spool"].ClosePrinter(hpr)
    finally:
        doc.close()


def _render_page(page, geo, mode, nup, j, cols, lineart):
    """算一页的位置和缩放，渲染成 DIB。返回 (pix, dib, dw, dh, x, y)，空页返回 None。

    geo = (w, h, dpi, paper_landscape, cell_w, cell_h)，就是可打印区域和格子尺寸。
    """
    w, h, dpi, paper_landscape, cell_w, cell_h = geo
    pw, ph = page.rect.width / 72.0, page.rect.height / 72.0      # 英寸
    if pw <= 0 or ph <= 0:
        return None
    if nup > 1:
        # 拼版格子：不旋转（纸已按网格转过来了），每页独立等比缩进自己那格。
        # 缩放档对格子没有意义——塞不进格子就白打，所以一律按 fit。
        rot, cw, ch = 0, pw, ph
        k = min(cell_w / (cw * dpi), cell_h / (ch * dpi))
        r_, c_ = divmod(j, cols)
    else:
        rot = 90 if (pw > ph) != paper_landscape else 0
        cw, ch = (ph, pw) if rot else (pw, ph)                    # 旋转后的英寸尺寸
        fit = min(w / (cw * dpi), h / (ch * dpi))
        if mode == "fit":
            k = fit
        elif mode == "shrink":
            k = min(1.0, fit)                  # 只缩不放
        else:
            k = 1.0                            # noscale / custom（比例已烘进 PDF）
        r_, c_ = 0, 0
    px_w, px_h = cw * dpi * k, ch * dpi * k
    rk = k
    if px_w * px_h > MAX_PIXELS:           # 太大会吃爆内存，降分辨率再拉伸
        rk = k * (MAX_PIXELS / (px_w * px_h)) ** 0.5
    scale = dpi * rk / 72.0
    if scale <= 0:
        return None
    pix = _pixmap(page, scale, lineart)
    dib = _dib(pix, rot)
    dw, dh = max(1, int(round(px_w))), max(1, int(round(px_h)))
    if nup > 1:                            # 摆进自己那一格，格内居中
        x = max(0, int(round(c_ * cell_w + (cell_w - dw) / 2)))
        y = max(0, int(round(r_ * cell_h + (cell_h - dh) / 2)))
    else:
        x, y = max(0, (w - dw) // 2), max(0, (h - dh) // 2)        # 居中
    return pix, dib, dw, dh, x, y


def _lineart_of(page, lineart_mode):
    """--lineart auto 按页面自动判；on/off 是给调试和"这张图我一定要这么打"留的手动挡。"""
    if lineart_mode == "on":
        return True
    if lineart_mode == "off":
        return False
    return is_lineart(page)


def _run(doc, pages, hdc, printer, mode, color, path, nup=1,
         lineart_mode="auto", preview_png=None):
    gdi32 = api()["gdi32"]
    w = gdi32.GetDeviceCaps(hdc, HORZRES)          # 可打印区域（设备像素）
    h = gdi32.GetDeviceCaps(hdc, VERTRES)
    dpi = gdi32.GetDeviceCaps(hdc, LOGPIXELSX) or 600
    phys_w = gdi32.GetDeviceCaps(hdc, PHYSICALWIDTH)
    phys_h = gdi32.GetDeviceCaps(hdc, PHYSICALHEIGHT)
    if min(w, h, dpi) <= 0:
        raise PrintError(f"打印机「{printer}」没给出可打印区域")
    paper_landscape = phys_w > phys_h

    # 多页缩印：把可打印区域切成 rows×cols 个格子。nup=1 时整张纸就是一格。
    rows, cols = NUP_GRID.get(nup, (1, 1))
    cell_w, cell_h = w / cols, h / rows
    geo = (w, h, dpi, paper_landscape, cell_w, cell_h)

    if preview_png:
        # 调试用：只把第一页渲染成"准备送印的那张位图"存盘，不建打印作业、不出纸
        page = doc[pages[0] - 1]
        res = _render_page(page, geo, mode, nup, 0, cols, _lineart_of(page, lineart_mode))
        if not res:
            raise PrintError("这一页没有可打印的内容", code=5)
        res[0].save(preview_png)
        return {"pages": 0, "sheets": 0, "printed": 0, "preview": preview_png}

    di = DOCINFOW(ctypes.sizeof(DOCINFOW), os.path.basename(path), None, None, 0)
    # StartDoc 的返回值就是 spooler 里的作业号：父进程取消打印时用它 SetJob 删作业，
    # 不然只能杀进程，作业会留在队列里接着出纸
    jid = gdi32.StartDocW(hdc, ctypes.byref(di))
    if jid <= 0:
        raise PrintError(f"StartDoc 失败「{printer}」（错误码 {ctypes.get_last_error()}）")
    total_sheets = (len(pages) + nup - 1) // nup
    _emit(f"JOB {jid}")
    _emit(f"P 0 {total_sheets}")

    sent = 0
    gdi32.SetMapMode(hdc, 1)                        # MM_TEXT：一个单位 = 一个设备像素
    # 单色机要让 GDI 自己去做抖动（否则灰阶层次全丢成纯黑纯白）；彩机逐色直传。
    # 线稿页渲染出来本来就是黑白两色，抖动对它没有副作用，不用单独改模式。
    gdi32.SetStretchBltMode(hdc, COLORONCOLOR if color else HALFTONE)
    if not color:
        gdi32.SetBrushOrgEx(hdc, 0, 0, None)
    try:
        for i in range(0, len(pages), nup):
            if gdi32.StartPage(hdc) <= 0:
                raise PrintError("StartPage 失败（打印队列可能被暂停）", code=3)
            for j, no in enumerate(pages[i:i + nup]):
                page = doc[no - 1]
                res = _render_page(page, geo, mode, nup, j, cols,
                                   _lineart_of(page, lineart_mode))
                if res is None:
                    continue
                pix, dib, dw, dh, x, y = res
                pix = None                          # 位图已经搬进 dib，别占着内存
                _blit(dib, hdc, dw, dh, x, y)
                dib = None
            if gdi32.EndPage(hdc) <= 0:
                raise PrintError("EndPage 失败（打印机报错、脱机或被取消）", code=3)
            sent += 1
            _emit(f"P {sent} {total_sheets}")       # 这一张已经交给驱动了
        if sent == 0:
            raise PrintError("没有可打印的页面", code=3)
        if gdi32.EndDoc(hdc) <= 0:
            raise PrintError("EndDoc 失败（驱动拒绝收尾）", code=3)
    except Exception:
        gdi32.AbortDoc(hdc)
        raise
    return {"pages": len(pages), "sheets": sent, "printed": sent}


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
    ap.add_argument("--nup", type=int, default=1, choices=sorted(NUP_GRID),
                    help="每张纸印几页（默认 1）")
    ap.add_argument("--lineart", default="auto", choices=["auto", "on", "off"],
                    help="线稿页关抗锯齿：auto 按页判定（默认），on 强制，off 关闭")
    ap.add_argument("--preview-png", metavar="PNG", default=None,
                    help="只把第一页准备送印的位图存成 PNG，不出纸（调试用）")
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
                      page_spec=args.page_spec, nup=args.nup, lineart=args.lineart,
                      preview_png=args.preview_png)
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
