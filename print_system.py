# -*- coding: utf-8 -*-
"""LinkPrint —— 热文件夹 PDF 打印系统（预检 + 批量下单）

启动：启动打印系统.bat 或 python print_system.py
流程：PDF 丢进「热文件夹」→ 网页里逐份调参数 → 一次性发出打印。
"""
import csv
import glob
import json
import os
import queue
import re
import socket
import subprocess
import sys
import threading
import time
import webbrowser
import zlib
from collections import Counter
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import parse_qs, unquote, urlparse

try:
    import pymupdf as fitz
except ImportError:
    import fitz

BASE = os.path.dirname(os.path.abspath(__file__))
CFG_PATH = os.path.join(BASE, "config.json")
WEB_DIR = os.path.join(BASE, "web")

DEFAULT_CFG = {
    "port": 8777,
    "hotfolder": "热文件夹",
    # 按关键字匹配本机打印机名，改名换关键词即可
    "printers": ["D95", "D125", "C7580"],
    "papers": ["A4", "A3"],
    "sumatra": "SumatraPDF.exe",
    # 送印引擎：gdi = 自己走 GDI（快，实测 7 页从 54s 降到 2s 量级）；sumatra = 老办法
    "engine": "gdi",
    # gdi 引擎在「还没往打印机送任何数据」的失败下，自动换 sumatra 再试一次
    "engine_fallback": True,
    # 不同打印机之间并行送印（同一台机器仍然严格按顺序，不会乱）
    "print_parallel": True,
    # Office -> PDF 转换器（WPS COM）。找不到它就只支持 PDF
    "convert2pdf": r"D:\脚本文件\多文件合并PDF\Convert2PDF.ps1",
    # 打印效果预览里，纸张四边扣掉多少毫米当"不可打印区"（估算）。
    # 驱动真实边距拿不到，4mm 大致对应 A4 满幅时 Acrobat 显示的 96%。
    "preview_margin_mm": 4,
}

# 标准纸张 mm（宽 x 高），容忍 ±2mm
STD_SIZES = {
    "A0": (841, 1189), "A1": (594, 841), "A2": (420, 594), "A3": (297, 420),
    "A4": (210, 297), "A5": (148, 210), "A6": (105, 148),
    "B4": (250, 353), "B5": (176, 250), "B5-JIS": (182, 257),
    "Letter": (216, 279), "Legal": (216, 356), "11x17": (279, 432),
    "SRA3": (320, 450),
}


def load_cfg():
    if not os.path.exists(CFG_PATH):
        with open(CFG_PATH, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CFG, f, ensure_ascii=False, indent=2)
        return dict(DEFAULT_CFG)
    with open(CFG_PATH, encoding="utf-8") as f:
        cfg = dict(DEFAULT_CFG)
        cfg.update(json.load(f))
    return cfg


CFG = load_cfg()
HOT = CFG["hotfolder"]
if not os.path.isabs(HOT):
    HOT = os.path.join(BASE, HOT)
SUMATRA = CFG["sumatra"]
if not os.path.isabs(SUMATRA):
    SUMATRA = os.path.join(BASE, SUMATRA)

# Office 文件靠 Convert2PDF.ps1（WPS/Office COM）转 PDF，程序目录里有同名文件就优先用它
C2P = CFG["convert2pdf"]
if not os.path.isfile(C2P) and os.path.isfile(os.path.join(BASE, "Convert2PDF.ps1")):
    C2P = os.path.join(BASE, "Convert2PDF.ps1")
POWERSHELL = r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
CONV_DIR = os.path.join(BASE, "_转换缓存")

# 送印引擎。gdi_print.py 直接用 Windows GDI 把位图交给驱动：实测同一份 7 页讲义，
# SumatraPDF 要 53.9s（其中它自己 CPU 只用了 1.7s，基本在干等），走 GDI 只要 2.3s。
try:
    import gdi_print
except Exception as _e:                      # 缺 numpy 之类，退回 SumatraPDF
    gdi_print = None
    print("[提示] gdi_print 不可用，送印改用 SumatraPDF：", _e)

ENGINE = str(CFG.get("engine", "gdi")).lower()
if ENGINE == "gdi" and (gdi_print is None or not gdi_print.available()):
    print("[提示] 当前环境不支持 GDI 送印，自动改用 SumatraPDF")
    ENGINE = "sumatra"


def resolve_printers(keywords):
    """把关键字解析成本机真实打印机名（含 emoji 前缀）。"""
    try:
        import win32print
        names = [p["pPrinterName"] for p in win32print.EnumPrinters(6, None, 4)]
    except Exception as e:
        print("[警告] 无法枚举打印机:", e)
        return []
    out = []
    for kw in keywords:
        hit = next((n for n in names if kw.lower() in n.lower()), None)
        if hit and hit not in out:
            out.append(hit)
        elif not hit:
            print(f"[警告] 找不到打印机关键字: {kw}")
    return out


PRINTERS = resolve_printers(CFG["printers"])


def classify(w, h):
    """mm 尺寸 -> 标准纸型名，不是标准尺寸返回 None。
    容差 max(2mm, 2.5%)：带出血的 A4 常见 213x303，仍按 A4 认。"""
    w, h = sorted((w, h))
    for name, (sw, sh) in STD_SIZES.items():
        if abs(w - sw) <= max(2, sw * 0.025) and abs(h - sh) <= max(2, sh * 0.025):
            return name
    return None


CACHE = {}  # path -> ((mtime, size), info)

# 后台转换的状态：WANT 是「已排进队列/正在转」的文件（值=文件指纹），
# FAIL 记着转失败的文件（值=(指纹, 原因)），避免每个刷新周期都去重试一遍。
CONV_WANT = {}
CONV_FAIL = {}
CONV_STATE_LOCK = threading.Lock()
CONV_Q = queue.Queue()

# ---------------- Office -> PDF ----------------
# 分类名同时用作界面上的小标签，扩展名集合就是「能被打印」的 Office 格式
KIND = {}
for _k, _e in (("Word", ".doc .docx .docm .dot .dotx .dotm .wps .wpt"),
               ("表格", ".xls .xlsx .xlsm .xlsb .et .ett .csv"),
               ("PPT", ".ppt .pptx .pptm .pps .ppsx .ppsm .pot .potx .potm .dps .dpt")):
    for _x in _e.split():
        KIND[_x] = _k
OFFICE_EXT = set(KIND)
# 图片没有 Office 那套转换器，走 PyMuPDF 自己转（见 img_to_pdf）
IMG_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp", ".gif"}
CONV_LOCK = threading.Lock()

# 页面缩放：前三档是 SumatraPDF 原生参数；custom 没有原生支持，
# 先用 PyMuPDF 把页面按比例重排成新 PDF，再按 noscale 打
SCALE = {"fit": "fit", "shrink": "shrink", "noscale": "noscale", "custom": "noscale"}
SCALE_LOCK = threading.Lock()


def ext_of(p):
    return os.path.splitext(p)[1].lower()


def conv_target(src):
    """缓存 PDF 的路径。名字带 mtime+size：源文件一改，缓存自动作废。"""
    st = os.stat(src)
    stem = os.path.splitext(os.path.basename(src))[0]
    return os.path.join(CONV_DIR, f"{stem}_{st.st_size}_{int(st.st_mtime)}.pdf")


def cached_pdf(src):
    t = conv_target(src)
    return t if os.path.isfile(t) else None


def run_convert(paths):
    """调 Convert2PDF.ps1 转一批文件，返回 {源文件: 生成的 PDF}。"""
    env = dict(os.environ, CONVERT2PDF_OUTPUT_DIR=CONV_DIR, CONVERT2PDF_NOOPEN="1")
    subprocess.run([POWERSHELL, "-NoProfile", "-NonInteractive", "-ExecutionPolicy",
                    "Bypass", "-File", C2P] + list(paths),
                   env=env, timeout=900, capture_output=True, creationflags=0x08000000)
    made = {}
    map_file = os.path.join(CONV_DIR, "_mapping.tsv")   # 脚本每跑一次重写这份对照表
    if os.path.isfile(map_file):
        with open(map_file, encoding="utf-8-sig", errors="replace") as f:
            for line in f:
                src, _, pdf = line.rstrip("\r\n").partition("\t")
                if src and pdf:
                    made[os.path.normcase(os.path.abspath(src))] = pdf
    return {p: made.get(os.path.normcase(os.path.abspath(p))) for p in paths}


def ensure_pdfs(paths):
    """把还没转过的 Office 文件批量转一次（一次 PS 启动搞定一批，别一个文件起一次）。"""
    todo = [p for p in paths if ext_of(p) in OFFICE_EXT and not cached_pdf(p)]
    if not todo:
        return
    with CONV_LOCK:
        todo = [p for p in todo if not cached_pdf(p)]      # 排队期间别人可能已经转好
        if not todo:
            return
        os.makedirs(CONV_DIR, exist_ok=True)
        for src, pdf in run_convert(todo).items():
            if pdf and os.path.isfile(pdf):
                try:
                    os.replace(pdf, conv_target(src))       # 改名成带版本号的缓存名
                except OSError:
                    pass


def img_to_pdf(src):
    """图片 -> 缓存 PDF。页面按图片自身 DPI 折算大小，打印时再由 fit 缩到纸上。"""
    t = conv_target(src)
    if os.path.isfile(t):
        return t
    doc = fitz.open(src)                 # fitz 原生支持 jpg/png/bmp/tif/webp/gif
    try:
        data = doc.convert_to_pdf()      # 多页 TIFF 会转成多页
    finally:
        doc.close()
    os.makedirs(CONV_DIR, exist_ok=True)
    part = t + ".part"                   # 先写临时名，半截文件不会被当成有效缓存
    with open(part, "wb") as f:
        f.write(data)
    os.replace(part, t)
    return t


def printable_pdf(src):
    """Office 文件返回它的缓存 PDF（没有就现转一份），图片转成 PDF，PDF 原样返回。"""
    if ext_of(src) in IMG_EXT:
        return img_to_pdf(src)
    if ext_of(src) not in OFFICE_EXT:
        return src
    hit = cached_pdf(src)
    if not hit:
        ensure_pdfs([src])
        hit = cached_pdf(src)
    if not hit:
        raise RuntimeError("转换失败（WPS 没转出 PDF：文件损坏/加密，或 WPS 正弹着对话框）")
    return hit


def needs_convert(path):
    """这份文件要送去打印前，还得先转成 PDF 吗？"""
    return (ext_of(path) in OFFICE_EXT or ext_of(path) in IMG_EXT) and not cached_pdf(path)


def request_convert(path):
    """把文件排进后台转换队列。已经在队列里/正在转的不会重复排。"""
    try:
        st = os.stat(path)
    except OSError:
        return
    key = (st.st_mtime, st.st_size)
    with CONV_STATE_LOCK:
        if CONV_WANT.get(path) == key:
            return
        CONV_WANT[path] = key
    CONV_Q.put(path)


def scaled_target(src, pct):
    """自定义比例的缓存路径。带上源文件的 mtime+size，源一改缓存自动作废。"""
    st = os.stat(src)
    stem = os.path.splitext(os.path.basename(src))[0]
    return os.path.join(CONV_DIR, f"{stem}_{st.st_size}_{int(st.st_mtime)}_s{int(pct)}.pdf")


def _slow_rescale(doc, k):
    """重建页面：show_pdf_page 只搬内容流是矢量缩放，但每页都要按 zlib 最高级别
    重压一遍内容流（1.5MB 内容 ≈ 220ms/页），慢在这一步。带 /Rotate 的页面走这条路最稳。"""
    out = fitz.open()
    try:
        for pg in doc:
            np_ = out.new_page(width=pg.rect.width * k, height=pg.rect.height * k)
            np_.show_pdf_page(np_.rect, doc, pg.number)
        return out.tobytes(deflate=True, garbage=3)
    finally:
        out.close()


def _fast_rescale(doc, k):
    """原地拷页，只在内容流最外层套一句 q..cm..Q 做缩放。

    等价于 show_pdf_page 的矢量缩放（不重采样、画质一样），但省掉了重新压缩内容流：
    自己按 zlib 低级别打包，实测 4 页文档 0.90s -> 0.13s，文件大约大 18%。
    /Rotate 的页面不用这条路（旋转下裁剪框语义不好伺候），调用方会退回 _slow_rescale。
    """
    out = fitz.open()
    try:
        out.insert_pdf(doc)
        for pg in out:
            rc = pg.cropbox
            xrefs = pg.get_contents()
            if xrefs:
                raw = b"\n".join(out.xref_stream(x) for x in xrefs)
                # 缩放 + 把裁剪框左下角挪到原点（源页裁剪框可能不在 0,0）
                head = b"q %.6f 0 0 %.6f %.6f %.6f cm\n" % (k, k, -rc.x0 * k, -rc.y0 * k)
                data = zlib.compress(head + raw + b"\nQ\n", 1)
                out.update_stream(xrefs[0], data, compress=False)
                out.xref_set_key(xrefs[0], "Filter", "/FlateDecode")
                out.xref_set_key(xrefs[0], "Length", str(len(data)))
                for x in xrefs[1:]:          # 多段内容流合并成一段了，其余清空
                    out.update_stream(x, b"", compress=False)
                    out.xref_set_key(x, "Filter", "/FlateDecode")
            # 不先删掉 CropBox，MediaBox 就改不动（MuPDF 要求 CropBox 必须落在 MediaBox 内）
            out.xref_set_key(pg.xref, "CropBox", "null")
            out.reload_page(pg).set_mediabox(fitz.Rect(0, 0, rc.width * k, rc.height * k))
        return out.tobytes(deflate=False, garbage=1)
    finally:
        out.close()


def scaled_pdf(src, pct):
    """把 src 的可打印 PDF 按 pct% 精确重排成一份新 PDF，返回缓存路径。
    重排出来的页面本身已经是 pct% 大小，所以送印时配 noscale。"""
    pct = int(pct)
    if not 10 <= pct <= 400:
        raise ValueError(f"自定义比例需在 10-400%：{pct}")
    t = scaled_target(src, pct)
    if os.path.isfile(t) and os.path.getsize(t):
        return t
    with SCALE_LOCK:
        if os.path.isfile(t) and os.path.getsize(t):   # 排队期间别人可能已经做好
            return t
        doc = fitz.open(printable_pdf(src))
        try:
            if doc.needs_pass:
                raise RuntimeError("PDF 有密码，无法按比例缩放")
            if doc.page_count > 2000:
                raise RuntimeError(f"页数太多（{doc.page_count}），请改用「适合」或「实际大小」")
            n = doc.page_count
            try:
                # show_pdf_page 只搬页面内容流，不拍平会丢掉注释和表单域；
                # 原地缩放更要把注释先烘进内容，否则注释停在原来的位置和大小上
                doc.bake(annots=True, widgets=True)
            except Exception:
                pass
            k = pct / 100.0
            data = None
            if not any(pg.rotation for pg in doc):
                try:
                    data = _fast_rescale(doc, k)
                except Exception as e:      # 各种畸形页面都不该让打不出来
                    print("[提示] 快速缩放失败，改用逐页重建:", e)
            if data is None:
                data = _slow_rescale(doc, k)
        finally:
            doc.close()
        os.makedirs(CONV_DIR, exist_ok=True)
        part = t + ".part"          # 半截文件不会被当成有效缓存
        with open(part, "wb") as f:
            f.write(data)
        try:
            os.replace(part, t)
        except OSError:             # Sumatra 正读着同名文件时会拒绝替换，已有就复用
            if not os.path.isfile(t):
                raise
        finally:
            if os.path.exists(part):
                try:
                    os.remove(part)
                except OSError:
                    pass
        # 校验：页数不对就宁可报错，别拿一份缺页的 PDF 去打印
        try:
            with fitz.open(t) as chk:
                bad = chk.page_count != n
        except Exception:
            bad = True
        if bad:
            try:
                os.remove(t)
            except OSError:
                pass
            raise RuntimeError("缩放后的 PDF 校验不通过（页数不符），请改用「适合」打印")
    return t


# ---------------- 多页（N-up） ----------------
# 把 N 页缩印到一张纸上，和 Acrobat Reader 打印对话框的「每页张数」一个意思。
# SumatraPDF 命令行没有 N-up 参数，所以和自定义比例一样：先用 PyMuPDF 拼成
# 每张纸一份页面的新 PDF（缓存名带 _nup{N}），再按 noscale 送印。
NUP_GRID = {1: (1, 1), 2: (1, 2), 4: (2, 2), 6: (2, 3), 9: (3, 3), 16: (4, 4)}


def nup_target(src, nup):
    st = os.stat(src)
    stem = os.path.splitext(os.path.basename(src))[0]
    return os.path.join(CONV_DIR, f"{stem}_{st.st_size}_{int(st.st_mtime)}_nup{nup}.pdf")


def nup_pdf(src, nup, paper):
    """把 src 每 nup 页缩印到一张 paper 纸上，返回缓存 PDF 路径。

    纸张方向跟着网格走：列比行多就用横向纸（宽高对调），否则纵向。
    每页按原始宽高比缩放进格子、不旋转，整体居中；页数和金额都按纸张数算。
    """
    nup = int(nup)
    rows, cols = NUP_GRID.get(nup, (1, 1))
    if nup == 1:
        return printable_pdf(src)
    t = nup_target(src, nup)
    if os.path.isfile(t) and os.path.getsize(t):
        return t
    with SCALE_LOCK:
        if os.path.isfile(t) and os.path.getsize(t):
            return t
        doc = fitz.open(printable_pdf(src))
        try:
            if doc.needs_pass:
                raise RuntimeError("PDF 有密码，无法多页拼版")
            if doc.page_count > 2000:
                raise RuntimeError(f"页数太多（{doc.page_count}）")
            n = doc.page_count
            try:
                doc.bake(annots=True, widgets=True)
            except Exception:
                pass
            wmm, hmm = paper_mm(paper)
            if cols > rows:                       # 横向网格用横向纸
                pw, ph = max(wmm, hmm), min(wmm, hmm)
            else:
                pw, ph = min(wmm, hmm), max(wmm, hmm)
            pw, ph = pw * 72 / 25.4, ph * 72 / 25.4
            cell_w, cell_h = pw / cols, ph / rows
            out = fitz.open()
            try:
                for i in range(0, n, nup):
                    np_ = out.new_page(width=pw, height=ph)
                    for j in range(nup):
                        pno = i + j
                        if pno >= n:
                            break
                        pg = doc[pno]
                        # 每页独立缩放，塞进格子（不旋转，保持原方向）
                        k = min(cell_w / pg.rect.width, cell_h / pg.rect.height)
                        tw, th = pg.rect.width * k, pg.rect.height * k
                        r_, c_ = divmod(j, cols)
                        x0 = c_ * cell_w + (cell_w - tw) / 2
                        y0 = r_ * cell_h + (cell_h - th) / 2
                        target = fitz.Rect(x0, y0, x0 + tw, y0 + th)
                        if pg.rotation:
                            src_dpi = max(120, min(300, int(round(300 / max(k, 0.05)))))
                            np_.insert_image(target, pixmap=_rotate_pixmap(
                                pg.get_pixmap(dpi=src_dpi), 0))
                        else:
                            np_.show_pdf_page(target, doc, pno)
                data = out.tobytes(deflate=True, garbage=1)
            finally:
                out.close()
        finally:
            doc.close()
        os.makedirs(CONV_DIR, exist_ok=True)
        part = t + ".part"
        with open(part, "wb") as f:
            f.write(data)
        try:
            os.replace(part, t)
        except OSError:
            if not os.path.isfile(t):
                raise
        finally:
            if os.path.exists(part):
                try:
                    os.remove(part)
                except OSError:
                    pass
        want = (n + nup - 1) // nup
        try:
            with fitz.open(t) as chk:
                bad = chk.page_count != want
        except Exception:
            bad = True
        if bad:
            try:
                os.remove(t)
            except OSError:
                pass
            raise RuntimeError("多页拼版校验不通过（张数不符）")
    return t


def prune_cache(days=7):
    """清掉过期的缓存 PDF（源文件改过或删掉后，旧缓存再也不会被用到）。"""
    if not os.path.isdir(CONV_DIR):
        return
    cut = time.time() - days * 86400
    for f in os.scandir(CONV_DIR):
        if f.is_file() and f.name != "_mapping.tsv" and f.stat().st_mtime < cut:
            try:
                os.remove(f.path)
            except OSError:
                pass
    # 合并出来的临时 PDF 正常打完就删了，只有进程被杀才会剩下，顺手扫掉
    if os.path.isdir(MERGE_DIR):
        live = {j["path"] for j in list(JOBS.values()) if j.get("path")}
        for f in os.scandir(MERGE_DIR):
            if f.is_file() and f.path not in live and f.stat().st_mtime < cut:
                try:
                    os.remove(f.path)
                except OSError:
                    pass


# ---------------- 打印效果预览 ----------------
# 和 Acrobat 打印对话框右侧那格一个意思：把「照当前参数打出来会是什么样」画成一张纸。
# 它不负责看清内容（那是原文件预览的事），而是看缩放比、有没有被裁、横向页有没转对。

def paper_mm(paper):
    """纸张名 -> (宽 mm, 高 mm)。不认识的按 A4。"""
    return STD_SIZES.get(paper) or STD_SIZES["A4"]


def print_layout(page_w, page_h, paper, mode="fit", pct=100, margin_mm=None):
    """模拟 SumatraPDF 的排版规则，返回 (旋转角, 缩放比, 目标矩形, 是否超出纸面)。

    坐标用 pt、原点在纸张左上角。规则对应 build_settings 发出的参数：
    - 自动旋转：页面与纸张的横竖方向不一致就转 90°（Sumatra 的行为）
    - fit：缩放到「纸张减去不可打印边距」，可缩也可放大
    - shrink：同上但只缩不放，小页面维持 100%
    - noscale：固定 100%
    - custom：固定 pct%，超过 100% 的部分被纸裁掉（和 Acrobat 一致）
    """
    if margin_mm is None:
        try:
            margin_mm = float(CFG.get("preview_margin_mm", 4))
        except (TypeError, ValueError):
            margin_mm = 4.0
    wmm, hmm = paper_mm(paper)
    pw, ph = wmm * 72 / 25.4, hmm * 72 / 25.4
    m = max(0.0, margin_mm) * 72 / 25.4
    area_w, area_h = max(pw - 2 * m, 1.0), max(ph - 2 * m, 1.0)

    rot = 90 if (page_w > page_h) != (pw > ph) else 0
    w, h = (page_h, page_w) if rot else (page_w, page_h)
    if mode == "noscale":
        k = 1.0
    elif mode == "custom":
        k = min(4.0, max(0.1, float(pct) / 100.0))
    elif mode == "shrink":
        k = min(1.0, min(area_w / w, area_h / h))
    else:                                     # fit
        k = min(area_w / w, area_h / h)
    tw, th = w * k, h * k
    x0, y0 = (pw - tw) / 2, (ph - th) / 2
    # 只有超出纸张本身才算会被裁（贴着可打印边距不算，Acrobat 也不提示）
    over = tw > pw + 0.5 or th > ph + 0.5
    return rot, k, fitz.Rect(x0, y0, x0 + tw, y0 + th), over


def _rotate_pixmap(pix, rot):
    """把位图按顺时针 90° 的整数倍旋转。用来绕开一个坑：
    源页自带 /Rotate 时 show_pdf_page 不认它，自己转更可靠。"""
    rot %= 360
    if not rot:
        return pix
    import numpy as np
    a = np.frombuffer(pix.samples, dtype=np.uint8)
    a = a.reshape(pix.height, pix.stride // pix.n, pix.n)[:, :pix.width, :]
    a = np.ascontiguousarray(np.rot90(a, k=(4 - rot // 90) % 4))   # np.rot90 是逆时针
    h, w = a.shape[:2]
    cs = fitz.csGRAY if pix.n == 1 else fitz.csRGB
    return fitz.Pixmap(cs, w, h, a.tobytes(), False)


def preview_png(src, paper, mode="fit", pct=100, page_no=1, gray=False, dpi=120, nup=1):
    """按打印参数渲染整张纸的预览图，返回 (png bytes, 元信息 dict)。

    nup>1 时先多页拼版成每张纸一份页面，再照 noscale 渲染那张纸；page_no / meta 里的
    页数都按「纸张数」算，和真正送印时打出来的张数一致。
    """
    doc = fitz.open(nup_pdf(src, nup, paper) if nup > 1 else printable_pdf(src))
    try:
        if doc.needs_pass:
            raise RuntimeError("PDF 有密码，无法预览")
        n = doc.page_count
        if not n:
            raise RuntimeError("文档里没有页面")
        pno = max(1, min(int(page_no), n))
        pg = doc[pno - 1]
        page_rot = pg.rotation          # 源页自带的 /Rotate，get_pixmap 会应用
        # 多页拼版后的页面已经是纸张大小，按 noscale 100% 打（其余缩放档对它无意义）
        m = "noscale" if nup > 1 else mode
        wmm, hmm = paper_mm(paper)
        if nup > 1:
            rows, cols = NUP_GRID.get(nup, (1, 1))
            if cols > rows:                       # 拼版时横向网格用了横向纸
                wmm, hmm = hmm, wmm
        rot, k, rect, over = print_layout(pg.rect.width, pg.rect.height,
                                          paper, m, pct)
        pw, ph = wmm * 72 / 25.4, hmm * 72 / 25.4
        if nup > 1:
            # 拼版页本身就是这张纸，noscale 不转不缩，直接满幅居中（不扣边距，看整张纸）
            rot, k = 0, 1.0
            rect = fitz.Rect(0, 0, pw, ph)
            over = False
        out = fitz.open()
        try:
            # 自定义比例 >100% 时内容会超到纸外。show_pdf_page / insert_image
            # 遇到越界矩形会静默画空白，所以先把页面放大到「纸张 ∪ 内容」的包围盒，
            # 画完再在渲染时用 clip 只取纸张那一块。
            # +0.01pt 是给 new_page 的单精度量化留余量。
            bx0, by0 = min(0.0, rect.x0), min(0.0, rect.y0)
            bx1, by1 = max(pw, rect.x1), max(ph, rect.y1)
            np_ = out.new_page(width=bx1 - bx0 + 0.01, height=by1 - by0 + 0.01)
            target = fitz.Rect(rect.x0 - bx0, rect.y0 - by0,
                               rect.x1 - bx0, rect.y1 - by0)
            if page_rot:
                # 带 /Rotate 的页面走位图：先按呈现方向渲染（get_pixmap 认
                # /Rotate），自己转到位再贴。多一次重采样，但方向不会错。
                src_dpi = max(120, min(300, int(round(dpi / max(k, 0.05)))))
                np_.insert_image(target, pixmap=_rotate_pixmap(pg.get_pixmap(dpi=src_dpi), rot))
            else:
                # 普通页面走矢量：show_pdf_page 只搬内容流，不重采样
                np_.show_pdf_page(target, doc, pno - 1, rotate=rot)
            paper_rect = fitz.Rect(-bx0, -by0, pw - bx0, ph - by0)
            pix = np_.get_pixmap(clip=paper_rect, dpi=max(60, min(int(dpi), 200)),
                                 colorspace=fitz.csGRAY if gray else fitz.csRGB,
                                 alpha=False)
            png = pix.tobytes("png")
        finally:
            out.close()
    finally:
        doc.close()
    meta = {"page": pno, "pages": n, "pct": int(round(k * 100)),
            "rot": rot, "w": wmm, "h": hmm, "over": bool(over)}
    return png, meta


def preview_multi(names, paper, nup, page_no=1, gray=False, dpi=120):
    """把多个文件首尾相接、再按 nup 拼版，渲染第 page_no 张纸。跨文件多页预览用。

    复用 preview_png：先把多份 PDF 合并成一份临时 PDF，再让 nup 拼版作用于它。
    返回 (png bytes, meta)。"""
    names = [os.path.basename(n) for n in names]
    pdfs = []
    for n in names:
        fp = os.path.join(HOT, n)
        if not os.path.isfile(fp):
            raise RuntimeError(f"文件不存在：{n}")
        info = preflight(fp)
        if info.get("error"):
            raise RuntimeError(f"{n}：{info['detail']}")
        pdfs.append(info["pdf"])
    if not pdfs:
        raise RuntimeError("没有可预览的文件")
    merged = merge_pdfs(pdfs) if len(pdfs) > 1 else pdfs[0]
    try:
        return preview_png(merged, paper, "fit", 100, page_no, gray, dpi, nup)
    finally:
        if len(pdfs) > 1 and os.path.isfile(merged):
            try:
                os.remove(merged)
            except OSError:
                pass



def preflight(path, allow_convert=True):
    """读页数 / 尺寸 / 方向。按 mtime+size 缓存，避免每次刷新重解析。

    allow_convert=False 是给网页刷新用的：绝不在这里等 WPS 转 PDF，只把没转好的
    文件标成「转换中」并甩给后台转换线程，界面就不会因为一份 Word 卡住不动。
    """
    st = os.stat(path)
    stamp = (st.st_mtime, st.st_size)      # 注意别用 key 这个名字，下面还要用
    hit = CACHE.get(path)
    if hit and hit[0] == stamp:
        return hit[1]

    info = {"path": path, "name": os.path.basename(path),
            "kind": KIND.get(ext_of(path)) or ("图片" if ext_of(path) in IMG_EXT else ""),
            "mb": round(st.st_size / 1048576, 2), "mtime": st.st_mtime}
    if not allow_convert and needs_convert(path):
        fail = CONV_FAIL.get(path)
        if fail and fail[0] == stamp:      # 转过了、失败了，别再一遍遍重试
            info.update({"pages": 0, "size": "—", "orientation": "—", "paper": "A4",
                         "detail": f"解析失败：{fail[1]}", "error": fail[1]})
            CACHE[path] = (stamp, info)
            return info
        request_convert(path)
        info.update({"pages": 0, "size": "—", "orientation": "—", "paper": "A4",
                     "detail": "转换中…", "pending": True})
        return info                        # 不缓存：转好后要换成真实页数
    try:
        pdf = printable_pdf(path)      # Office/图片在这里转成 PDF
        info["pdf"] = pdf
        doc = fitz.open(pdf)
        pairs = Counter()
        exact = {}
        for pg in doc:
            r = pg.rect
            w = round(r.width * 25.4 / 72)
            h = round(r.height * 25.4 / 72)
            sk = (classify(w, h) or f"{min(w,h)}×{max(w,h)}mm",
                  "横向" if w > h else "纵向")
            pairs[sk] += 1
            exact.setdefault(sk, (min(w, h), max(w, h)))
        doc.close()
        sizes = {k[0] for k in pairs}
        orients = {k[1] for k in pairs}
        info["pages"] = sum(pairs.values())
        info["size"] = sizes.pop() if len(sizes) == 1 else "混合"
        info["orientation"] = orients.pop() if len(orients) == 1 else "混合"
        info["detail"] = "，".join(f"{s} {o} {n}页"
                                   for (s, o), n in pairs.most_common())
        top = max(pairs.items(), key=lambda kv: kv[1])[0]
        # 尺寸与标准纸有偏差（如出血版）时把实际 mm 标出来
        if len(pairs) == 1 and top[0] in STD_SIZES:
            ew, eh = exact[top]
            sw, sh = STD_SIZES[top[0]]
            if abs(ew - sw) > 1 or abs(eh - sh) > 1:
                info["detail"] += f"（实际 {ew}×{eh}mm）"
        info["paper"] = top[0] if top[0] in CFG["papers"] else "A4"
    except Exception as e:
        info.update({"pages": 0, "size": "—", "orientation": "—",
                     "detail": f"解析失败：{e}", "paper": "A4",
                     "error": str(e)})
    CACHE[path] = (stamp, info)
    return info


def set_hotfolder(path):
    """切换热文件夹并记住到 config.json。"""
    global HOT
    path = os.path.abspath(path)
    if not os.path.isdir(path):
        raise ValueError("文件夹不存在：" + path)
    HOT = path
    CFG["hotfolder"] = path
    CACHE.clear()                       # 换目录了，上一目录的预检/转换状态都作废
    with CONV_STATE_LOCK:
        CONV_WANT.clear()
        CONV_FAIL.clear()
    with open(CFG_PATH, "w", encoding="utf-8") as f:
        json.dump(CFG, f, ensure_ascii=False, indent=2)
    return HOT


PICK_LOCK = threading.Lock()


def pick_folder():
    """弹系统文件夹选择框（本机 GUI）。取消返回空串。"""
    import tkinter as tk
    from tkinter import filedialog
    with PICK_LOCK:                     # 连点两下只弹一个框
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)   # 否则对话框可能藏在浏览器后面
        try:
            return filedialog.askdirectory(initialdir=HOT, title="选择热文件夹",
                                           mustexist=True)
        finally:
            root.destroy()


def list_files():
    """列热文件夹里能打印的文件。这里只读缓存，不在这里等转换 —— 新丢进来的
    Word/表格/图片先显示「转换中」，后台线程转好了下一轮刷新就换成真实页数。"""
    out = []
    if not os.path.isdir(HOT):
        return out
    paths = []
    for name in os.listdir(HOT):
        p = os.path.join(HOT, name)
        if (ext_of(name) in OFFICE_EXT or ext_of(name) in IMG_EXT
                or ext_of(name) == ".pdf") and os.path.isfile(p):
            paths.append(p)
    for p in paths:
        try:
            out.append(preflight(p, allow_convert=False))
        except OSError:
            pass
    out.sort(key=lambda x: -x["mtime"])
    return out


# ---------------- 打印队列 ----------------
# 三道工序分开排队：PREP_Q（转 PDF / 缩放 / 合并）→ READY_Q（汇总）→ 各打印机自己的队列（送印）。
# 分开的好处：下单请求立刻返回；打前一份的时候，后一份已经在准备了。
# 送印这一层按打印机拆开：同一台机器必须串行（多任务同时送会乱页、还会抢带宽），
# 但 D95 在打的时候彩机 C7580 没道理跟着排队 —— 不同机器之间并行。
PREP_Q = queue.Queue()
READY_Q = queue.Queue()
JOBS = {}          # id -> job
JOBS_LOCK = threading.Lock()
_seq = [0]
MERGE_DIR = os.path.join(CONV_DIR, "_合并")
# 还在跑的状态。清空队列遇到这几种要拦住，等打完再说
ACTIVE_STATES = ("排队中", "准备中", "打印中")

PARALLEL = bool(CFG.get("print_parallel", True))
if PARALLEL:
    PRINTER_Q = {p: queue.Queue() for p in PRINTERS}          # 一台机器一条队列
elif PRINTERS:
    # 串行模式（print_parallel=false）：所有机器共用一条队列，就是老行为
    _common = queue.Queue()
    PRINTER_Q = {p: _common for p in PRINTERS}
else:
    PRINTER_Q = {}
# 没有匹配到打印机时也得有地方放活儿，别在分派线程里炸
PRINT_QUEUES = list(dict.fromkeys(PRINTER_Q.values())) or [queue.Queue()]

DUPLEX = {"single": "simplex", "long": "duplexlong", "short": "duplexshort"}
# DMPAPER_* 编号。本机驱动的纸名叫 "A4(210x297mm)"，Sumatra 的 paper=A4 匹配不上
# （它只认完整同名或 "A4 " 带空格前缀），会退回"按文档尺寸自定义纸张"，驱动找不到
# 这种纸就落到 A3 —— 233x479mm 的图打成 A3 就是这么来的。paperkind 直接给编号。
# 这张表只在 SumatraPDF 兜底时才用；gdi_print 里有一份同样的（自检会比对两份是否一致）。
PAPERKIND = {"A4": 9, "A3": 8, "A5": 11, "B4": 12, "B5": 13, "Letter": 1, "Legal": 5, "11x17": 3}

# ---------------- 页码范围 ----------------
# SumatraPDF 的 -print-settings 原生认页码（5 / 2-6 / 5- / -1 末页），但合并打印会把多份
# PDF 首尾相接、页号整体后移，所以统一先解析成具体页号，prepare 里按各份页数偏移后再压回
# "1-3,5" 这种紧凑写法送印。
_PG_CHAR = str.maketrans({"０": "0", "１": "1", "２": "2", "３": "3", "４": "4", "５": "5",
                          "６": "6", "７": "7", "８": "8", "９": "9", "，": ",", "、": ",",
                          "；": ",", ";": ",", "－": "-", "—": "-", "–": "-", "~": "-",
                          "～": "-"})
_PG_ONE = re.compile(r"^(-?\d+)$")
_PG_RANGE = re.compile(r"^(-?\d+)-(-?\d*)$")


def parse_pages(spec, total=None):
    """页码串 -> 升序去重的页号列表；空串表示全部（total 未知时返回 None）。

    total 已知就校验越界并抛 ValueError；不知道总页数（文件还在转）时只校验语法，
    需要总页数才能解释的写法（5- / 负数）留到 prepare 再算。
    """
    s = str(spec or "").strip()
    if not s:
        return list(range(1, total + 1)) if total else None
    t = s.translate(_PG_CHAR)
    t = re.sub(r"\s*-\s*", "-", t)          # 「1 - 3」也认
    t = re.sub(r"\s+", ",", t)              # 空格当分隔符
    nums = set()

    def resolve(v, tok):
        if v < 0:                            # 负数从末尾数，-1 是最后一页（同 SumatraPDF）
            if not total:
                raise ValueError(f"「{tok}」要等页数出来才能算")
            v = total + v + 1
        if v < 1:
            raise ValueError(f"页码不存在：{tok}")
        if total and v > total:
            raise ValueError(f"第 {v} 页超出范围（共 {total} 页）")
        return v

    for tok in [x for x in t.split(",") if x]:
        if tok.lower() in ("all", "全部", "全部页", "*"):
            if not total:
                raise ValueError("「全部」要等页数出来才能算")
            nums.update(range(1, total + 1))
        elif _PG_ONE.match(tok):
            nums.add(resolve(int(tok), tok))
        elif _PG_RANGE.match(tok):
            a, b = _PG_RANGE.match(tok).groups()
            hi = resolve(int(b), tok) if b else total
            if not hi:
                raise ValueError(f"「{tok}」要等页数出来才能算")
            lo = resolve(int(a), tok)
            nums.update(range(min(lo, hi), max(lo, hi) + 1))
        else:
            raise ValueError(f"看不懂的页码：{tok}")
    if not nums:
        raise ValueError("没有有效的页码")
    return sorted(nums)


def pages_text(nums):
    """页号列表 -> "1-3,5"。连着三页以上压成范围，省得参数串太长。"""
    out, i, nums = [], 0, list(nums)
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        if j - i >= 2:
            out.append(f"{nums[i]}-{nums[j]}")
        else:
            out.extend(str(x) for x in nums[i:j + 1])
        i = j + 1
    return ",".join(out)


def build_settings(job):
    s = []
    if job.get("range"):
        s.append(job["range"])          # 页码放最前，和 SumatraPDF 文档里的示例一致
    s.append(f'{job["copies"]}x')
    if job["copies"] > 1:
        # 显式要「逐份」：一份打完再打下一份。不写就听驱动的默认值，
        # 有的驱动按页刷（第1页xN、第2页xN…），合并打印会整叠错乱。
        s.append("collate")
    s.append(DUPLEX[job["duplex"]])
    s.append("color" if job["color"] else "monochrome")
    kind = PAPERKIND.get(job["paper"])
    s.append(f"paperkind={kind}" if kind else f'paper={job["paper"]}')
    # 不传 portrait/landscape：那两个参数只旋转内容、不改纸张方向，会和 SumatraPDF
    # 自带的自动旋转叠加成转两次（横向内容缩成 71% 塞在纵向纸里）。
    # 交给自动旋转：宽于高的页面自动转 90° 适配纸张，满版。
    mode = job.get("scale", "fit")      # 老调用方不带 scale 时维持原行为
    s.append(SCALE.get(mode, "fit"))
    if mode != "fit":
        s.append("center")              # 内容小于纸面时居中，和 Acrobat 观感一致
    return ",".join(s)


# ---------------- 计费 ----------------
LOG_PATH = os.path.join(BASE, "打印记录.csv")
A4_PRICE = {False: 0.5, True: 2.0}   # A4 每页：黑白 / 彩色；（True 在 dict 里等于 1）
LOG_COLS = ["时间", "文件", "页数", "份数", "纸张", "色彩", "打印机", "单价元/页", "金额元"]


def unit_price(paper, color):
    p = A4_PRICE[bool(color)]
    return p * 2 if paper == "A3" else p      # A3 翻倍


def log_job(job, ts=None):
    new = not os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", newline="", encoding="utf-8-sig" if new else "utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(LOG_COLS)
        w.writerow([time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(ts or time.time())),
                    job["name"], job["pages"], job["copies"], job["paper"],
                    "彩色" if job["color"] else "黑白", job["printer"],
                    f'{job["unit"]:.2f}', f'{job["amount"]:.2f}'])


def read_log():
    # shortcut: 每次全量读 CSV，几千条无压力；上万条再换 sqlite
    if not os.path.exists(LOG_PATH):
        return {"rows": [], "today": 0, "total": 0, "count": 0}
    with open(LOG_PATH, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    today = time.strftime("%Y-%m-%d")
    return {"rows": rows[-300:][::-1], "count": len(rows),
            "today": round(sum(float(r["金额元"]) for r in rows
                               if r["时间"].startswith(today)), 2),
            "total": round(sum(float(r["金额元"]) for r in rows), 2)}


def conv_worker():
    """后台把 Office/图片转成 PDF。攒一小会儿再一起转，一次 PowerShell/WPS 启动搞定一批。"""
    while True:
        batch = [CONV_Q.get()]
        try:
            while len(batch) < 12:
                batch.append(CONV_Q.get(timeout=0.15))      # 同一批刚丢进来的凑一次
        except queue.Empty:
            pass
        office = [b for b in batch if ext_of(b) in OFFICE_EXT]
        if office:
            try:
                ensure_pdfs(office)
            except Exception as e:
                print("[警告] Office 转换出错:", e)
        for b in batch:
            err = ""
            if ext_of(b) in IMG_EXT:
                try:
                    img_to_pdf(b)
                except Exception as e:
                    err = f"图片转换失败：{e}"
            ok = not err and bool(cached_pdf(b))
            if not ok and not err:
                err = "转换失败（WPS 没转出 PDF：文件损坏/加密，或 WPS 正弹着对话框）"
            with CONV_STATE_LOCK:
                stamp = CONV_WANT.pop(b, None)
                if ok:
                    CONV_FAIL.pop(b, None)
                elif stamp:
                    CONV_FAIL[b] = (stamp, err)
            CACHE.pop(b, None)          # 状态变了，下一轮刷新重新解析


def select_pages(pdf, nums):
    """抽出 nums 指定的页（1 起），写成一份临时 PDF。多页拼版前用来真正剔除
    不打的页（拼版后页码变了，不能再靠 SumatraPDF 的页码参数）。"""
    nums = sorted(set(nums))
    doc = fitz.open(pdf)
    try:
        out = fitz.open()
        try:
            for p in nums:
                out.insert_pdf(doc, from_page=p - 1, to_page=p - 1)
            os.makedirs(MERGE_DIR, exist_ok=True)
            t = os.path.join(MERGE_DIR, os.urandom(4).hex() + ".pdf")
            part = t + ".part"
            out.save(part, deflate=True, garbage=1)
            os.replace(part, t)
        finally:
            out.close()
    finally:
        doc.close()
    return t


def merge_pdfs(pdfs):
    """把多份 PDF 首尾相接写成一份，返回新文件路径。各页尺寸保持不变。"""
    out = fitz.open()
    try:
        for p in pdfs:
            d = fitz.open(p)
            try:
                if d.needs_pass:
                    raise RuntimeError(f"{os.path.basename(p)} 有密码，无法合并")
                out.insert_pdf(d)
            finally:
                d.close()
        os.makedirs(MERGE_DIR, exist_ok=True)
        t = os.path.join(MERGE_DIR, os.urandom(4).hex() + ".pdf")
        part = t + ".part"                 # 半截文件不会被当成品
        # garbage=1 只丢没人引用的对象；garbage=3 还要去重、重排，纯浪费（实测慢一倍多，体积一样）
        out.save(part, deflate=True, garbage=1)
        os.replace(part, t)
    finally:
        out.close()
    return t


# 只有这几项完全相同才能拼成一份打：拼起来就是一次送印，参数只能有一套
MERGE_KEY = ("printer", "copies", "color", "paper", "duplex", "scale", "pct", "nup")


def group_jobs(preps):
    """按送印参数分组。参数不同的各成一组，组内可以拼成一份一次送印。"""
    groups = {}
    for p in preps:
        groups.setdefault(tuple(p[k] for k in MERGE_KEY), []).append(p)
    return list(groups.values())


def _merge_group(ps):
    """把一组（同参数）的预备任务并成一个任务。真正拼 PDF 留到 prepare 里做，
    下单请求就不用等合并了。"""
    if len(ps) == 1:
        return ps[0]
    j = dict(ps[0])
    j["srcs"] = [p["srcs"][0] for p in ps]
    j["names"] = [p["name"] for p in ps]
    j["name"] = f'{ps[0]["name"]} 等{len(ps)}份（合并）'
    j["specs"] = [(p.get("specs") or [""])[0] for p in ps]   # 每份各自的页码，prepare 里按偏移翻译
    j["range"] = ""                             # 合并后的页码要重算，别留第一份的
    j["pages"] = sum(p.get("pages") or 0 for p in ps)
    j["tmp"] = True
    return j


def _validate(it):
    """只做轻量校验 + 读缓存里的页数，返回下单字段。耗时的转换/缩放/合并留给 prepare。

    pages 可能是 0（文件刚丢进来还没转成 PDF），prepare 时补上真实页数。
    """
    path = it.get("path") or os.path.join(HOT, os.path.basename(it.get("name", "")))
    path = os.path.abspath(path)
    if os.path.dirname(path) != os.path.abspath(HOT) or not os.path.isfile(path):
        raise ValueError(f"文件不在热文件夹内: {path}")
    printer = it.get("printer")
    if printer not in PRINTERS:
        raise ValueError(f"打印机不在列表中: {printer}")
    copies = int(it.get("copies", 1))
    if not 1 <= copies <= 999:
        raise ValueError("份数需在 1-999")
    paper = it.get("paper")
    if paper not in CFG["papers"]:
        raise ValueError(f"纸张不在列表中: {paper}")
    duplex = it.get("duplex", "single")
    if duplex not in DUPLEX:
        raise ValueError("双面参数非法")
    color = bool(it.get("color"))
    scale = it.get("scale", "fit")
    if scale not in SCALE:
        raise ValueError(f"缩放方式非法: {scale}")
    pct = 100
    if scale == "custom":
        try:
            pct = int(it.get("scale_pct", 100))
        except (TypeError, ValueError):
            raise ValueError("自定义比例必须是整数")
        if not 10 <= pct <= 400:
            raise ValueError("自定义比例需在 10-400%")
    try:
        nup = int(it.get("nup", 1) or 1)
    except (TypeError, ValueError):
        raise ValueError("每页张数必须是整数")
    if nup not in NUP_GRID:
        raise ValueError("每页张数只能选 1/2/4/6/9/16")
    # 缓存里已经有解析结果就直接用；还没转好就先当 0 页，别在这里等 WPS
    info = preflight(path, allow_convert=False)
    if info.get("error"):
        raise ValueError(f"{os.path.basename(path)}：{info['detail']}")
    unit = unit_price(paper, color)
    total = info.get("pages") or 0
    # 页码范围：空 = 全部。页数还没解析出来（转换中）就先存原文，prepare 时按真实页数校验
    spec = str(it.get("pages_spec") or "").strip()
    pick = parse_pages(spec, total) if spec and total else None
    pages = len(pick) if pick else total
    # 多页拼版：账单按纸张数算（N 页缩印到一张纸上），不是按文档页数
    sheets = (pages + nup - 1) // nup if pages else 0
    return {"name": os.path.basename(path), "names": [os.path.basename(path)],
            "srcs": [path], "printer": printer, "copies": copies, "color": color,
            "paper": paper, "duplex": duplex, "scale": scale, "pct": pct, "nup": nup,
            "specs": [spec], "range": pages_text(pick) if pick else "",
            "pages": pages, "unit": unit,
            "amount": round(unit * copies * sheets, 2)}


def prepare(job):
    """真正备货：转 PDF、按比例重排、多份合并，最后算出页数和金额。

    这道工序放在后台线程里跑，下单请求不用等它；它和「送印」是两条流水线，
    所以在打前一份的时候，后面几份就已经在转了。
    """
    srcs = job["srcs"]
    specs = job.get("specs") or [""] * len(srcs)
    office = [s for s in srcs if ext_of(s) in OFFICE_EXT]
    if office:
        ensure_pdfs(office)                   # 一次 PowerShell 启动转一批
    pdfs, counts = [], []
    for s in srcs:
        info = preflight(s)                   # 此刻已经转好，缓存里是真实页数
        if info.get("error"):
            raise RuntimeError(f"{os.path.basename(s)}：{info['detail']}")
        counts.append(info.get("pages") or 0)
        pdfs.append(info["pdf"])
    if job["scale"] == "custom" and job["pct"] != 100:
        pdfs = [scaled_pdf(s, job["pct"]) for s in srcs]   # 用原始文件算缓存名，源一改自动失效
    job["path"] = merge_pdfs(pdfs) if len(pdfs) > 1 else pdfs[0]

    # 页码：合并后各份的页号整体后移，所以按每份自己的页数翻译成本份内页号再加偏移。
    # 真页数到这里才准（下单时文件可能还在转），越界也在这里报错，不发出去打。
    nums, off = [], 0
    for i, (n, spec) in enumerate(zip(counts, specs)):
        if spec:
            try:
                nums.extend(off + p for p in parse_pages(spec, n))
            except ValueError as e:
                raise RuntimeError(f"{os.path.basename(srcs[i])} 页码范围：{e}")
        else:
            nums.extend(range(off + 1, off + n + 1))
        off += n
    job["pages"] = len(nums)
    job["range"] = "" if job["pages"] == off else pages_text(nums)   # 全选就不用写页码
    nup = job.get("nup", 1)
    if nup > 1:
        # 多页拼版后页数变了，页码参数不能再交给 SumatraPDF；先把要打的页真正抽成
        # 一份 PDF（有页码筛选时），再整份拼版，最后按 noscale 打。
        if job["range"]:
            job["path"] = select_pages(job["path"], nums)
            job["range"] = ""
        nup_path = nup_pdf(job["path"], nup, job["paper"])
        if nup_path != job["path"]:
            job["path"] = nup_path
        job["sheets"] = (job["pages"] + nup - 1) // nup
        job["scale"] = "noscale"          # 拼版后的页面已是纸张大小，送印不缩放
    job["settings"] = build_settings(job)
    amount_pages = job.get("sheets", job["pages"])
    job["amount"] = round(job["unit"] * job["copies"] * amount_pages, 2)
    return job


def enqueue(items, merge=False):
    """校验后立刻入队，返回 [{id,name,state}]。转换/缩放/合并在后台做，请求马上返回。

    merge=True 时同参数的文件归成一组，等准备阶段再拼成一份 PDF 一次送印。
    """
    preps = [_validate(it) for it in items]
    jobs = [_merge_group(g) for g in group_jobs(preps)] if merge else preps
    ok = []
    for p in jobs:
        _seq[0] += 1
        job = {"id": _seq[0], "state": "排队中", "msg": "", **p}
        job["settings"] = build_settings(job)
        with JOBS_LOCK:
            JOBS[job["id"]] = job
        PREP_Q.put(job)
        ok.append({"id": job["id"], "name": job["name"], "names": job["names"],
                   "state": job["state"], "settings": job["settings"]})
    return ok


def prep_worker():
    while True:
        job = PREP_Q.get()
        job["state"] = "准备中"
        try:
            prepare(job)
            job["state"] = "排队中"
            READY_Q.put(job)
        except Exception as e:
            job["state"] = "失败"
            job["msg"] = str(e)
            job["end"] = time.time()


# SumatraPDF 的退出码就是它自己的 PrintResult 枚举，比它打印的日志好看得多
SUMATRA_ERR = {1: "打印失败", 2: "打不开这个文件", 3: "文档不允许打印",
               4: "找不到打印机", 5: "打印机/驱动报错", 6: "没有打印权限"}


def _first_line(raw, n=160):
    """子进程的输出可能是一大坨日志，只留第一行有用的。"""
    txt = (raw or b"").decode("utf-8", "ignore").strip()
    line = next((x for x in txt.splitlines() if x.strip()), "")
    return line[:n]


def send_sumatra(job):
    """老办法：交给 SumatraPDF。现在只在 GDI 送印不可用/兜底时走。"""
    args = [SUMATRA, "-print-to", job["printer"],
            "-print-settings", job["settings"], "-silent", job["path"]]
    r = subprocess.run(args, capture_output=True, timeout=600,
                       creationflags=0x08000000)
    if r.returncode == 0:
        return True, ""
    msg = SUMATRA_ERR.get(r.returncode) or _first_line(r.stderr or r.stdout)
    return False, f"SumatraPDF：{msg}"


def send_gdi(job):
    """自己走 GDI 送印。返回 (成功, 错误信息, 失败时能否安全换引擎重试)。

    gdi_print.py 放子进程里跑：万一某个驱动把进程搞崩，崩的只是这个子进程，
    网页服务还在；超时也和原来一样能掐掉。
    退出码 2 = 一页都没送出去（可以换引擎再来）；3/5 = 别重试（会重复出纸或本来就打不了）
    """
    args = [sys.executable, os.path.join(BASE, "gdi_print.py"), job["path"], job["printer"],
            "--copies", str(job["copies"]), "--paper", job["paper"],
            "--duplex", job["duplex"], "--scale", job.get("scale", "fit"),
            "--color", "1" if job["color"] else "0"]
    if job.get("range"):
        args += ["--range", job["range"]]
    r = subprocess.run(args, capture_output=True, timeout=600,
                       creationflags=0x08000000)
    if r.returncode == 0:
        return True, "", False
    msg = _first_line(r.stderr) or f"送印失败（退出码 {r.returncode}）"
    return False, msg, r.returncode == 2


def print_job(job):
    """送一份进打印机：成功才记账，失败把原因写进 job 里给页面看。"""
    job["state"] = "打印中"
    job["start"] = time.time()
    try:
        if ENGINE == "gdi":
            ok, msg, retry_safe = send_gdi(job)
            # 一页都没送出去就失败了，换 SumatraPDF 再试一次是安全的；
            # 已经送出去过就别重试，否则重复出纸、重复计费
            if not ok and retry_safe and CFG.get("engine_fallback", True):
                ok, msg2 = send_sumatra(job)
                if ok:
                    print(f"[提示] {job['name']} GDI 送印失败（{msg}），已改用 SumatraPDF")
                else:
                    msg = f"{msg}；{msg2}"
        else:
            ok, msg = send_sumatra(job)
        if ok:
            job["state"] = "完成"
            log_job(job)              # 只有真正送印成功才记账
        else:
            job["state"] = "失败"
            job["msg"] = msg
    except subprocess.TimeoutExpired:
        job["state"] = "失败"
        job["msg"] = "超时（>10 分钟）"
    except Exception as e:
        job["state"] = "失败"
        job["msg"] = str(e)
    job["end"] = time.time()
    if job.get("tmp"):            # 合并出来的临时 PDF 打完就没用了
        try:
            os.remove(job["path"])
        except OSError:
            pass


def dispatch_worker():
    """把备好的活按打印机分到各自的队列里。"""
    while True:
        job = READY_Q.get()
        PRINTER_Q.get(job["printer"], PRINT_QUEUES[0]).put(job)


def print_worker(pq):
    """一台打印机一条队列：同机严格按顺序，不同机互不排队。"""
    while True:
        print_job(pq.get())


_threads = [conv_worker, prep_worker, dispatch_worker]
for _q in PRINT_QUEUES:
    _threads.append(lambda q=_q: print_worker(q))
for _t in _threads:
    threading.Thread(target=_t, daemon=True).start()
print(f"[送印] 引擎 {ENGINE}，{len(PRINT_QUEUES)} 条送印队列"
      f"（{'按打印机并行' if PARALLEL else '全局串行'}）")


def recent_jobs():
    with JOBS_LOCK:
        js = sorted(JOBS.values(), key=lambda j: -j["id"])[:80]
        return [{k: v for k, v in j.items() if k not in ("path", "srcs")} for j in js]


# ---------------- 清空队列 ----------------

def recycle(paths):
    """把文件送进回收站（可还原），返回没能删掉的路径。"""
    import ctypes
    from ctypes import wintypes

    class OP(ctypes.Structure):
        _fields_ = [("hwnd", wintypes.HWND), ("wFunc", wintypes.UINT),
                    ("pFrom", wintypes.LPCWSTR), ("pTo", wintypes.LPCWSTR),
                    ("fFlags", ctypes.c_ushort), ("aborted", wintypes.BOOL),
                    ("hNameMappings", ctypes.c_void_p),
                    ("lpszProgressTitle", wintypes.LPCWSTR)]

    # FOF_ALLOWUNDO|NOCONFIRMATION|SILENT|NOERRORUI，wFunc=3 是 FO_DELETE
    op = OP(None, 3, "\0".join(paths) + "\0\0", None, 0x40 | 0x10 | 0x04 | 0x400,
            False, None, None)
    ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    return [p for p in paths if os.path.exists(p)]    # 还在盘上的就是没删掉


def clear_files(names):
    """清空队列：删掉热文件夹里这些文件（进回收站），连带它们的缓存 PDF
    和打剩的合并临时文件。只认名单里的文件名，热文件夹里别的东西不碰。"""
    root = os.path.abspath(HOT)
    paths = []
    for n in names:
        p = os.path.abspath(os.path.join(root, os.path.basename(n)))
        if os.path.dirname(p) != root:
            raise ValueError(f"文件不在热文件夹内: {p}")
        if os.path.isfile(p):
            paths.append(p)
    if not paths:
        return {"removed": 0, "cache": 0, "merged": 0, "failed": []}
    live = [j for j in list(JOBS.values()) if j["state"] in ACTIVE_STATES]
    want = {os.path.basename(p) for p in paths}
    # 合并任务的名字是一串「a.pdf 等N份」，所以按任务的 names 判断有没有在打
    busy = [n for j in live for n in j.get("names") or [] if n in want]
    if busy:
        raise ValueError("还有任务在打，等打完再清：" + "、".join(busy[:3]))

    cache = 0
    for p in paths:
        t = conv_target(p)
        # 缓存名是 名字_大小_时间.pdf，自定义比例再带 _s{pct}、多页拼版带 _nup{N}。
        # 只清这一份的，同名的别的文件不误伤
        for f in [t] + glob.glob(t[:-4] + "_s*.pdf") + glob.glob(t[:-4] + "_nup*.pdf"):
            if os.path.isfile(f):
                try:
                    os.remove(f)
                    cache += 1
                except OSError:
                    pass
    merged = 0
    inuse = {j["path"] for j in live if j.get("path")}   # 正在打的那份合并 PDF 不能删
    if os.path.isdir(MERGE_DIR):
        for f in os.scandir(MERGE_DIR):
            if f.is_file() and f.path not in inuse:
                try:
                    os.remove(f.path)
                    merged += 1
                except OSError:
                    pass
    failed = recycle(paths)
    return {"removed": len(paths) - len(failed), "cache": cache, "merged": merged,
            "failed": [os.path.basename(p) for p in failed]}


class Handler(BaseHTTPRequestHandler):
    server_version = "LinkPrint"
    # 长连接：前端每秒轮询状态，不用每次都重新 TCP 握手/关闭（Windows 上还会留一堆
    # TIME_WAIT）。每个响应都带 Content-Length，所以 1.1 的 keep-alive 是安全的。
    protocol_version = "HTTP/1.1"
    timeout = 120          # 空闲连接最多挂 2 分钟，超了就关，别攒线程

    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8", headers=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False)
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (headers or {}).items():     # 自定义头只放 ASCII，中文别往里塞
            self.send_header(k, str(v))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        u = urlparse(self.path)
        p = u.path
        if p in ("/", "/index.html"):
            with open(os.path.join(WEB_DIR, "index.html"), encoding="utf-8") as f:
                return self._send(200, f.read(), "text/html; charset=utf-8")
        if p == "/api/files":
            return self._send(200, {"folder": HOT, "printers": PRINTERS,
                                    "papers": CFG["papers"],
                                    "prices": {f"{pp}|{c}": unit_price(pp, c)
                                               for pp in CFG["papers"] for c in (0, 1)},
                                    "files": list_files()})
        if p == "/api/log":
            return self._send(200, read_log())
        if p == "/api/status":
            return self._send(200, {"jobs": recent_jobs()})
        if p == "/api/open-folder":
            os.startfile(HOT)
            return self._send(200, {"ok": True})
        if p == "/api/pick-folder":
            chosen = pick_folder()
            if not chosen:
                return self._send(200, {"folder": HOT, "canceled": True})
            try:
                return self._send(200, {"folder": set_hotfolder(chosen)})
            except Exception as e:
                return self._send(400, {"error": str(e)})
        if p.startswith("/icons/"):
            name = os.path.basename(unquote(p))
            fp = os.path.join(WEB_DIR, "icons", name)
            if not os.path.isfile(fp):
                return self._send(404, {"error": "not found"})
            ctype = {"svg": "image/svg+xml", "png": "image/png",
                     "ico": "image/x-icon"}.get(name.rsplit(".", 1)[-1].lower(),
                                                "application/octet-stream")
            with open(fp, "rb") as f:
                return self._send(200, f.read(), ctype)
        if p == "/api/printpreview":
            q = parse_qs(u.query)
            g = lambda k, d: (q.get(k) or [d])[0]
            name = os.path.basename(g("name", ""))
            fp = os.path.join(HOT, name)
            if not name or not os.path.isfile(fp):
                return self._send(404, {"error": "文件不存在"})
            try:
                png, meta = preview_png(
                    fp, g("paper", "A4"), g("scale", "fit"),
                    float(g("pct", "100")), int(g("page", "1")),
                    g("gray", "0") in ("1", "true"), int(g("dpi", "120")),
                    int(g("nup", "1")))
            except Exception as e:
                return self._send(400, {"error": str(e)})
            return self._send(200, png, "image/png", {
                "X-LP-Pct": meta["pct"], "X-LP-Rot": meta["rot"],
                "X-LP-Page": meta["page"], "X-LP-Pages": meta["pages"],
                "X-LP-W": meta["w"], "X-LP-H": meta["h"],
                "X-LP-Over": 1 if meta["over"] else 0})
        if p == "/api/preview-multi":
            q = parse_qs(u.query)
            g = lambda k, d: (q.get(k) or [d])[0]
            names = [n for n in g("names", "").split(",") if n.strip()]
            if not names:
                return self._send(404, {"error": "没有指定文件"})
            try:
                png, meta = preview_multi(
                    names, g("paper", "A4"), int(g("nup", "1")),
                    int(g("page", "1")),
                    g("gray", "0") in ("1", "true"), int(g("dpi", "120")))
            except Exception as e:
                return self._send(400, {"error": str(e)})
            return self._send(200, png, "image/png", {
                "X-LP-Pct": meta["pct"], "X-LP-Rot": meta["rot"],
                "X-LP-Page": meta["page"], "X-LP-Pages": meta["pages"],
                "X-LP-W": meta["w"], "X-LP-H": meta["h"],
                "X-LP-Over": 1 if meta["over"] else 0})
        if p.startswith("/pdf/"):
            name = os.path.basename(unquote(p[5:]))
            fp = os.path.join(HOT, name)
            if not os.path.isfile(fp):
                return self._send(404, {"error": "文件不存在"})
            info = preflight(fp)
            if not info.get("pdf"):
                return self._send(400, {"error": info.get("detail", "无法预览")})
            with open(info["pdf"], "rb") as f:
                return self._send(200, f.read(), "application/pdf")
        return self._send(404, {"error": "not found"})

    def do_POST(self):
        p = urlparse(self.path).path
        try:
            n = int(self.headers.get("Content-Length", 0))
            data = json.loads(self.rfile.read(n) or b"{}")
            if p == "/api/print":
                return self._send(200, {"queued": enqueue(data.get("jobs", []),
                                                          bool(data.get("merge")))})
            if p == "/api/clear":
                return self._send(200, clear_files(data.get("names", [])))
            return self._send(404, {"error": "not found"})
        except Exception as e:
            return self._send(400, {"error": str(e)})


def main():
    if not os.path.isdir(HOT):
        os.makedirs(HOT, exist_ok=True)
    prune_cache()
    if not PRINTERS:
        print("[警告] 未匹配到任何打印机，请检查 config.json 的 printers 关键字")
    if not os.path.isfile(C2P):
        print("[提示] 没找到 Convert2PDF.ps1，Word/表格/PPT 将无法打印（config.json 的 convert2pdf）")
    port = CFG["port"]
    srv = None
    for cand in range(port, port + 20):
        try:
            srv = ThreadingHTTPServer(("127.0.0.1", cand), Handler)
            port = cand
            break
        except OSError:
            continue
    if srv is None:
        sys.exit("端口 8777-8796 全部被占用")

    print("=" * 56)
    print(" LinkPrint 就绪   http://127.0.0.1:%d" % port)
    print(" 热文件夹  ", HOT)
    print(" 打印机    ", " | ".join(PRINTERS) or "（无）")
    print(" 关闭此窗口即停止服务")
    print("=" * 56)
    if "--no-browser" not in sys.argv:
        threading.Timer(1.0, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    srv.serve_forever()


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
