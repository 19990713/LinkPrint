# -*- coding: utf-8 -*-
"""LinkPrint —— 热文件夹 PDF 打印系统（预检 + 批量下单）

启动：启动打印系统.bat 或 python print_system.py
流程：PDF 丢进「热文件夹」→ 网页里逐份调参数 → 一次性发出打印。
"""
import csv
import json
import os
import queue
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from collections import Counter
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import unquote, urlparse

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
    # Office -> PDF 转换器（WPS COM）。找不到它就只支持 PDF
    "convert2pdf": r"D:\脚本文件\多文件合并PDF\Convert2PDF.ps1",
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



def preflight(path):
    """读页数 / 尺寸 / 方向。按 mtime+size 缓存，避免每次刷新重解析。"""
    st = os.stat(path)
    key = (st.st_mtime, st.st_size)
    hit = CACHE.get(path)
    if hit and hit[0] == key:
        return hit[1]

    info = {"path": path, "name": os.path.basename(path),
            "kind": KIND.get(ext_of(path)) or ("图片" if ext_of(path) in IMG_EXT else ""),
            "mb": round(st.st_size / 1048576, 2), "mtime": st.st_mtime}
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
            key = (classify(w, h) or f"{min(w,h)}×{max(w,h)}mm",
                   "横向" if w > h else "纵向")
            pairs[key] += 1
            exact.setdefault(key, (min(w, h), max(w, h)))
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
    CACHE[path] = (key, info)
    return info


def set_hotfolder(path):
    """切换热文件夹并记住到 config.json。"""
    global HOT
    path = os.path.abspath(path)
    if not os.path.isdir(path):
        raise ValueError("文件夹不存在：" + path)
    HOT = path
    CFG["hotfolder"] = path
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
    out = []
    if not os.path.isdir(HOT):
        return out
    paths = []
    for name in os.listdir(HOT):
        p = os.path.join(HOT, name)
        if (ext_of(name) in OFFICE_EXT or ext_of(name) in IMG_EXT
                or ext_of(name) == ".pdf") and os.path.isfile(p):
            paths.append(p)
    try:
        ensure_pdfs(paths)          # 先把这一批新的 Office 文件一次转完
    except Exception as e:
        print("[警告] Office 转换出错:", e)
    for p in paths:
        try:
            out.append(preflight(p))
        except OSError:
            pass
    out.sort(key=lambda x: -x["mtime"])
    return out


# ---------------- 打印队列 ----------------
QUEUE = queue.Queue()
JOBS = {}          # id -> job
JOBS_LOCK = threading.Lock()
_seq = [0]

DUPLEX = {"single": "simplex", "long": "duplexlong", "short": "duplexshort"}
# DMPAPER_* 编号。本机驱动的纸名叫 "A4(210x297mm)"，Sumatra 的 paper=A4 匹配不上
# （它只认完整同名或 "A4 " 带空格前缀），会退回"按文档尺寸自定义纸张"，驱动找不到
# 这种纸就落到 A3 —— 233x479mm 的图打成 A3 就是这么来的。paperkind 直接给编号。
PAPERKIND = {"A4": 9, "A3": 8, "A5": 11, "B4": 12, "B5": 13, "Letter": 1, "Legal": 5, "11x17": 3}


def build_settings(job):
    s = [f'{job["copies"]}x', DUPLEX[job["duplex"]]]
    s.append("color" if job["color"] else "monochrome")
    kind = PAPERKIND.get(job["paper"])
    s.append(f"paperkind={kind}" if kind else f'paper={job["paper"]}')
    # 不传 portrait/landscape：那两个参数只旋转内容、不改纸张方向，会和 SumatraPDF
    # 自带的自动旋转叠加成转两次（横向内容缩成 71% 塞在纵向纸里）。
    # 交给自动旋转：宽于高的页面自动转 90° 适配纸张，满版。
    s.append("fit")  # 适合纸张：按比例缩放到可打印区域，会缩也会放大
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


def worker():
    while True:
        job = QUEUE.get()
        job["state"] = "打印中"
        job["start"] = time.time()
        args = [SUMATRA, "-print-to", job["printer"],
                "-print-settings", job["settings"], "-silent", job["path"]]
        try:
            r = subprocess.run(args, capture_output=True, timeout=600,
                               creationflags=0x08000000)
            if r.returncode == 0:
                job["state"] = "完成"
                log_job(job)          # 只有真正送印成功才记账
            else:
                job["state"] = "失败"
                job["msg"] = (r.stderr or r.stdout or b"")[:300].decode("utf-8", "ignore").strip()
        except subprocess.TimeoutExpired:
            job["state"] = "失败"
            job["msg"] = "超时（>10 分钟）"
        except Exception as e:
            job["state"] = "失败"
            job["msg"] = str(e)
        job["end"] = time.time()


threading.Thread(target=worker, daemon=True).start()


def enqueue(items):
    """校验后入队，返回 [{id,name,state}]。"""
    ok = []
    for it in items:
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
        info = preflight(path)
        if info.get("error"):
            raise ValueError(f"{os.path.basename(path)}：{info['detail']}")
        pages = info["pages"]
        unit = unit_price(paper, color)
        _seq[0] += 1
        job = {"id": _seq[0], "name": os.path.basename(path), "path": info["pdf"],
               "printer": printer, "copies": copies, "color": color,
               "paper": paper, "duplex": duplex, "state": "排队中", "msg": "",
               "pages": pages, "unit": unit,
               "amount": round(unit * copies * pages, 2)}
        job["settings"] = build_settings(job)
        with JOBS_LOCK:
            JOBS[job["id"]] = job
        QUEUE.put(job)
        ok.append({"id": job["id"], "name": job["name"], "settings": job["settings"]})
    return ok


def recent_jobs():
    with JOBS_LOCK:
        js = sorted(JOBS.values(), key=lambda j: -j["id"])[:80]
        return [{k: v for k, v in j.items() if k != "path"} for j in js]


class Handler(BaseHTTPRequestHandler):
    server_version = "LinkPrint"

    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False)
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
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
        if urlparse(self.path).path != "/api/print":
            return self._send(404, {"error": "not found"})
        try:
            n = int(self.headers.get("Content-Length", 0))
            data = json.loads(self.rfile.read(n) or b"{}")
            return self._send(200, {"queued": enqueue(data.get("jobs", []))})
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
