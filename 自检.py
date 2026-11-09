# -*- coding: utf-8 -*-
"""自检：纸型识别、打印参数字符串、真实 PDF 预检。python 自检.py"""
import glob
import os

import print_system as ps
from print_system import fitz

# 纸型识别
assert ps.classify(210, 297) == "A4"
assert ps.classify(297, 210) == "A4"      # 横向也是 A4
assert ps.classify(297, 420) == "A3"
assert ps.classify(420, 297) == "A3"
assert ps.classify(209, 296) == "A4"      # 1mm 误差
assert ps.classify(213, 303) == "A4"      # 带出血的 A4
assert ps.classify(200, 300) is None      # 非标
print("纸型识别 OK")

# 打印参数串（不能带 portrait/landscape，否则和自动旋转叠加转两次；缩放默认 fit 适合纸张）
j = {"copies": 2, "duplex": "short", "color": False, "paper": "A4"}
assert ps.build_settings(j) == "2x,duplexshort,monochrome,paperkind=9,fit"
j.update(duplex="long", color=True)
assert ps.build_settings(j) == "2x,duplexlong,color,paperkind=9,fit"
# 缩放四档：fit 不加 center，其余三档加（内容小于纸面时居中）
j.update(scale="shrink")
assert ps.build_settings(j) == "2x,duplexlong,color,paperkind=9,shrink,center"
j.update(scale="noscale")
assert ps.build_settings(j) == "2x,duplexlong,color,paperkind=9,noscale,center"
j.update(scale="custom", pct=85)
assert ps.build_settings(j) == "2x,duplexlong,color,paperkind=9,noscale,center"
print("打印参数 OK")

# 计费：A4 黑白 0.5 / 彩色 2，A3 翻倍
assert ps.unit_price("A4", False) == 0.5
assert ps.unit_price("A4", True) == 2.0
assert ps.unit_price("A3", False) == 1.0
assert ps.unit_price("A3", True) == 4.0

import tempfile
tmp = os.path.join(tempfile.gettempdir(), "lp_log_test.csv")
if os.path.exists(tmp):
    os.remove(tmp)
ps.LOG_PATH = tmp
ps.log_job({"name": "a.pdf", "pages": 3, "copies": 2, "paper": "A4",
            "color": False, "printer": "D95", "unit": 0.5, "amount": 3.0})
ps.log_job({"name": "b.pdf", "pages": 1, "copies": 1, "paper": "A3",
            "color": True, "printer": "C7580", "unit": 4.0, "amount": 4.0})
d = ps.read_log()
assert d["count"] == 2 and abs(d["total"] - 7.0) < 1e-9, d
assert d["today"] == 7.0, d
print("计费 OK  合计", d["total"], "首条", d["rows"][0]["文件"], d["rows"][0]["金额元"])
os.remove(tmp)

# 真实 PDF 预检
here = os.path.dirname(os.path.abspath(__file__))
files = glob.glob(os.path.join(here, "测试文件", "*.pdf"))
assert files, "没有找到测试 PDF"
for f in files:
    i = ps.preflight(f)
    assert i["pages"] > 0, f
    assert i["paper"] in ("A4", "A3", "A5", "B5", "Letter"), f
    print(f"{i['pages']:>4}页  {i['size']:<4} {i['orientation']}  "
          f"{i['mb']}MB  {i['detail']}  <- {i['name']}")
print("预检 OK")

# ---------------- Office 转 PDF ----------------
assert ps.ext_of("a.DOCX") == ".docx"
assert ".docx" in ps.OFFICE_EXT and ".xlsm" in ps.OFFICE_EXT and ".dps" in ps.OFFICE_EXT
assert ".pdf" not in ps.OFFICE_EXT
assert ps.printable_pdf(files[0]) == files[0], "PDF 应该原样返回，不该转"

# 真实 Word：转换 → 缓存 → 第二次命中缓存不再转
doc = os.path.join(here, "热文件夹", "声明.doc")
if os.path.isfile(doc):
    p1 = ps.printable_pdf(doc)
    assert os.path.isfile(p1) and os.path.getsize(p1) > 0, p1
    assert ps.printable_pdf(doc) == p1, "第二次应命中缓存"
    assert ps.preflight(doc)["pages"] > 0
    print("Word 转换 OK  ->", os.path.basename(p1))

# 转换失败必须抛错（否则会拿个空文件去打印、还照原文件页数计费）
_bak = ps.run_convert
ps.run_convert = lambda paths: {}
try:
    bad = os.path.join(tempfile.gettempdir(), "lp_fail_test.docx")
    open(bad, "wb").close()
    try:
        ps.printable_pdf(bad)
        raise AssertionError("转换失败时应该抛异常")
    except RuntimeError:
        pass
    finally:
        os.remove(bad)
finally:
    ps.run_convert = _bak
print("转换失败拦截 OK")

# ---------------- 图片转 PDF ----------------
assert ".png" in ps.IMG_EXT and ".png" not in ps.OFFICE_EXT
png = os.path.join(tempfile.gettempdir(), "lp_img_test.png")
pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 600, 450), 0)
pix.clear_with(200)
pix.save(png)
ip = ""
try:
    ip = ps.printable_pdf(png)
    assert ip != png and ip.endswith(".pdf") and os.path.getsize(ip) > 0, ip
    assert ps.printable_pdf(png) == ip, "第二次应命中缓存"
    ii = ps.preflight(png)
    assert ii["pages"] == 1 and ii["kind"] == "图片" and not ii.get("error"), ii
    print(f"图片转 PDF OK  {ii['size']} {ii['orientation']}  {ii['detail']}")
finally:
    os.remove(png)
    if os.path.isfile(ip):
        os.remove(ip)

# ---------------- 自定义比例缩放 ----------------
sp = ""
try:
    src_pdf = files[0]
    try:
        ps.scaled_pdf(src_pdf, 5)
        raise AssertionError("比例超范围应该报错")
    except ValueError:
        pass
    sp = ps.scaled_pdf(src_pdf, 85)
    a, b = fitz.open(src_pdf), fitz.open(sp)
    try:
        assert len(b) == len(a), "缩放后页数不能变"
        for i in range(len(a)):
            assert abs(b[i].rect.width / a[i].rect.width - 0.85) < 1e-3, i
    finally:
        a.close()
        b.close()
    assert ps.scaled_pdf(src_pdf, 85) == sp, "第二次应命中缩放缓存"
    print(f"自定义比例 OK  85% -> {os.path.basename(sp)}")
finally:
    if sp and os.path.isfile(sp):
        os.remove(sp)

# ---------------- 切换热文件夹 ----------------
import json
orig_cfg = json.load(open(ps.CFG_PATH, encoding="utf-8"))
tmpdir = tempfile.mkdtemp(prefix="lp_hot_")
assert ps.set_hotfolder(tmpdir) == os.path.abspath(tmpdir)
assert json.load(open(ps.CFG_PATH, encoding="utf-8"))["hotfolder"] == os.path.abspath(tmpdir)
try:
    ps.set_hotfolder(os.path.join(tmpdir, "不存在"))
    raise AssertionError("不存在的目录应该被拒")
except ValueError:
    pass
# 还原成测试前的值
old = orig_cfg["hotfolder"]
ps.CFG["hotfolder"] = old
ps.HOT = old if os.path.isabs(old) else os.path.join(ps.BASE, old)
json.dump(ps.CFG, open(ps.CFG_PATH, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
os.rmdir(tmpdir)
print("切换热文件夹 OK，已还原为", ps.HOT)

# 选择框本身：把弹窗换成固定返回值，只验证 Tk 在本进程能起能关（服务里是在 HTTP 线程里起）
from tkinter import filedialog
_fb = filedialog.askdirectory
filedialog.askdirectory = lambda **kw: "X"
try:
    assert ps.pick_folder() == "X"
finally:
    filedialog.askdirectory = _fb
print("文件夹选择框 OK")
