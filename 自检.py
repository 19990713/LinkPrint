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
assert ps.build_settings(j) == "2x,collate,duplexshort,monochrome,paperkind=9,fit"
j.update(duplex="long", color=True)
assert ps.build_settings(j) == "2x,collate,duplexlong,color,paperkind=9,fit"
# 缩放四档：fit 不加 center，其余三档加（内容小于纸面时居中）
j.update(scale="shrink")
assert ps.build_settings(j) == "2x,collate,duplexlong,color,paperkind=9,shrink,center"
j.update(scale="noscale")
assert ps.build_settings(j) == "2x,collate,duplexlong,color,paperkind=9,noscale,center"
j.update(scale="custom", pct=85)
assert ps.build_settings(j) == "2x,collate,duplexlong,color,paperkind=9,noscale,center"
# 多份必须带 collate（逐份），单份不用
assert "collate" in ps.build_settings(j)
j.update(copies=1)
assert ps.build_settings(j) == "1x,duplexlong,color,paperkind=9,noscale,center"
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

# ---------------- 打印效果预览 ----------------
# 排版规则必须和 build_settings 发出去的一致，否则预览会撒谎
A4W, A4H = 210 * 72 / 25.4, 297 * 72 / 25.4
rot, k, _r, over = ps.print_layout(A4W, A4H, "A4", "fit")
assert rot == 0 and not over and abs(k - (210 - 8) / 210) < 1e-3, (k, over)   # A4 打 A4 = 96%
rot, k, _r, over = ps.print_layout(A4W, A4H, "A3", "fit")
assert abs(k - 289 / 210) < 0.01, k                    # A4 内容打 A3 会放大到 ~138%
rot, k, _r, over = ps.print_layout(A4H, A4W, "A4", "fit")
assert rot == 90, rot                                  # 横向页自动转 90°
rot, k, _r, over = ps.print_layout(A4W, A4H, "A4", "custom", 130)
assert abs(k - 1.3) < 1e-6 and over, (k, over)         # 130% 超出纸面，要标"会被裁"
rot, k, _r, over = ps.print_layout(A4W, A4H, "A4", "noscale")
assert abs(k - 1.0) < 1e-6 and not over, (k, over)     # 实际大小不算超纸
png, meta = ps.preview_png(files[0], "A4", "fit", 100, 1, True, 80)
assert png[:8] == b"\x89PNG\r\n\x1a\n", "应该是 PNG"
with fitz.open(files[0]) as _d:
    assert meta["pages"] == _d.page_count, meta
assert meta["pct"] == 96 and meta["rot"] == 0 and (meta["w"], meta["h"]) == (210, 297), meta
print(f"打印效果预览 OK  首屏 {meta['pct']}% {meta['w']}×{meta['h']}mm {len(png)//1024}KB")

# ---------------- 合并打印 ----------------
# 同参数的一组归成一份任务，准备阶段才真正拼 PDF（下单请求不等合并）
assert len(files) >= 2, "合并打印自检需要至少两份测试 PDF"


def _prep(f, **kw):
    d = {"name": os.path.basename(f), "names": [os.path.basename(f)], "srcs": [f],
         "printer": "D95", "copies": 1, "color": False, "paper": "A4",
         "duplex": "single", "scale": "fit", "pct": 100, "nup": 1}
    d.update(kw)
    d["unit"] = ps.unit_price(d["paper"], d["color"])
    return d


_a, _b = _prep(files[0]), _prep(files[1])
_c = _prep(files[0], paper="A3")            # 纸张不同，必须单独一组
groups = ps.group_jobs([_a, _b, _c])
assert len(groups) == 2, groups
_pa, _pb = ps.preflight(files[0])["pages"], ps.preflight(files[1])["pages"]
job = ps.prepare(ps._merge_group([g for g in groups if len(g) > 1][0]))
assert job["pages"] == _pa + _pb, job
assert job["tmp"] is True and os.path.isfile(job["path"]), job
assert abs(job["amount"] - (_a["unit"] * (_pa + _pb))) < 1e-9, job
with fitz.open(job["path"]) as _d:
    assert _d.page_count == job["pages"], (_d.page_count, job["pages"])
os.remove(job["path"])
print(f"合并打印 OK  {_pa}+{_pb}页 -> 一份 PDF，分成 {len(groups)} 组")


# ---------------- 页码范围 ----------------
assert ps.parse_pages("", 10) == list(range(1, 11))     # 空 = 全部
assert ps.parse_pages("", 0) is None                    # 页数未知时也返回「全部」
assert ps.parse_pages("3", 10) == [3]
assert ps.parse_pages("2-5", 10) == [2, 3, 4, 5]
assert ps.parse_pages("1-3,5", 10) == [1, 2, 3, 5]
assert ps.parse_pages("5-", 10) == [5, 6, 7, 8, 9, 10]  # 到末尾
assert ps.parse_pages("-2", 10) == [9]                  # -2 = 倒数第 2 页（同 SumatraPDF）
assert ps.parse_pages("-3--1", 10) == [8, 9, 10]        # 最后三页
assert ps.parse_pages("3-1", 10) == [1, 2, 3]           # 反向范围自动转正
assert ps.parse_pages("1，3 5、7", 10) == [1, 3, 5, 7]  # 中文逗号/顿号/空格都当分隔符
assert ps.parse_pages("1 - 3", 10) == [1, 2, 3]
assert ps.parse_pages("all", 4) == [1, 2, 3, 4]
assert ps.parse_pages("1-3", 0) == [1, 2, 3]            # 页数未知只校验语法
for bad in ("0", "11", "1-11", "abc", "1,,2-x", "-"):
    try:
        ps.parse_pages(bad, 10)
        raise AssertionError(f"应该报错：{bad}")
    except ValueError:
        pass
assert ps.pages_text([1, 2, 3, 5]) == "1-3,5"
assert ps.pages_text([1, 2]) == "1,2"
assert ps.pages_text([4]) == "4"
assert ps.pages_text(list(range(1, 11))) == "1-10"
# 页码排在参数串最前
assert ps.build_settings({"copies": 1, "duplex": "single", "color": False,
                          "paper": "A4", "scale": "fit", "range": "1-3"}) \
       == "1-3,1x,simplex,monochrome,paperkind=9,fit"
# 单份：4 页的文档只打前 2 页，页数/金额/参数串都跟着变
_p4 = ps.preflight(files[1])["pages"]
_r = _prep(files[1], specs=["1-2"])
job = ps.prepare(_r)
assert job["pages"] == 2 and job["range"] == "1,2", job
assert abs(job["amount"] - _r["unit"] * 2) < 1e-9, job
assert job["settings"].startswith("1,2,"), job["settings"]
# 多份合并：第二份的页码要按第一份的页数整体后移
job = ps.prepare(ps._merge_group([_prep(files[0], specs=["2-"]), _prep(files[1], specs=["1"])]))
want = [2, 3]                                   # files[0]=2页取第2页；files[1]取第1页 -> 合并后是第3页
assert job["pages"] == len(want), (job["pages"], want)
assert ps.parse_pages(job["range"], _pa + _p4) == want, job["range"]
assert abs(job["amount"] - _r["unit"] * len(want)) < 1e-9, job
os.remove(job["path"])
# 选满全部页时不该写页码，让 Sumatra 打整份
job = ps.prepare(_prep(files[1], specs=["1-"]))
assert job["range"] == "" and job["pages"] == _p4, job

# 网页下单那一层：pages_spec 一路传到 _validate，越界在下单时就挡掉（不用等准备阶段）
_old_hot, ps.HOT = ps.HOT, os.path.join(here, "测试文件")
try:
    _it = {"name": os.path.basename(files[1]), "printer": ps.PRINTERS[0], "copies": 1,
           "paper": "A4", "duplex": "single", "scale": "fit", "pages_spec": "1-2"}
    d = ps._validate(dict(_it))
    assert d["pages"] == 2 and d["range"] == "1,2" and d["specs"] == ["1-2"], d
    try:
        ps._validate(dict(_it, pages_spec="1-99"))
        raise AssertionError("越界页码应该在下单时就被挡掉")
    except ValueError:
        pass
    d = ps._validate(dict(_it, pages_spec=""))
    assert d["pages"] == _p4 and d["range"] == "", d
finally:
    ps.HOT = _old_hot
print(f"页码范围 OK  {_p4}页文档抽 1-2 页，合并时自动偏移")

# ---------------- 清空队列 ----------------
# 只删名单里的源文件，热文件夹里别的东西不碰；缓存 PDF 和合并临时文件一起清
import shutil
_old_hot, _old_merge = ps.HOT, ps.MERGE_DIR
hot2 = tempfile.mkdtemp(prefix="lp_clear_")
ps.HOT = hot2
ps.MERGE_DIR = os.path.join(hot2, "_m")
try:
    src = os.path.join(hot2, "要删.pdf")
    shutil.copyfile(files[0], src)
    cache = ps.conv_target(src)                 # 伪造一份转换缓存
    with open(cache, "wb") as f:
        f.write(b"cache")
    keep = os.path.join(hot2, "别动.txt")
    with open(keep, "w") as f:
        f.write("x")
    r = ps.clear_files(["要删.pdf"])
    assert r["removed"] == 1 and r["cache"] == 1 and not r["failed"], r
    assert not os.path.exists(src) and not os.path.exists(cache), "源文件和缓存都该没了"
    assert os.path.exists(keep), "非打印文件不该被删"
    # 名字只按 basename 解析，所以 "../x.pdf" 出不了热文件夹
    outside = os.path.join(os.path.dirname(hot2), "lp_outside.pdf")
    with open(outside, "wb") as f:
        f.write(b"x")
    try:
        assert ps.clear_files(["../lp_outside.pdf"])["removed"] == 0
        assert os.path.exists(outside), "不该删到热文件夹外面的文件"
    finally:
        os.remove(outside)
    print(f"清空队列 OK  删 {r['removed']} 个源文件 + {r['cache']} 个缓存 PDF")
finally:
    ps.HOT, ps.MERGE_DIR = _old_hot, _old_merge
    shutil.rmtree(hot2, ignore_errors=True)

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

# ---------------- 多页（N-up）拼版 ----------------
_nup = ""
try:
    # 网格映射：1/2/4/6/9/16 -> rows, cols
    assert ps.NUP_GRID == {1: (1, 1), 2: (1, 2), 4: (2, 2), 6: (2, 3), 9: (3, 3), 16: (4, 4)}
    # nup=1 原样返回可打印 PDF，不拼版
    assert ps.nup_pdf(files[0], 1, "A4") == ps.printable_pdf(files[0])
    # 4 页文档 nup=4 -> 1 张；nup=2 -> 2 张，横向纸（841.9x595.3pt）
    src4 = files[1]                       # 【A4】...双科.pdf 是 4 页
    n4 = ps.preflight(src4)["pages"]
    assert n4 == 4, n4
    p = ps.nup_pdf(src4, 4, "A4")
    with fitz.open(p) as d:
        assert d.page_count == 1, d.page_count
    os.remove(p)
    p = ps.nup_pdf(src4, 2, "A4")
    with fitz.open(p) as d:
        assert d.page_count == 2, d.page_count
        assert abs(d[0].rect.width - 841.9) < 1 and abs(d[0].rect.height - 595.3) < 1, \
            (d[0].rect.width, d[0].rect.height)
    os.remove(p)
    # 预览：nup=2 返回横向纸 meta（297x210mm）
    png, meta = ps.preview_png(src4, "A4", "fit", 100, 1, False, 120, nup=2)
    assert meta["w"] == 297 and meta["h"] == 210 and meta["pages"] == 2, meta
    # prepare：金额按纸张数计（4 页 nup=2 -> 2 张）
    job = _prep(src4, copies=3, nup=2)
    ps.prepare(job)
    assert job.get("sheets") == 2 and job["scale"] == "noscale", job
    assert abs(job["amount"] - 0.5 * 3 * 2) < 1e-9, job["amount"]
    assert "noscale" in job["settings"] and "2x" not in job["settings"], job["settings"]
    os.remove(job["path"])
    # 页码筛选 + nup：4 页只打前 2 页，nup=2 -> 1 张
    job = _prep(src4, nup=2, specs=["1-2"])
    ps.prepare(job)
    assert job.get("sheets") == 1, job
    os.remove(job["path"])
    print(f"多页拼版 OK  每张 2/4 页，金额按纸张数计，预览出横向纸")
finally:
    if _nup and os.path.isfile(_nup):
        os.remove(_nup)



# ---------------- GDI 送印引擎 ----------------
assert ps.ENGINE in ("gdi", "sumatra"), ps.ENGINE
# 纸型编号两张表必须一致（gdi_print 自己带一份，免得为了一个常量反向依赖）
if ps.gdi_print:
    assert ps.PAPERKIND == ps.gdi_print.PAPER_ID, (ps.PAPERKIND, ps.gdi_print.PAPER_ID)
    assert ps.gdi_print.available(), "GDI 送印不可用（缺 numpy？）"
    assert ps.gdi_print.parse_range("", 3) == [1, 2, 3]
    assert ps.gdi_print.parse_range("1-3,5", 10) == [1, 2, 3, 5]
    assert ps.gdi_print.parse_range("2-", 4) == [2, 3, 4]
    assert ps.gdi_print.parse_range("99", 3) == []        # 全越界 -> 空，送印前就会拦住
    print("GDI 引擎 OK  ", ps.ENGINE, "并行" if ps.PARALLEL else "串行",
          f"{len(ps.PRINT_QUEUES)} 条队列")

# 真实打印机的可打印区域（只读，不出纸）。离线/没装的跳过，不当失败
if ps.gdi_print and ps.PRINTERS:
    for _p in ps.PRINTERS:
        try:
            _c = ps.gdi_print.printer_caps(_p)
            assert _c["w"] > 0 and _c["h"] > 0 and _c["dpi_x"] > 0, _c
            assert _c["w"] < _c["phys_w"] and _c["h"] < _c["phys_h"], _c   # 可打印区必须小于纸
            print(f"  {_p}  可打印 {_c['w']}×{_c['h']}px @{_c['dpi_x']}dpi  "
                  f"纸 {_c['paper_mm'][0]}×{_c['paper_mm'][1]}mm")
        except Exception as _e:
            print(f"  [跳过] {_p}：{_e}")

# 送印队列：任务只会进它那台打印机的队列
if ps.PRINTER_Q:
    _p0 = ps.PRINTERS[0]
    assert ps.PRINTER_Q[_p0] in ps.PRINT_QUEUES
    if ps.PARALLEL and len(ps.PRINTERS) > 1:
        assert ps.PRINTER_Q[ps.PRINTERS[0]] is not ps.PRINTER_Q[ps.PRINTERS[1]], "并行模式队列不该共用"
    else:
        assert len(ps.PRINT_QUEUES) == 1, "串行模式应该只有一条队列"
    print("送印队列 OK")
