# -*- coding: utf-8 -*-
"""由 icon.svg / icon-small.svg 生成全套位图与 favicon.ico（可重复执行）

用法：python web/icons/build-icons.py
     （需先跑 web/icons/gen-icon.py 生成 SVG）

约定：
  · 16/32 像素用无文字的 icon-small.svg（带字标在这个尺寸只会糊成一团）
  · 48 及以上用完整的 icon.svg
  · favicon.ico 打包 16,24,32,48,64,128,256 七帧
"""
import os
import subprocess

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
CHROME = r"C:/Program Files/Google/Chrome/Application/chrome.exe"

MASTER_SMALL = 16, 32
MASTER_MAIN = 48, 64, 128, 256, 512


def render(svg_path, out_png, px, scale=2):
    """用 Chrome 把 SVG 渲染成 px*scale 的 PNG（内联进 HTML，规避 file:// 强缓存）。"""
    svg = open(svg_path, encoding="utf-8").read()
    html = ('<!doctype html><meta charset="utf-8">'
            '<style>html,body{margin:0;background:transparent}'
            'svg{display:block;width:%dpx;height:%dpx}</style>%s' % (px, px, svg))
    tmp = os.path.join(HERE, "_tmp_render.html")
    open(tmp, "w", encoding="utf-8").write(html)
    subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars",
                    "--force-device-scale-factor=%d" % scale,
                    "--window-size=%d,%d" % (px, px),
                    "--default-background-color=00000000",
                    "--screenshot=%s" % os.path.abspath(out_png),
                    "file:///" + tmp.replace("\\", "/")],
                   capture_output=True, check=False)
    os.remove(tmp)


def write_ico(frames, path):
    """frames: [(size, PIL.Image)] 顺序即尺寸顺序；每帧存 PNG（Vista+ 支持）。"""
    import io
    import struct

    blobs = []
    for im in frames:
        buf = io.BytesIO()
        im.save(buf, format="PNG", optimize=True)
        blobs.append(buf.getvalue())

    n = len(frames)
    header = struct.pack("<HHH", 0, 1, n)
    offset = 6 + 16 * n
    entries = b""
    for im, blob in zip(frames, blobs):
        w, h = im.size
        entries += struct.pack("<BBBBHHII",
                               0 if w >= 256 else w,
                               0 if h >= 256 else h,
                               0, 0, 1, 32, len(blob), offset)
        offset += len(blob)
    with open(path, "wb") as fh:
        fh.write(header + entries + b"".join(blobs))


def main():
    main_svg = os.path.join(HERE, "icon.svg")
    small_svg = os.path.join(HERE, "icon-small.svg")

    big = os.path.join(HERE, "_master_main.png")
    sml = os.path.join(HERE, "_master_small.png")
    render(main_svg, big, 512, 2)
    render(small_svg, sml, 512, 2)
    M = Image.open(big).convert("RGBA")
    S = Image.open(sml).convert("RGBA")
    M.resize((1024, 1024), Image.LANCZOS).save(os.path.join(HERE, "icon-source-1024.png"))

    frames = {}
    for px in MASTER_MAIN:
        M.resize((px, px), Image.LANCZOS).save(os.path.join(HERE, "icon-%d.png" % px))
        frames[px] = M.resize((px, px), Image.LANCZOS)
    for px in MASTER_SMALL:
        S.resize((px, px), Image.LANCZOS).save(os.path.join(HERE, "icon-%d.png" % px))
        frames[px] = S.resize((px, px), Image.LANCZOS)

    # favicon.ico：小尺寸用无文字版，大尺寸用完整版；手工拼 ICO 容器（Pillow 的 sizes= 会漏帧）
    ico_sizes = [16, 24, 32, 48, 64, 128, 256]
    ico_frames = []
    for px in ico_sizes:
        src = S if px <= 32 else M
        ico_frames.append(src.resize((px, px), Image.LANCZOS))
    write_ico(ico_frames, os.path.join(HERE, "favicon.ico"))

    os.remove(big)
    os.remove(sml)
    print("生成完成：", ", ".join("icon-%d.png" % p for p in sorted(frames)))
    print("favicon.ico 帧：", ico_sizes)


if __name__ == "__main__":
    main()
