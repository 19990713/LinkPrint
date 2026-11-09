# -*- coding: utf-8 -*-
"""LinkPrint 图标生成器（第五版）

输出 web/icons/icon.svg（完整版）与 icon-small.svg（无字标小尺寸版）。
坐标是「参考图白砖裁成 512x512 后的画布坐标」，再统一过一个 <g> 变换映射到真正的白砖坐标
（量测用的旧裁切框比白砖多留了下边约 15px，所以内容做了 -0.27% / +1.35% 的补偿）。

量测值（旧裁切框坐标）：
  白砖   512x512, rx 112
  机身   x 58..454, y 193..335, 圆角 39
  白缝   直段 y 217.6..227.4、x 98..412；两端顺肩线向下垂落成细尾，左尾尖 (81,238)、右尾尖 (427,230)
  纸     左 123 / 折痕起点 308 / 右 375，顶 85.5，描边 13.5，圆角 20
  折角   灰块 左竖边 x308.5（y85.5..136.5）、底边 y152.4、左下角小抹角
  指示灯 (411,254) r 10
  文字   Trebuchet MS Bold，ink 156..321 / 135..203，等比例排版（横向 HSCALE 1.045），字距反解
  彩线   胶囊线宽 18.5，中心/角度/长度见 CAPS（HSV 掩膜 + PCA 实测）
  CMYK   梯形 (157,288)-(356,288)-(389,386)-(108,386)，渐变按参考图逐点采样重写
"""

from fontTools.ttLib import TTFont
from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.pens.transformPen import TransformPen
from fontTools.pens.boundsPen import BoundsPen
from fontTools.misc.transform import Transform

# ---------------- 文字：等比例排版（sx≈sy），字距反解到目标 ink 宽度 ----------------
FONT, TEXT = r"C:/Windows/Fonts/trebucbd.ttf", "Link"
INK = (156.0, 135.0, 321.0, 203.0)    # 目标 ink 外框（含同色描边外扩）
STROKE = 3.2                          # 同色描边加字重（参考图笔画 14~15px）
HSCALE = 1.045                        # 参考图字形比 Trebuchet 略宽，横向 +4.5%（<5% 不显变形）

f = TTFont(FONT)
gs, cmap, upem = f.getGlyphSet(), f.getBestCmap(), f["head"].unitsPerEm


def layout(track, text=TEXT):
    """给定字距返回 (x0, x1, y0, y1, placed)，单位 em。"""
    x_em = 0.0
    a0 = b0 = 1e9
    a1 = b1 = -1e9
    placed = []
    for ch in text:
        g = cmap[ord(ch)]
        bp = BoundsPen(gs)
        gs[g].draw(bp)
        if bp.bounds:
            bx0, by0, bx1, by1 = bp.bounds
            a0 = min(a0, x_em + bx0 / upem)
            a1 = max(a1, x_em + bx1 / upem)
            b0 = min(b0, by0 / upem)
            b1 = max(b1, by1 / upem)
        placed.append((x_em, g))
        x_em += gs[g].width / upem + track / 100.0
    return a0, a1, b0, b1, placed


pad = STROKE / 2.0
tx0, ty0, tx1, ty1 = INK[0] + pad, INK[1] + pad, INK[2] - pad, INK[3] - pad
w_t, h_t = tx1 - tx0, ty1 - ty0

_, _, ny0, ny1, _ = layout(0.0)
sy0 = h_t / (ny1 - ny0)          # 由字高定比例
sx0 = sy0 * HSCALE               # 横向按 HSCALE 微扩

lo, hi = -30.0, 20.0
for _ in range(60):
    tr = (lo + hi) / 2.0
    a0, a1, _, _, _ = layout(tr)
    if (a1 - a0) * sx0 > w_t:    # 太宽 → 收紧字距
        hi = tr
    else:
        lo = tr
TRACK = (lo + hi) / 2.0

x0, x1, y0, y1, placed = layout(TRACK)
sx = w_t / (x1 - x0)
sy = h_t / (y1 - y0)
ox, oy = tx0 - x0 * sx, ty1 + y0 * sy

def num(v):
    return ("%.2f" % v).rstrip("0").rstrip(".")

d_text = ""
for px, g in placed:
    sp = SVGPathPen(gs, ntos=num)
    gs[g].draw(TransformPen(sp, Transform(sx / upem, 0, 0, -sy / upem, ox + px * sx, oy)))
    d_text += sp.getCommands()

print("text  sx=%.2f sy=%.2f 横压=%.3f track=%.2f  ink %.1f x %.1f"
      % (sx, sy, sx / sy, TRACK, w_t, h_t))

# ---------------- 几何 ----------------
BODY = (58.0, 193.0, 454.0, 335.0, 39.0)
PAPER_L, PAPER_R, PAPER_T = 123.0, 375.0, 85.5
PAPER_SW, PAPER_RAD, PAPER_BOT = 13.5, 20.0, 227.4
CREASE = (308.0, 85.5, 375.0, 150.5)
SEAM_T, SEAM_B = 217.6, 226.7
DOT = (411.0, 254.0, 10.0)
SHEET = "M144 287H369L414 397Q416 403 410 403H97Q91 403 93 397Z"
CMYK = "M157 288H356L389 386Q391 390 386 390H108Q103 390 105 386Z"

# 白缝：中段 y 217.6..226.7 等宽直条；两端不收成针尖，而是顺肩线向下垂落成带圆头的细尾。
# 量测（参考图 512 坐标）：左尾尖 (82,236)、右尾尖 (427.5,230.5)，直段 x 98..412。
SEAM = """M83.4 236.4
          C83.9 226.6 92.0 217.6 108.0 217.6
          H412
          C419.6 219.2 425.8 224.8 427.2 230.8
          A1.1 1.1 0 0 1 426.0 232.3
          C423.8 230.0 419.2 227.6 415.0 227.6
          H109
          C96.0 227.6 86.4 229.4 84.2 237.6
          A1.7 1.7 0 0 0 83.4 236.4
          Z"""

# 折角灰块：左竖边 x308.5（y 85.5..136.5），底边 y152.3，左下角只做很小的抹角。
FOLD = "M308.5 136.5V85.5L376 152.4H321C316.2 152.4 310.6 148.4 308.5 136.5Z"

CAPS = [  # 中心, 方向(单位向量), 胶囊总长(含圆头), 颜色, 线宽
    ((400.7, 104.5), (-0.300, 0.954), 50.9, "#019bfd", 18.5),
    ((434.6, 125.9), (-0.756, 0.654), 55.5, "#fc2074", 18.5),
    ((449.0, 161.8), (-0.976, 0.213), 49.9, "#fec105", 18.5),
]

def capsule(c, u, total, w):
    half = (total - w) / 2.0
    return "M%.1f %.1fL%.1f %.1f" % (
        c[0] + u[0] * half, c[1] + u[1] * half,
        c[0] - u[0] * half, c[1] - u[1] * half)

bx0, by0, bx1, by1, br = BODY

CAPS_SVG = "".join('    <path d="%s" stroke="%s" stroke-width="%.1f"/>\n'
                   % (capsule(c, u, L, w), col, w) for c, u, L, col, w in CAPS)


def build(with_text=True, shadow_op=0.70):
    text = ('  <!-- Link -->\n  <path d="%s" fill="#1a1a20" stroke="#1a1a20" stroke-width="%s"'
            '\n        stroke-linejoin="round"/>\n' % (d_text, STROKE)) if with_text else ""
    return f"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" width="512" height="512">
  <defs>
    <linearGradient id="cmyk" x1="0" y1="0" x2="1" y2="0">
      <stop offset="0" stop-color="#00baff"/><stop offset="0.14" stop-color="#01a1fd"/>
      <stop offset="0.30" stop-color="#3e7efa"/><stop offset="0.37" stop-color="#6a63f1"/>
      <stop offset="0.45" stop-color="#ce43cc"/><stop offset="0.51" stop-color="#f037af"/>
      <stop offset="0.56" stop-color="#fd3b94"/><stop offset="0.66" stop-color="#fe635d"/>
      <stop offset="0.71" stop-color="#fe8245"/><stop offset="0.76" stop-color="#fda131"/>
      <stop offset="0.81" stop-color="#ffb82a"/><stop offset="0.91" stop-color="#fed530"/>
      <stop offset="1" stop-color="#ffe63c"/>
    </linearGradient>
    <linearGradient id="fold" x1="0.55" y1="0" x2="0.2" y2="1">
      <stop offset="0" stop-color="#c7ccd7"/><stop offset="0.62" stop-color="#d6dae4"/>
      <stop offset="1" stop-color="#edeef4"/>
    </linearGradient>
    <linearGradient id="lip" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0" stop-color="#000000" stop-opacity="0.45"/>
      <stop offset="1" stop-color="#000000" stop-opacity="0"/>
    </linearGradient>
    <filter id="sh" x="-30%" y="-25%" width="160%" height="170%">
      <feDropShadow dx="0" dy="6" stdDeviation="20" flood-color="#22305c" flood-opacity="{shadow_op}"/>
    </filter>
    <clipPath id="sheetClip"><path d="{SHEET}"/></clipPath>
  </defs>

  <rect width="512" height="512" rx="112" fill="#ffffff"/>
  <rect x="0.8" y="0.8" width="510.4" height="510.4" rx="111" fill="none"
        stroke="#f1f3f7" stroke-width="1.4"/>

  <!-- 内容整体坐标校正：量测是在“旧裁切框”里做的，旧框比白砖多留了下边 ~15px，
       这里把内容映射回真正的白砖坐标（横向 -0.27%，纵向 +1.35%） -->
  <g transform="translate(0.959 0.487) scale(0.99726 1.01346)">

  <!-- 机身 -->
  <path d="M{bx0+br} {by0}H{bx1-br}A{br} {br} 0 0 1 {bx1} {by0+br}V{by1-br}
           A{br} {br} 0 0 1 {bx1-br} {by1}H{bx0+br}A{br} {br} 0 0 1 {bx0} {by1-br}
           V{by0+br}A{br} {br} 0 0 1 {bx0+br} {by0}Z" fill="#1a1a20"/>

  <!-- 上盖与机身之间的白缝：中段等宽，两端顺肩线向下垂落成带圆头的细尾 -->
  <path d="{SEAM}" fill="#ffffff"/>

  <!-- 指示灯 -->
  <circle cx="{DOT[0]}" cy="{DOT[1]}" r="{DOT[2]}" fill="#ffffff"/>

  <!-- 纸张白底：下缘与白缝下沿齐平，白色连成一片 -->
  <path d="M{PAPER_L} {PAPER_BOT}V{PAPER_T+PAPER_RAD}
           A{PAPER_RAD} {PAPER_RAD} 0 0 1 {PAPER_L+PAPER_RAD} {PAPER_T}
           H{CREASE[0]}L{CREASE[2]} {CREASE[3]}V{PAPER_BOT}Z" fill="#ffffff"/>

  <!-- 折角：压在折痕内侧，先画，随后纸的描边会盖住它的斜边，接缝不留白 -->
  <path d="{FOLD}" fill="url(#fold)"/>

  <!-- 纸的描边：左、顶、折痕、右（底边不描边，止于白缝上沿） -->
  <path d="M{PAPER_L} {SEAM_T}V{PAPER_T+PAPER_RAD}
           A{PAPER_RAD} {PAPER_RAD} 0 0 1 {PAPER_L+PAPER_RAD} {PAPER_T}
           H{CREASE[0]}L{CREASE[2]} {CREASE[3]}V{SEAM_T}"
        fill="none" stroke="#1a1a20" stroke-width="{PAPER_SW}"
        stroke-linejoin="round" stroke-linecap="butt"/>

{text}  <!-- CMYK 出纸 -->
  <g filter="url(#sh)"><path d="{SHEET}" fill="#ffffff"/></g>
  <path d="{CMYK}" fill="url(#cmyk)"/>
  <rect x="60" y="286" width="392" height="8" fill="url(#lip)" clip-path="url(#sheetClip)"/>

  <!-- 右上三条彩线 -->
  <g stroke-linecap="round" fill="none">
{CAPS_SVG}  </g>

  </g>
</svg>
"""


SVG = build(True, 0.70)
open("web/icons/icon.svg", "w", encoding="utf-8").write(SVG)
print("written web/icons/icon.svg", len(SVG), "bytes")

SMALL = build(False, 0.55)
open("web/icons/icon-small.svg", "w", encoding="utf-8").write(SMALL)
print("written web/icons/icon-small.svg", len(SMALL), "bytes")
