# -*- coding: utf-8 -*-
"""把 compute_report 产出的报告文字排成白底黑字 PNG（微信可直接查看）。

用法：
  python render_report_image.py <positions.json> [--out 路径] [--date YYYY-MM-DD]
  python render_report_image.py --text-file 报告.txt [--out 路径]

不指定 --out 时输出到 图表数据/机构持仓报告_YYYYMMDD.png。
字体用微软雅黑（C:/Windows/Fonts/msyh*.ttc），缺字体时退回系统默认中文字体。
"""
import argparse
import os
import re
import sys
from datetime import date, datetime

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    sys.exit("需要 Pillow：pip install pillow")

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.environ.get("INST_HOLDINGS_HOME") or os.path.abspath(
    os.path.join(HERE, "..", "..", "..", ".."))
OUT_DIR = os.environ.get("INST_HOLDINGS_OUT") or os.path.join(ROOT, "图表数据")

W = 1240
MARGIN_L = 78
MARGIN_R = 78
MARGIN_T = 74
MARGIN_B = 66

FONT_R = "C:/Windows/Fonts/msyh.ttc"
FONT_B = "C:/Windows/Fonts/msyhbd.ttc"
FALLBACK_R = ["C:/Windows/Fonts/Deng.ttf", "C:/Windows/Fonts/simhei.ttf"]
FALLBACK_B = ["C:/Windows/Fonts/Dengb.ttf", "C:/Windows/Fonts/simhei.ttf"]

BLACK = (0, 0, 0)
WHITE = (255, 255, 255)
RULE = (232, 232, 232)
GREY = (90, 90, 90)

SZ_TITLE = 46
SZ_H = 36
SZ_B = 31
SZ_F = 22

GAP_TITLE = 46
GAP_H = 34
GAP_B = 15
GAP_SEC = 30


def _pick(primary, fallbacks):
    for p in [primary] + fallbacks:
        if os.path.exists(p):
            return p
    return None


FR_PATH = _pick(FONT_R, FALLBACK_R)
FB_PATH = _pick(FONT_B, FALLBACK_B) or FR_PATH


def font(path, size):
    if path:
        return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def parse_report(text):
    """把报告文字拆成 (kind, text) 序列。"""
    items = []
    lines = [ln.rstrip() for ln in text.strip().splitlines() if ln.strip()]
    if not lines:
        sys.exit("报告文字为空")
    items.append(("title", lines[0]))
    items.append(("rule", ""))
    for ln in lines[1:]:
        if re.match(r"^[一二三四五六七八九十]、", ln):
            items.append(("head", ln))
        else:
            items.append(("body", ln))
    return items


def layout(items, footer):
    fr = font(FR_PATH, SZ_B)
    probe = ImageDraw.Draw(Image.new("RGB", (10, 10)))
    max_w = W - MARGIN_L - MARGIN_R
    y = MARGIN_T
    out = []
    first_head = True
    for kind, text in items:
        if kind == "head" and not first_head:
            y += 24
        if kind == "head":
            first_head = False
        if kind == "title":
            out.append((kind, text, y, font(FB_PATH, SZ_TITLE)))
            y += SZ_TITLE + GAP_TITLE
        elif kind == "rule":
            out.append((kind, text, y, None))
            y += 24
        elif kind == "head":
            out.append((kind, text, y, font(FB_PATH, SZ_H)))
            y += SZ_H + GAP_H
        else:
            f = fr
            if probe.textlength(text, font=f) <= max_w:
                out.append((kind, text, y, f))
                y += SZ_B + GAP_B
            else:
                cur = ""
                for ch in text:
                    if probe.textlength(cur + ch, font=f) > max_w:
                        out.append((kind, cur, y, f))
                        y += SZ_B + GAP_B
                        cur = ch
                    else:
                        cur += ch
                if cur:
                    out.append((kind, cur, y, f))
                    y += SZ_B + GAP_B
    y += GAP_SEC
    out.append(("foot", footer, y, font(FR_PATH, SZ_F)))
    y += SZ_F + 6
    return out, y


def render(items, footer, out_path):
    laid, end_y = layout(items, footer)
    H = int(end_y + MARGIN_B)
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    for kind, text, y, f in laid:
        if kind == "rule":
            d.line([(MARGIN_L, y), (W - MARGIN_R, y)], fill=RULE, width=2)
        elif kind == "foot":
            d.text((MARGIN_L, y), text, font=f, fill=GREY)
        else:
            d.text((MARGIN_L, y), text, font=f, fill=BLACK)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    img.save(out_path, "PNG")
    return img.size


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("json", nargs="?", help="positions_auto.json")
    ap.add_argument("--text-file", help="直接给报告文字文件（跳过 json）")
    ap.add_argument("--out", default=None, help="输出 PNG 路径")
    ap.add_argument("--date", default=None, help="数据日 YYYY-MM-DD，用于文件名与页脚")
    a = ap.parse_args()

    if a.text_file:
        text = open(a.text_file, encoding="utf-8").read()
        d = a.date
        if not d:
            m = re.search(r"(\d{1,2})\s*月\s*(\d{1,2})\s*日", text)
            d = date.today().isoformat()
    elif a.json:
        import json
        sys.path.insert(0, HERE)
        import compute_report
        data = json.load(open(a.json, encoding="utf-8"))
        text = compute_report.build(data)
        d = a.date or data.get("date")
    else:
        sys.exit("需要 json 或 --text-file")

    dd = datetime.strptime(str(d), "%Y-%m-%d").date() if d else date.today()
    stamp = dd.strftime("%Y%m%d")
    footer = "数据来源：东方财富龙虎榜 · 数据日 {}".format(dd.strftime("%Y-%m-%d"))

    out = a.out or os.path.join(OUT_DIR, "机构持仓报告_{}.png".format(stamp))
    items = parse_report(text)
    size = render(items, footer, out)
    print(text)
    print("\n✅ 报告图：{}（{}×{}）".format(out, size[0], size[1]))


if __name__ == "__main__":
    main()
