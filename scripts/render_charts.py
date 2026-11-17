#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""机构持仓合并图（研报简洁风 · v2，含豆粕05）

左豆粕01、中豆粕05、右菜粕01，取 Sheet1 最近 N 个交易日，直接读表渲染成同一张 PNG。
只读工作簿：不写表、不改图表对象，Excel 开着也能跑（读的是磁盘文件）。
K–P 列（豆粕05）若整列为空，自动退回两图（豆粕01 / 菜粕01）。

用法：
  python render_charts.py                       # 最近 5 日，文件名取表内最新日期
  python render_charts.py --date 2026-11-04     # 指定文件名日期
  python render_charts.py --days 10             # 周报用 10 日窗口
  python render_charts.py --no-meal05           # 强制只画豆粕01 + 菜粕01

输出：<out-dir>/机构持仓_YYYYMMDD.png
"""
import argparse
import os
import sys
from datetime import datetime

XLSM = os.environ.get("INST_HOLDINGS_XLSM") or r"D:/WorkBuddyWorkSpace/鲁信工作/机构持仓表.xlsm"
OUT_DIR = os.environ.get("INST_HOLDINGS_OUT") or r"D:/WorkBuddyWorkSpace/鲁信工作/图表数据"

MEAL = ["高盛", "摩根", "中粮", "国泰君安", "东证", "中信"]
RAPE = ["高盛", "摩根", "中粮"]
MEAL_COLS = dict(zip(MEAL, [2, 3, 4, 5, 6, 7]))      # B..G   (H 是合计列，不画)
MEAL05_COLS = dict(zip(MEAL, [11, 12, 13, 14, 15, 16]))  # K..P  (Q 是合计列，不画)
RAPE_COLS = dict(zip(RAPE, [29, 30, 31]))            # AC..AE (AF 是合计列)

PAL = {"高盛": "#E4572E", "摩根": "#2E7EBF", "中粮": "#333333",
       "国泰君安": "#D99A2B", "东证": "#7C5CFF", "中信": "#0FA37F"}


def die(msg, code=2):
    print("❌ " + msg, file=sys.stderr)
    sys.exit(code)


def load(path, n_days):
    try:
        from openpyxl import load_workbook
        from openpyxl.utils.datetime import from_excel
    except ImportError:
        die("缺 openpyxl，先 pip install openpyxl")
    if not os.path.exists(path):
        die("找不到工作簿：{}".format(path))
    wb = load_workbook(path, read_only=True, data_only=True)
    if "Sheet1" not in wb.sheetnames:
        die("工作簿里没有 Sheet1")
    ws = wb["Sheet1"]
    rows = []
    r = 2
    while len(rows) < n_days:
        a = ws.cell(row=r, column=1).value
        if a is None:
            break
        d = a.date() if isinstance(a, datetime) else from_excel(a).date()
        rows.append((d,
                     {n: ws.cell(row=r, column=c).value for n, c in MEAL_COLS.items()},
                     {n: ws.cell(row=r, column=c).value for n, c in MEAL05_COLS.items()},
                     {n: ws.cell(row=r, column=c).value for n, c in RAPE_COLS.items()}))
        r += 1
    wb.close()
    if not rows:
        die("Sheet1 第 2 行起没有数据")
    rows.reverse()                      # 表内倒序 → 图要正序
    dates = [x[0] for x in rows]
    meal = {n: [float(x[1][n] or 0) for x in rows] for n in MEAL}
    meal05 = {n: [float(x[2][n] or 0) for x in rows] for n in MEAL}
    rape = {n: [float(x[3][n] or 0) for x in rows] for n in RAPE}
    return dates, meal, meal05, rape


def has_data(series):
    return any(abs(v) > 0 for vals in series.values() for v in vals)


def build(dates, meal, meal05, rape, use_meal05=True):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.lines import Line2D
    from matplotlib.ticker import FuncFormatter

    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "Source Han Sans CN", "SimHei"]
    plt.rcParams["axes.unicode_minus"] = False

    def K(v, _=None):
        return "{:,.0f}".format(v)

    panels = [(MEAL, meal, "豆粕 2701（主力）")]
    if use_meal05:
        panels.append((MEAL, meal05, "豆粕 2705"))
    panels.append((RAPE, rape, "菜粕 2701"))
    n = len(panels)

    fig, axes = plt.subplots(1, n, figsize=(6.6 * n - 0.4, 7.0), facecolor="white",
                             gridspec_kw=dict(wspace=0.30 if n == 3 else 0.46,
                                              left=0.045, right=0.955,
                                              top=0.745, bottom=0.115))
    if n == 1:
        axes = [axes]

    labels = ["豆粕01"] + (["豆粕05"] if use_meal05 else []) + ["菜粕01"]
    title = "机构持仓 · " + " / ".join(labels)
    fig.text(0.045, 0.968, title, fontsize=19, fontweight="bold",
             color="#111", va="top")
    fig.text(0.045, 0.908, "{} — {}　累计净持仓（手）　净多为正、净空为负".format(
        dates[0].strftime("%Y/%m/%d"), dates[-1].strftime("%Y/%m/%d")),
        fontsize=10.5, color="#8C8C8C", va="top")

    handles = [Line2D([0], [0], color=PAL[nm], lw=2.4, solid_capstyle="round")
               for nm in MEAL]
    leg = fig.legend(handles, MEAL, loc="center", bbox_to_anchor=(0.5, 0.845),
                     ncol=len(MEAL), frameon=False, fontsize=9.6,
                     handlelength=1.7, columnspacing=1.9, handletextpad=0.7)
    for t in leg.get_texts():
        t.set_color("#4A4A4A")

    for ax, (names, series, ttl) in zip(axes, panels):
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)
        for s in ("left", "bottom"):
            ax.spines[s].set_color("#DADADA")
            ax.spines[s].set_linewidth(0.9)
        ax.tick_params(colors="#555", labelsize=9.5, length=0, pad=7)
        ax.yaxis.set_major_formatter(FuncFormatter(K))
        ax.grid(axis="y", color="#E9E9E9", linewidth=0.8, linestyle=(0, (2, 3)))
        ax.set_axisbelow(True)
        ax.axhline(0, color="#C4C4C4", lw=0.9, ls=(0, (4, 3)), zorder=1)

        for nm in names:
            ax.plot(range(len(dates)), series[nm], color=PAL[nm], lw=2.1,
                    solid_capstyle="round", zorder=4)
            ax.plot(len(dates) - 1, series[nm][-1], "o", ms=5.6, color=PAL[nm],
                    mec="white", mew=1.5, zorder=9)
        ax.set_xlim(-0.22, len(dates) - 1 + 0.50)
        ax.set_xticks(range(len(dates)))
        ax.set_xticklabels([d.strftime("%m/%d") for d in dates], color="#555",
                           fontsize=9.5)
        ax.set_title(ttl, fontsize=12.5, color="#1A1A1A", loc="left", pad=10,
                     fontweight="bold")

        # 上下各留一点空间，给右端标签避让腾位置
        lo, hi = ax.get_ylim()
        pad = (hi - lo) * 0.09
        ax.set_ylim(lo - pad, hi + pad)

        # 右端直标：机构名 + 末值，按像素避让重叠
        # 注意 gap 必须用坐标区像素高度算（不是单个数据单位的像素数）
        items = sorted([[series[nm][-1], nm] for nm in names], key=lambda t: t[0])
        ypix = [ax.transData.transform((len(dates) - 1, v))[1] for v, _ in items]
        gap = max(0.058 * ax.bbox.height, 32.0)
        for i in range(1, len(ypix)):
            if ypix[i] - ypix[i - 1] < gap:
                ypix[i] = ypix[i - 1] + gap
        over = ypix[-1] - (ax.bbox.y1 - 0.015 * ax.bbox.height)
        if over > 0:
            ypix = [y - over for y in ypix]
        under = (ax.bbox.y0 + 0.015 * ax.bbox.height) - ypix[0]
        if under > 0:
            ypix = [y + under for y in ypix]
        inv = ax.transData.inverted()
        for (v, nm), yp in zip(items, ypix):
            ax.annotate("{} {}".format(nm, K(v)), xy=(len(dates) - 1, v),
                        xytext=(len(dates) - 1 + 0.30, inv.transform((0, yp))[1]),
                        textcoords="data", va="center", ha="left", fontsize=9.2,
                        fontweight="bold", color=PAL[nm], annotation_clip=False,
                        arrowprops=dict(arrowstyle="-", color=PAL[nm], lw=0.8,
                                        alpha=0.5, shrinkA=0, shrinkB=2))

    fig.text(0.045, 0.030, "数据来源：东方财富 龙虎榜　·　单位：手", fontsize=8.6,
             color="#A0A0A0", va="bottom")
    return fig


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", default=XLSM, help="xlsm 路径")
    ap.add_argument("--days", type=int, default=5, help="取最近几个交易日，默认 5")
    ap.add_argument("--date", default=None, help="文件名日期 YYYY-MM-DD，默认取表内最新日期")
    ap.add_argument("--out-dir", default=OUT_DIR, help="输出目录")
    ap.add_argument("--no-meal05", action="store_true", help="只画豆粕01 + 菜粕01")
    a = ap.parse_args()

    if a.days < 2:
        die("--days 至少 2")
    dates, meal, meal05, rape = load(a.file, a.days)

    use5 = (not a.no_meal05) and has_data(meal05)
    if not a.no_meal05 and not use5:
        print("⚠️ K–P 列（豆粕05）没有数据，本次只画豆粕01 / 菜粕01", file=sys.stderr)

    if a.date:
        try:
            name_date = datetime.strptime(a.date, "%Y-%m-%d")
        except ValueError:
            die("--date 需要 YYYY-MM-DD")
    else:
        name_date = dates[-1]

    try:
        fig = build(dates, meal, meal05, rape, use5)
    except ImportError:
        die("缺 matplotlib：pip install matplotlib（或改用 export_charts.py 走 Excel 原生导出）", 3)

    os.makedirs(a.out_dir, exist_ok=True)
    out = os.path.join(a.out_dir, "机构持仓_{}.png".format(name_date.strftime("%Y%m%d")))
    fig.savefig(out, dpi=170, facecolor=fig.get_facecolor())
    print("图数据：{} — {}（{} 个交易日）".format(
        dates[0].strftime("%Y/%m/%d"), dates[-1].strftime("%Y/%m/%d"), len(dates)))
    print("✅ 已出图：{}".format(out))
    if a.date and name_date.date() != dates[-1]:
        print("⚠️ 文件名日期 {} 与表内最新日期 {} 不一致".format(
            name_date.date(), dates[-1]), file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
