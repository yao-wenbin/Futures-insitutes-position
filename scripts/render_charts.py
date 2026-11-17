#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""机构持仓合并图（研报简洁风）：按 config.json 的 objects 逐合约出一张面板。

每个 object 读自己的 CSV（contract → 文件名），取最近 N 个交易日（各合约日期并集），
渲染成一张并排多面板 PNG。只读 CSV，不写数据。
面板与线条完全由 objects 的 chart_orgs 决定。

用法：
  python render_charts.py                       # 最近 5 日，文件名取最新日期
  python render_charts.py --date 2026-11-04     # 指定文件名日期
  python render_charts.py --days 10             # 周报用 10 日窗口
  python render_charts.py --data-dir ./data     # 指定 CSV 目录
  python render_charts.py --group meal_rape     # 只画一个分组

输出：<out-dir>/机构持仓_YYYYMMDD[_分组名].png
"""
import argparse
import os
import sys
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import config_loader  # pyright: ignore[reportImplicitRelativeImport]
import csv_store  # pyright: ignore[reportImplicitRelativeImport]
import window_data  # pyright: ignore[reportImplicitRelativeImport]

# 固定机构配色（按名称稳定着色）
PAL = {"高盛": "#E4572E", "摩根": "#2E7EBF", "中粮": "#333333",
       "国泰君安": "#D99A2B", "东证": "#7C5CFF", "中信": "#0FA37F"}
EXTRA = ["#C0392B", "#2980B9", "#27AE60", "#8E44AD",
         "#D35400", "#16A085", "#2C3E50", "#7F8C8D"]


def color_of(name):
    if name in PAL:
        return PAL[name]
    return EXTRA[sum(ord(c) for c in str(name)) % len(EXTRA)]


def die(msg, code=2):
    print("❌ " + msg, file=sys.stderr)
    sys.exit(code)


def load(data_dir, n_days, objects):
    """从各合约 CSV 读取最近 n_days 个交易日，返回 (dates, panels)。"""
    tables = {}
    all_dates = set()
    for obj in objects:
        fp = os.path.join(data_dir, csv_store.contract_filename(obj["contract"]))
        dates, data = csv_store.read_csv(fp)
        tables[obj["contract"]] = data
        all_dates.update(dates)
    if not all_dates:
        die("CSV 目录里没有数据：{}".format(data_dir))
    dates = sorted(all_dates)[-n_days:]
    date_objs = [datetime.strptime(d, "%Y-%m-%d").date() for d in dates]

    panels = []
    for obj in objects:
        tbl = tables.get(obj["contract"], {})
        orgs = list(obj["chart_orgs"])
        series = {n: [float(tbl.get(n, {}).get(d) or 0) for d in dates] for n in orgs}
        ct = config_loader.contract_digits(obj["contract"])
        variety = config_loader.variety_name(obj)
        title = "{} {}".format(variety, ct) + ("（主力）" if obj["primary"] else "")
        panels.append({"name": obj["name"], "title": title, "orgs": orgs, "series": series})
    return date_objs, panels


def load_dates(data_dir, date_strs, objects):
    """按指定日期列表（升序 YYYY-MM-DD）读取各合约数据，返回 (dates, panels)。"""
    tables = {}
    for obj in objects:
        fp = os.path.join(data_dir, csv_store.contract_filename(obj["contract"]))
        _, data = csv_store.read_csv(fp)
        tables[obj["contract"]] = data
    date_objs = [datetime.strptime(d, "%Y-%m-%d").date() for d in date_strs]
    panels = []
    for obj in objects:
        tbl = tables.get(obj["contract"], {})
        orgs = list(obj["chart_orgs"])
        series = {n: [float(tbl.get(n, {}).get(d) or 0) for d in date_strs] for n in orgs}
        ct = config_loader.contract_digits(obj["contract"])
        variety = config_loader.variety_name(obj)
        title = "{} {}".format(variety, ct) + ("（主力）" if obj["primary"] else "")
        panels.append({"name": obj["name"], "title": title, "orgs": orgs, "series": series})
    return date_objs, panels


def has_data(series):
    return any(abs(v) > 0 for vals in series.values() for v in vals)


def build(dates, panels):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.lines import Line2D
    from matplotlib.ticker import FuncFormatter

    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "Source Han Sans CN", "SimHei"]
    plt.rcParams["axes.unicode_minus"] = False

    def K(v, _=None):
        return "{:,.0f}".format(v)

    n = len(panels)
    fig, axes = plt.subplots(1, n, figsize=(6.6 * n - 0.4, 7.0), facecolor="white",
                             gridspec_kw=dict(wspace=0.30 if n == 3 else 0.46,
                                              left=0.045, right=0.955,
                                              top=0.745, bottom=0.115))
    if n == 1:
        axes = [axes]

    labels = [p["name"] for p in panels]
    title = "机构持仓 · " + " / ".join(labels)
    fig.text(0.045, 0.968, title, fontsize=19, fontweight="bold",
             color="#111", va="top")
    fig.text(0.045, 0.908, "{} — {}　累计净持仓（手）　净多为正、净空为负".format(
        dates[0].strftime("%Y/%m/%d"), dates[-1].strftime("%Y/%m/%d")),
        fontsize=10.5, color="#8C8C8C", va="top")

    # 图例：所有面板 chart_orgs 的并集（按首次出现顺序）
    legend_orgs = []
    for p in panels:
        for o in p["orgs"]:
            if o not in legend_orgs:
                legend_orgs.append(o)
    handles = [Line2D([0], [0], color=color_of(nm), lw=2.4, solid_capstyle="round")
               for nm in legend_orgs]
    leg = fig.legend(handles, legend_orgs, loc="center", bbox_to_anchor=(0.5, 0.845),
                     ncol=len(legend_orgs), frameon=False, fontsize=9.6,
                     handlelength=1.7, columnspacing=1.9, handletextpad=0.7)
    for t in leg.get_texts():
        t.set_color("#4A4A4A")

    for ax, p in zip(axes, panels):
        names = p["orgs"]
        series = p["series"]
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
            ax.plot(range(len(dates)), series[nm], color=color_of(nm), lw=2.1,
                    solid_capstyle="round", zorder=4)
            ax.plot(len(dates) - 1, series[nm][-1], "o", ms=5.6, color=color_of(nm),
                    mec="white", mew=1.5, zorder=9)
        ax.set_xlim(-0.22, len(dates) - 1 + 0.50)
        ax.set_xticks(range(len(dates)))
        ax.set_xticklabels([d.strftime("%m/%d") for d in dates], color="#555",
                           fontsize=9.5)
        ax.set_title(p["title"], fontsize=12.5, color="#1A1A1A", loc="left", pad=10,
                     fontweight="bold")

        # 上下各留一点空间，给右端标签避让腾位置
        lo, hi = ax.get_ylim()
        pad = (hi - lo) * 0.09
        ax.set_ylim(lo - pad, hi + pad)

        # 右端直标：机构名 + 末值，按像素避让重叠
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
                        fontweight="bold", color=color_of(nm), annotation_clip=False,
                        arrowprops=dict(arrowstyle="-", color=color_of(nm), lw=0.8,
                                        alpha=0.5, shrinkA=0, shrinkB=2))

    fig.text(0.045, 0.030, "数据来源：东方财富 龙虎榜　·　单位：手", fontsize=8.6,
             color="#A0A0A0", va="bottom")
    return fig


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-dir", default=None, help="CSV 目录（默认 config.json 的 data_dir）")
    ap.add_argument("--days", type=int, default=5, help="窗口交易日数，默认 5")
    ap.add_argument("--date", default=None, help="文件名日期 YYYY-MM-DD，默认取窗口最新日")
    ap.add_argument("--out-dir", default=None, help="输出目录")
    ap.add_argument("--group", default=None,
                    help="只画指定分组（config.json 的分组 id 或名称）")
    ap.add_argument("--config", default=None, help="config.json 路径（默认项目根）")
    ap.add_argument("--no-auto", action="store_true",
                    help="不按 16:00 规则自动算窗口/补数，改用 CSV 里最近 N 个日期")
    a = ap.parse_args()

    if a.days < 2:
        die("--days 至少 2")

    cfg = config_loader.load_config(a.config)
    objects = config_loader.get_objects(cfg)
    label = ""
    if a.group:
        grp = config_loader.find_group(cfg, a.group)
        if not grp:
            die("config.json 里没有分组：{}".format(a.group))
        objects = grp["objects"]
        label = grp["name"]
    if not objects:
        die("config.json 没有 objects 配置")
    data_dir = config_loader.data_dir(a.data_dir, cfg)
    out_dir = (a.out_dir or os.environ.get("INST_HOLDINGS_OUT")
               or config_loader.resolve_path(None, cfg.get("out_dir"), os.path.join("..", "图表数据")))

    if a.no_auto:
        dates, panels = load(data_dir, a.days, objects)
    else:
        win = window_data.window_dates(a.days)
        date_strs = [d.isoformat() for d in win]
        print("窗口（按 16:00 规则）：{} — {}（{} 个交易日）".format(
            date_strs[0], date_strs[-1], len(date_strs)))
        window_data.ensure(data_dir, objects, win)
        dates, panels = load_dates(data_dir, date_strs, objects)
    # 跳过全空面板（如某合约刚配置、窗口内还没数据）
    visible = [p for p in panels if has_data(p["series"])] or panels[:1]
    if len(visible) < len(panels):
        skipped = [p["title"] for p in panels if p not in visible]
        print("⚠️ 跳过无数据面板：{}".format("、".join(skipped)), file=sys.stderr)

    if a.date:
        try:
            name_date = datetime.strptime(a.date, "%Y-%m-%d")
        except ValueError:
            die("--date 需要 YYYY-MM-DD")
    else:
        name_date = dates[-1]

    try:
        fig = build(dates, visible)
    except ImportError:
        die("缺 matplotlib：pip install matplotlib", 3)

    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "机构持仓_{}{}.png".format(
        name_date.strftime("%Y%m%d"), "_{}".format(label) if label else ""))
    fig.savefig(out, dpi=170, facecolor=fig.get_facecolor())
    print("图数据：{} — {}（{} 个交易日）".format(
        dates[0].strftime("%Y/%m/%d"), dates[-1].strftime("%Y/%m/%d"), len(dates)))
    print("✅ 已出图：{}".format(out))
    if a.date and name_date.date() != dates[-1]:
        print("⚠️ 文件名日期 {} 与 CSV 最新日期 {} 不一致".format(
            name_date.date(), dates[-1]), file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
