#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""机构持仓表写入：在 Sheet1 第 2 行插入一行空白，写入当天豆粕01 / 菜粕01 的累计净持仓。

为什么走 Excel COM：机构持仓表.xlsm 里有 6 张图表 + VBA 宏，openpyxl 存盘会把它们丢掉；
行插入也只有在 Excel 里做才不会打乱图表锚点、公式区和条件格式。

列位（Sheet1）：
  A        日期（写真实日期，显示 mm"月"dd"日"）
  B–H      豆粕01：高盛 / 摩根 / 中粮 / 国泰君安 / 东证 / 中信 / 高盛+摩根
  AC–AF    菜粕01：高盛 / 摩根 / 中粮 / 高盛+摩根
  净空一律写负数；高盛+摩根 = 两者带符号相加。

用法：
  python write_positions.py data.json --dry-run
  python write_positions.py data.json
  python write_positions.py data.json --file "D:/WorkBuddyWorkSpace/鲁信工作/机构持仓表.xlsm"
"""

import argparse
import glob
import json
import os
import re
import shutil
import sys
import zipfile
from datetime import date, datetime

DEFAULT_FILE = os.environ.get("INST_HOLDINGS_XLSM") or r"D:/WorkBuddyWorkSpace/鲁信工作/机构持仓表.xlsm"
HEADER_TAG = "豆粕01"

# 图表 1 / 图表 6 固定显示最近 N 天（用户 2026-10-29 要求由 6 改为 5）
CHART_WINDOW = 5

# 列序：B..H
MEAL_COLS = ["高盛", "摩根", "中粮", "国泰君安", "东证", "中信"]
MEAL_C0 = 2          # B
MEAL05_C0 = 11       # K（豆粕05，2026-11-04 起纳入，列序与豆粕01 相同）
RAPE_C0 = 29         # AC
RAPE_COLS = ["高盛", "摩根", "中粮"]

ALIASES = {
    "高盛期货": "高盛", "高盛": "高盛",
    "摩根大通": "摩根", "摩根": "摩根", "摩根大通期货": "摩根",
    "中粮期货": "中粮", "中粮": "中粮",
    "国泰君安期货": "国泰君安", "国泰君安": "国泰君安", "国泰": "国泰君安",
    "东证期货": "东证", "东证": "东证",
    "中信期货": "中信", "中信": "中信",
}
LONG_WORDS = {"long", "多", "净多", "l", "多头"}
SHORT_WORDS = {"short", "空", "净空", "s", "空头"}


def die(msg, code=2):
    print("❌ " + msg, file=sys.stderr)
    sys.exit(code)


def fmt_signed(v):
    if v == 0:
        return "0"
    return ("+" if v > 0 else "-") + str(abs(v))


def to_int(v, what):
    if isinstance(v, (int, float)):
        return int(round(v))
    s = str(v).replace(",", "").replace("，", "").strip()
    try:
        return int(round(float(s)))
    except ValueError:
        die(f"{what} 不是数字：{v!r}")


def signed(rec, label):
    side = str(rec.get("side", "")).strip().lower()
    if side in LONG_WORDS:
        sign = 1
    elif side in SHORT_WORDS:
        sign = -1
    else:
        die(f"{label} 多空方向无法识别：{rec.get('side')!r}")
    cum = rec.get("cum")
    if cum is None:
        die(f"{label} 缺累计持仓")
    cum = abs(to_int(cum, label + "累计"))
    return sign * cum


def pick(store, name, label):
    for k, v in store.items():
        if ALIASES.get(str(k).strip().replace(" ", ""), str(k).strip()) == name:
            return v
    die(f"{label} 数据里找不到「{name}」")


def build_values(data):
    meal_raw = data.get("meal") or data.get("豆粕") or {}
    rape_raw = data.get("rapeseed") or data.get("菜粕") or {}
    if not meal_raw or not rape_raw:
        die("data.json 里 meal / rapeseed 不能为空")

    meal = {}
    for n in MEAL_COLS:
        meal[n] = signed(pick(meal_raw, n, "豆粕"), "豆粕" + n)
    meal_comb = meal["高盛"] + meal["摩根"]

    # 豆粕05（次主力）：可选。缺了就不写 K–Q，绝不拿豆粕01 的数顶上去。
    m5_raw = data.get("meal05") or data.get("豆粕05") or {}
    meal5 = {}
    if m5_raw:
        for n in MEAL_COLS:
            meal5[n] = signed(pick(m5_raw, n, "豆粕05"), "豆粕05" + n)
        meal5_comb = meal5["高盛"] + meal5["摩根"]

    rape = {}
    for n in RAPE_COLS:
        rape[n] = signed(pick(rape_raw, n, "菜粕"), "菜粕" + n)
    rape_comb = rape["高盛"] + rape["摩根"]

    d = data.get("date")
    if d:
        try:
            dd = datetime.strptime(str(d), "%Y-%m-%d").date()
        except ValueError:
            die(f"date 应为 YYYY-MM-DD：{d!r}")
    else:
        dd = date.today()

    row = {}
    for i, n in enumerate(MEAL_COLS):
        row[MEAL_C0 + i] = meal[n]
    row[MEAL_C0 + 6] = meal_comb                      # H = 高盛+摩根
    for i, n in enumerate(MEAL_COLS):
        if meal5:
            row[MEAL05_C0 + i] = meal5[n]
    if meal5:
        row[MEAL05_C0 + 6] = meal5_comb               # Q = 高盛+摩根
    for i, n in enumerate(RAPE_COLS):
        row[RAPE_C0 + i] = rape[n]
    row[RAPE_C0 + 3] = rape_comb                      # AF = 高盛+摩根
    return dd, row


def col_letter(idx):
    s = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def check_lock(path):
    d = os.path.dirname(path)
    base = os.path.basename(path)
    locks = glob.glob(os.path.join(d, "~$" + base))
    locks += glob.glob(os.path.join(d, "~$" + base[2:]))
    # Excel 对超长文件名会截断锁文件名，故按词干兜底匹配；
    # 但同目录其它工作簿的 ~$ 文件（如「头寸明细表」）与本次写入无关，不能算占用。
    if not locks:
        stem = os.path.splitext(base)[0]
        for lk in glob.glob(os.path.join(d, "~$*.xls*")):
            if stem and stem in os.path.basename(lk):
                locks.append(lk)
    # 兜底：真正被 Excel 独占时，连写句柄都拿不到。
    if not locks:
        try:
            with open(path, "r+b"):
                pass
        except OSError:
            locks.append(path)
    if locks:
        die("文件正被 Excel 打开（存在 ~$ 锁文件），请先关闭 Excel 再跑。原文件未改动。", 3)


def backup(path):
    d = os.path.join(os.path.dirname(path), "备份")
    os.makedirs(d, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem, ext = os.path.splitext(os.path.basename(path))
    dst = os.path.join(d, f"{stem}_{stamp}{ext}")
    shutil.copy2(path, dst)
    return dst


def find_sheet(ws_iter):
    for s in ws_iter:
        try:
            if str(s.Range("A1").Value or "").strip() == HEADER_TAG:
                return s
        except Exception:
            continue
    return None


RANGE_RE = re.compile(r"(\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3}\$?\d+)")
SIDE_RE = re.compile(r"^(\$?[A-Z]{1,3})(\$?)(\d+)$")


def _split(ref):
    m = SIDE_RE.match(ref)
    return m.group(1), m.group(2), int(m.group(3))


def shift_formula(f, insert_row=2):
    """按「在第 insert_row 行插入一行」重算图表引用。

    起点在插入行之上的：不动；
    起点正好是 insert_row 的（滚动窗口图）：起点留在 insert_row，
        末行固定为 insert_row + CHART_WINDOW - 1 —— 老的一天挤出、新的一天进来；
    起点在插入行之下：起点、末行都 +1（整块数据跟着下移，窗口大小不变）。
    """
    def rep(m):
        c1, d1, r1 = _split(m.group(1))
        c2, d2, r2 = _split(m.group(2))
        if r1 < insert_row:
            return m.group(0)
        if r1 > insert_row:
            return "{}{}{}:{}{}{}".format(c1, d1, r1 + 1, c2, d2, r2 + 1)
        return "{}{}{}:{}{}{}".format(c1, d1, r1, c2, d2,
                                     r1 + CHART_WINDOW - 1)
    return RANGE_RE.sub(rep, f)


def _strip_sheet(ref):
    return ref.rsplit("!", 1)[-1] if ref else ref


def parse_series(f):
    """拆 =SERIES(name,x,y,order) → (x, y)；name/order 不用。"""
    body = (f or "").strip()
    if not body.startswith("=SERIES(") or not body.endswith(")"):
        return None
    body = body[len("=SERIES("):-1]
    try:
        rest, _order = body.rsplit(",", 1)   # 末尾是绘图顺序
        head, y = rest.rsplit(",", 1)
        _name, x = head.rsplit(",", 1)
    except ValueError:
        return None
    x = _strip_sheet(x.strip())
    y = _strip_sheet(y.strip())
    if not RANGE_RE.fullmatch(x or ""):
        x = ""
    if not RANGE_RE.fullmatch(y or ""):
        return None
    return x, y


def read_series(ws):
    src = {}
    try:
        n = ws.ChartObjects().Count
    except Exception:
        return src
    for i in range(1, n + 1):
        ch = ws.ChartObjects(i).Chart
        try:
            cnt = ch.SeriesCollection().Count
        except Exception:
            continue
        for j in range(1, cnt + 1):
            try:
                p = parse_series(ch.SeriesCollection(j).Formula)
                if p:
                    src[(i, j)] = p
            except Exception:
                pass
    return src


def fix_series(ws, src):
    """把被 Excel 整体下移的数据源改回从新行开始。

    注意：给 Series.Formula 赋值会被 Excel 自动纠正回原范围（尤其次坐标轴上的系列），
    所以这里直接改数据层 XValues / Values，改完再回读校验。
    """
    fixed, bad = [], []
    for (i, j), (x, y) in sorted(src.items()):
        nx, ny = shift_formula(x), shift_formula(y)
        try:
            sc = ws.ChartObjects(i).Chart.SeriesCollection(j)
            # 注意：不能因为 nx==x 就跳过——滚动窗口图的期望范围与插行前完全相同，
            # 但 Excel 已经把它整体下移了一行，必须按「实际值 vs 期望值」判断。
            if parse_series(sc.Formula) == (nx, ny):
                continue
            if nx:
                sc.XValues = ws.Range(nx)
            sc.Values = ws.Range(ny)
            got = parse_series(sc.Formula)
            if got != (nx, ny):
                bad.append((i, j, nx, ny, got))
            else:
                fixed.append((i, j, "{} / {}".format(nx, ny)))
        except Exception as e:
            bad.append((i, j, nx, ny, str(e)))
    return fixed, bad



def snapshot(path):
    """用 openpyxl 只读取证：行数、A/B 列、图表与 VBA 部件。"""
    import warnings
    # 豆粕05 图是复制粘贴生成的，chart XML 里带一个 s:externalData 空引用，
    # openpyxl 解析时会警告；Excel 自己能正常读写，这里只屏蔽这条噪音。
    warnings.filterwarnings("ignore", message=".*Unable to read chart.*")
    import openpyxl
    wb = openpyxl.load_workbook(path, keep_vba=True, data_only=False)
    ws = None
    for s in wb.worksheets:
        if str(s.cell(row=1, column=1).value or "").strip() == HEADER_TAG:
            ws = s
            break
    if ws is None:
        die("工作簿里找不到 A1=豆粕01 的 sheet")
    snap = {
        "max_row": ws.max_row,
        "a": [ws.cell(row=r, column=1).value for r in (2, 3)],
        "b": [ws.cell(row=r, column=2).value for r in (2, 3)],
        "cells": {c: ws[c].value for r in (2,) for c in
                  ([f"A2"] + [f"{col_letter(i)}2" for i in range(2, 9)]
                   + [f"{col_letter(i)}2" for i in range(11, 18)]
                   + [f"{col_letter(i)}2" for i in range(29, 33)])},
    }
    wb.close()
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
    snap["charts"] = sorted(n for n in names if n.startswith("xl/charts/chart"))
    snap["vba"] = [n for n in names if n.endswith(".bin") and "vba" in n]
    return snap


def main():
    ap = argparse.ArgumentParser(description="机构持仓表：插入今日一行")
    ap.add_argument("json", help="数据 JSON 路径，- 从标准输入读")
    ap.add_argument("--file", default=DEFAULT_FILE, help="xlsm 路径")
    ap.add_argument("--dry-run", action="store_true", help="只打印将要写入的值")
    ap.add_argument("--force", action="store_true",
                    help="第 2 行已是同一天时仍然插入")
    ap.add_argument("--no-backup", action="store_true")
    args = ap.parse_args()

    path = os.path.abspath(args.file)
    if not os.path.exists(path):
        die(f"文件不存在：{path}", 2)
    if not path.lower().endswith((".xlsm", ".xlsx")):
        die("只支持 .xlsm / .xlsx", 2)

    if args.json == "-":
        data = json.load(sys.stdin)
    else:
        with open(args.json, "r", encoding="utf-8") as f:
            data = json.load(f)

    dd, row = build_values(data)

    # 打印将要写入的内容
    print("📅 日期：{}（{}月{}日）".format(dd.isoformat(), dd.month, dd.day))
    print("   豆粕01  B–H ：" + "  ".join(
        "{}={}".format(n, fmt_signed(row[MEAL_C0 + i]))
        for i, n in enumerate(MEAL_COLS + ["高盛+摩根"])))
    if any(c in row for c in range(MEAL05_C0, MEAL05_C0 + 7)):
        print("   豆粕05  K–Q ：" + "  ".join(
            "{}={}".format(n, fmt_signed(row[MEAL05_C0 + i]))
            for i, n in enumerate(MEAL_COLS + ["高盛+摩根"])))
    else:
        print("   豆粕05  K–Q ：（本次数据里没有 meal05，K–Q 保持空白）")
    print("   菜粕01  AC–AF：" + "  ".join(
        "{}={}".format(n, fmt_signed(row[RAPE_C0 + i]))
        for i, n in enumerate(RAPE_COLS + ["高盛+摩根"])))

    if args.dry_run:
        print("\n（--dry-run，未写文件）")
        return

    check_lock(path)

    before = snapshot(path)
    print("\n📄 写前：{} 行，第 2 行 A={} B={}，图表 {} 个，VBA {} 个".format(
        before["max_row"], before["a"][0], before["b"][0],
        len(before["charts"]), len(before["vba"])))

    if before["max_row"] > 1:
        old = before["a"][0]
        old_d = None
        if isinstance(old, datetime):
            old_d = old.date()
        elif isinstance(old, (int, float)):
            old_d = datetime(1899, 12, 30) + __import__("datetime").timedelta(days=int(old))
        if old_d == dd and not args.force:
            die("第 2 行已经是 {}，今天这行可能已经写过；确认要再插一行请加 --force".format(
                dd.isoformat()), 5)

    bak = None if args.no_backup else backup(path)
    if bak:
        print("💾 备份：{}".format(bak))

    try:
        import win32com.client as win32
    except ImportError:
        die("缺少 pywin32，无法安全插行（openpyxl 会丢图表/VBA）。先 pip install pywin32", 1)

    xl = None
    wb = None
    fixed = []
    chart_bad = []
    try:
        xl = win32.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False
        try:
            xl.AutomationSecurity = 3          # 禁用宏，防止 Workbook_Open 干扰
        except Exception:
            pass
        wb = xl.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False)
        ws = find_sheet(wb.Worksheets)
        if ws is None:
            die("工作簿里找不到 A1=豆粕01 的 sheet", 2)

        # 先记下图表数据源。Excel 插行时，数据范围「从第 2 行开始」的图表会整体
        # 下移一行（$A$2:$A$6 → $A$3:$A$7），新插入的这行就进不了图；插完要手动补回来。
        chart_src = read_series(ws)

        # 在第 2 行插入一行，格式沿用原第 2 行
        ws.Rows(2).Insert(-4121, 0)            # xlDown, xlFormatFromLeftOrAbove
        try:
            ws.Rows(3).Copy()
            ws.Rows(2).PasteSpecial(-4122)      # xlPasteFormats
            xl.CutCopyMode = False
        except Exception:
            pass
        ws.Rows(2).RowHeight = ws.Rows(3).RowHeight

        fixed, chart_bad = fix_series(ws, chart_src)

        # 直接写 Excel 序列号：pywin32 会把 naive datetime 当本地时间转 UTC，
        # 写 datetime 会存成「前一天的 16:00」，所以自己算序列号。
        ws.Range("A2").Value = (dd - date(1899, 12, 30)).days
        ws.Range("A2").NumberFormat = 'mm"月"dd"日"'
        for ci, val in row.items():
            c = ws.Range("{}{}".format(col_letter(ci), 2))
            c.Value = int(val)
            c.NumberFormat = "0"

        wb.Save()
        wb.Close(SaveChanges=False)
        wb = None
    except SystemExit:
        raise
    except Exception as e:
        die("写入失败：{}".format(e), 1)
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass

    # 回读自检
    after = snapshot(path)
    problems = []
    if after["max_row"] != before["max_row"] + 1:
        problems.append("行数 {} → {}，应 +1".format(before["max_row"], after["max_row"]))
    a2 = after["a"][0]
    if not isinstance(a2, datetime) or a2.date() != dd:
        problems.append("A2 不是 {}（读到 {!r}）".format(dd.isoformat(), a2))
    for ci, val in row.items():
        got = after["cells"].get("{}{}".format(col_letter(ci), 2))
        if got != val:
            problems.append("{}{} 应为 {}，实际 {!r}".format(col_letter(ci), 2, val, got))
    if after["b"][1] != before["b"][0]:
        problems.append("旧第 2 行没有下移到第 3 行（B3 应为 {!r}，实际 {!r}）".format(
            before["b"][0], after["b"][1]))
    if after["charts"] != before["charts"]:
        problems.append("图表部件变了：{} → {}".format(before["charts"], after["charts"]))
    if after["vba"] != before["vba"]:
        problems.append("VBA 部件变了：{} → {}".format(before["vba"], after["vba"]))

    if problems:
        print("\n⚠️ 自检未通过：", file=sys.stderr)
        for p in problems:
            print("   - " + p, file=sys.stderr)
        if bak:
            print("   备份可用于回滚：{}".format(bak), file=sys.stderr)
        sys.exit(4)

    print("\n✅ 完成")
    print("   在第 2 行插入新行，原第 2 行下移到第 3 行；总行数 {} → {}".format(
        before["max_row"], after["max_row"]))
    print("   A2 = {}（显示 {}月{}日）".format(dd.isoformat(), dd.month, dd.day))
    print("   图表 {} 个、VBA 部件 {} 个 均保留；数据源修正 {} 处".format(
        len(after["charts"]), len(after["vba"]), len(fixed)))
    for i, j, f in fixed:
        print("     图表{} 系列{} → {}".format(i, j, f))
    if chart_bad:
        print("\n⚠️ 有 {} 个图表系列的数据源未能自动修正，请手工核对：".format(len(chart_bad)),
              file=sys.stderr)
        for item in chart_bad:
            print("   图表{} 系列{} 期望 {} / {}，实际 {}".format(*item), file=sys.stderr)
        sys.exit(4)


if __name__ == "__main__":
    main()
