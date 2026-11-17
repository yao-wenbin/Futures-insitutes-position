#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把最近 N 个交易日的豆粕05（M2705）累计净持仓补进 Sheet1 的 K–Q 列。

为什么单独一个脚本：日常流程是「插第 2 行写当天」，而豆粕05 此前长期没维护，
需要一次把已有日期行补齐。补完以后每天由 write_positions.py 顺带写 K2–Q2。

列位（K–Q），与 B–H 完全同构：
  K 高盛 / L 摩根 / M 中粮 / N 国泰君安 / O 东证 / P 中信 / Q 高盛+摩根
净空写负数；Q = K + L（带符号相加）。

安全措施：
  * 逐行校验 A 列日期与接口返回的交易日完全一致，不一致就整批放弃；
  * 只写 K–Q，不碰 A–J、R 以后任何单元格；
  * 写前备份、写后回读校验；
  * 绝不用 openpyxl 存盘（会丢图表和 VBA）。

用法：
  python backfill_meal05.py --rows 5 --dry-run
  python backfill_meal05.py --rows 5
返回码：0 成功；1 写入异常；2 参数/路径问题；3 文件被 Excel 占用；4 自检未通过。
"""

import argparse
import datetime
import importlib.util
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_FILE = os.environ.get("INST_HOLDINGS_XLSM") or r"D:/WorkBuddyWorkSpace/鲁信工作/机构持仓表.xlsm"
CODE = "M2705"
COLS = ["高盛", "摩根", "中粮", "国泰君安", "东证", "中信"]
C0 = 11                      # K

EPOCH = datetime.date(1899, 12, 30)


def load_fetch():
    path = os.path.join(HERE, "fetch_positions.py")
    spec = importlib.util.spec_from_file_location("fetch_positions", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def col_letter(idx):
    s = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def to_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, (int, float)):
        return EPOCH + datetime.timedelta(days=int(v))
    return None


def check_lock(path):
    d = os.path.dirname(path)
    base = os.path.basename(path)
    for lk in os.listdir(d):
        if lk.startswith("~$") and os.path.splitext(base)[0] in lk:
            print("✗ 文件正被 Excel 打开（{}），请先关闭".format(lk), file=sys.stderr)
            return False
    try:
        with open(path, "r+b"):
            pass
    except OSError:
        print("✗ 文件无法写入（可能被占用）", file=sys.stderr)
        return False
    return True


def read_dates(path, rows):
    import warnings
    warnings.filterwarnings("ignore", message=".*Unable to read chart.*")
    import openpyxl
    wb = openpyxl.load_workbook(path, keep_vba=True, data_only=True)
    ws = wb["Sheet1"]
    out = []
    for i in range(rows):
        out.append(to_date(ws.cell(row=2 + i, column=1).value))
    wb.close()
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", default=DEFAULT_FILE)
    ap.add_argument("--rows", type=int, default=5, help="补多少行，默认 5（第 2 行起）")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--no-backup", action="store_true")
    a = ap.parse_args()

    path = os.path.abspath(a.file)
    if not os.path.isfile(path):
        print("✗ 找不到 {}".format(path), file=sys.stderr)
        return 2

    dates = read_dates(path, a.rows)
    print("表中前 {} 行日期：{}".format(
        a.rows, "、".join(d.isoformat() if d else "（空）" for d in dates)))
    if any(d is None for d in dates):
        print("✗ 前 {} 行里有空日期，先补好再跑".format(a.rows), file=sys.stderr)
        return 2

    fp = load_fetch()
    plan = []
    for d in dates:
        rows, real = fp.fetch_day(CODE, d.isoformat())
        if not rows:
            print("✗ {} 没有 {} 的持仓数据".format(d.isoformat(), CODE), file=sys.stderr)
            return 2
        if str(real)[:10] != d.isoformat():
            print("✗ {} 接口返回的是 {}，与表内日期对不上".format(d.isoformat(), real),
                  file=sys.stderr)
            return 2
        prev_rows, _pd = fp.prev_trading_rows(CODE, d.isoformat())
        prev_net = fp.net_map(prev_rows, fp.WANTED["meal05"]) if prev_rows else None
        got = fp.build(rows, fp.WANTED["meal05"], prev_net)
        miss = [n for n in COLS if n not in got]
        if miss:
            print("✗ {} 缺机构：{}".format(d.isoformat(), "、".join(miss)), file=sys.stderr)
            return 2
        vals = {}
        for n in COLS:
            v = got[n]
            vals[n] = v["cum"] if v["side"] == "long" else -v["cum"]
        vals["高盛+摩根"] = vals["高盛"] + vals["摩根"]
        plan.append((d, vals))

    print("\n将要写入 K–Q：")
    for d, vals in plan:
        print("  {}  {}".format(d.isoformat(), "  ".join(
            "{}={}".format(n, ("+" if vals[n] > 0 else "") + str(vals[n]))
            for n in COLS + ["高盛+摩根"])))

    if a.dry_run:
        print("\n（--dry-run，未写文件）")
        return 0

    if not check_lock(path):
        return 3

    if not a.no_backup:
        bak_dir = os.path.join(os.path.dirname(path), "备份")
        os.makedirs(bak_dir, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        bak = os.path.join(bak_dir, "机构持仓表_{}_补豆粕05前.xlsm".format(stamp))
        shutil.copy2(path, bak)
        print("\n💾 备份：{}".format(bak))

    import win32com.client as win32
    xl = None
    wb = None
    try:
        xl = win32.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False
        try:
            xl.AutomationSecurity = 3
        except Exception:
            pass
        wb = xl.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False)
        ws = wb.Worksheets("Sheet1")
        for i, (d, vals) in enumerate(plan):
            r = 2 + i
            for j, n in enumerate(COLS + ["高盛+摩根"]):
                c = ws.Cells(r, C0 + j)
                c.Value = int(vals[n])
                c.NumberFormat = "0"
        wb.Save()
        wb.Close(SaveChanges=False)
        wb = None
    except Exception as e:
        print("✗ 写入失败：{}".format(e), file=sys.stderr)
        return 1
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

    # 回读校验
    import warnings
    warnings.filterwarnings("ignore", message=".*Unable to read chart.*")
    import openpyxl
    wb2 = openpyxl.load_workbook(path, keep_vba=True, data_only=True)
    ws2 = wb2["Sheet1"]
    bad = []
    for i, (d, vals) in enumerate(plan):
        r = 2 + i
        for j, n in enumerate(COLS + ["高盛+摩根"]):
            got = ws2.cell(row=r, column=C0 + j).value
            if got != vals[n]:
                bad.append("{}{} 应为 {} 实际 {!r}".format(col_letter(C0 + j), r, vals[n], got))
    max_row = ws2.max_row
    wb2.close()

    if bad:
        print("\n⚠️ 回读不一致：", file=sys.stderr)
        for b in bad:
            print("   - " + b, file=sys.stderr)
        return 4

    print("\n✅ 完成：第 2–{} 行 K–Q 已写入（{} 行 × 7 列 = {} 个值），总行数 {} 未变".format(
        1 + a.rows, a.rows, a.rows * 7, max_row))
    return 0


if __name__ == "__main__":
    sys.exit(main())
