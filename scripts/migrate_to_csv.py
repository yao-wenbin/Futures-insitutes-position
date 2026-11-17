# -*- coding: utf-8 -*-
"""一次性迁移：把原「机构持仓表.xlsm」最近一轮数据按 config.json 的 objects 导出成 CSV。

读取 Sheet1（A1=豆粕01）从第 2 行起的数据，按合约把对应列映射到 config 的机构顺序：
  M2701 → B–G，M2705 → K–P，RM2701 → AC–AE
（xlsm 的合计列 H / Q / AF 忽略，由 combos 重新计算。）

遇到「所有配置合约当天全空」的行即判定为上一轮合约周期/空行，停止迁移，
因此只导出最近一轮连续数据（本轮为 2026-07-08 ~ 2026-11-16）。

用法：
  python migrate_to_csv.py --dry-run
  python migrate_to_csv.py
  python migrate_to_csv.py --file 机构持仓表.xlsm --data-dir ../data
"""
import argparse
import os
import sys
from datetime import date, datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import config_loader  # pyright: ignore[reportImplicitRelativeImport]
import csv_store  # pyright: ignore[reportImplicitRelativeImport]

HEADER_TAG = "豆粕01"
# 合约 → xlsm 列索引（A=0）：豆粕01 B–G / 豆粕05 K–P / 菜粕01 AC–AE
XLSM_COLS = {
    "M2701": [1, 2, 3, 4, 5, 6],
    "M2705": [10, 11, 12, 13, 14, 15],
    "RM2701": [28, 29, 30],
}


def die(msg, code=2):
    print("❌ " + msg, file=sys.stderr)
    sys.exit(code)


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=int(v))).date()
    return None


def to_int(v):
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return None
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        return None


def main():
    ap = argparse.ArgumentParser(description="xlsm → CSV 迁移（按 objects，最近一轮）")
    ap.add_argument("--file", default=os.path.join(HERE, "..", "机构持仓表.xlsm"), help="源 xlsm")
    ap.add_argument("--data-dir", default=None, help="CSV 输出目录")
    ap.add_argument("--config", default=None, help="config.json 路径（默认项目根）")
    ap.add_argument("--dry-run", action="store_true", help="只统计不写文件")
    a = ap.parse_args()

    try:
        import warnings
        warnings.filterwarnings("ignore")
        from openpyxl import load_workbook
    except ImportError:
        die("缺 openpyxl：pip install openpyxl")

    cfg = config_loader.load_config(a.config)
    objects = config_loader.get_objects(cfg)
    if not objects:
        die("config.json 里没有 objects 配置")
    data_dir = config_loader.data_dir(a.data_dir, cfg)

    path = os.path.abspath(a.file)
    if not os.path.exists(path):
        die("找不到源文件：{}".format(path))

    wb = load_workbook(path, read_only=True, data_only=True)
    ws = None
    for s in wb.worksheets:
        try:
            if str(s.cell(row=1, column=1).value or "").strip() == HEADER_TAG:
                ws = s
                break
        except Exception:
            continue
    if ws is None:
        die("工作簿里找不到 A1=豆粕01 的 sheet")

    per_obj = {}
    used = 0
    stopped_at = None
    for ri, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row:
            continue
        d = to_date(row[0] if len(row) > 0 else None)
        if d is None:
            stopped_at = ri
            break
        ds = d.isoformat()
        row_has = False
        for obj in objects:
            cols = XLSM_COLS.get(str(obj["contract"]).upper())
            if not cols:
                continue
            vals = {}
            for org, c in zip(obj["orgs"], cols):
                if c < len(row):
                    v = to_int(row[c])
                    if v is not None:
                        vals[org] = v
            if vals:
                row_has = True
                bucket = per_obj.setdefault(obj["contract"], {})
                for org, v in vals.items():
                    bucket.setdefault(org, {})[ds] = v
        if not row_has:
            stopped_at = ri
            break
        used += 1
    wb.close()

    print("迁移范围：Sheet1 第 2 行起 {} 行（到第 {} 行停止：所有配置合约全空或日期为空）".format(
        used, stopped_at))
    for obj in objects:
        data = per_obj.get(obj["contract"])
        if data:
            dates = sorted({x for org in data.values() for x in org})
            print("  {}（{}）：{} 个交易日（{} ~ {}）".format(
                csv_store.contract_filename(obj["contract"]), obj["name"],
                len(dates), dates[0], dates[-1]))
        else:
            print("  {}（{}）：无数据（xlsm 无对应列或该轮无记录）".format(
                csv_store.contract_filename(obj["contract"]), obj["name"]))

    if a.dry_run:
        print("\n（--dry-run，未写文件）")
        return 0

    os.makedirs(data_dir, exist_ok=True)
    for obj in objects:
        data = per_obj.get(obj["contract"])
        if not data:
            continue
        dates = sorted({x for org in data.values() for x in org})
        combos = obj["combos"]
        columns = list(obj["orgs"]) + [c["column"] for c in combos]
        for day in dates:
            csv_store.recompute_combos(data, combos, day)
        out = os.path.join(data_dir, csv_store.contract_filename(obj["contract"]))
        csv_store.write_csv(out, dates, data, columns)
        print("  ✅ 写入 {}".format(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
