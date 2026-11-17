#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""机构持仓写入 CSV：把当日累计净持仓按 config.json 的 objects 追加到各合约 CSV。

合约与机构完全由 config.json 决定；存储布局见 csv_store.py：
  - 每合约一个 CSV，文件名 = 合约小写，如 M2701 → m2701.csv；
  - 行=日期（升序），列=机构（含各 combos 的合计列）；
  - 净空写负数；合计列由 csv_store 按 combos 自动维护。

用法：
  python write_positions.py data.json --dry-run
  python write_positions.py data.json
  python write_positions.py data.json --data-dir ./data
返回码：0 成功；2 参数/数据问题；5 该日已存在（未改动，除非 --force）。
"""
import argparse
import json
import os
import shutil
import sys
from datetime import date, datetime
from typing import NoReturn

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import config_loader  # pyright: ignore[reportImplicitRelativeImport]
import csv_store  # pyright: ignore[reportImplicitRelativeImport]

ALIASES = {
    "高盛期货": "高盛", "高盛": "高盛",
    "摩根大通": "摩根", "摩根": "摩根", "摩根大通期货": "摩根", "JP摩根": "摩根",
    "中粮期货": "中粮", "中粮": "中粮",
    "国泰君安期货": "国泰君安", "国泰君安": "国泰君安", "国泰": "国泰君安",
    "东证期货": "东证", "东证": "东证", "东方证券": "东证",
    "中信期货": "中信", "中信": "中信",
}
LONG_WORDS = {"long", "多", "净多", "l", "多头"}
SHORT_WORDS = {"short", "空", "净空", "s", "空头"}


def die(msg, code=2) -> NoReturn:
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
        die("{} 不是数字：{!r}".format(what, v))


def signed(rec, label):
    side = str(rec.get("side", "")).strip().lower()
    if side in LONG_WORDS:
        sign = 1
    elif side in SHORT_WORDS:
        sign = -1
    else:
        die("{} 多空方向无法识别：{!r}".format(label, rec.get("side")))
    cum = rec.get("cum")
    if cum is None:
        die("{} 缺累计持仓".format(label))
    return sign * abs(to_int(cum, label + "累计"))


def pick(store, name, label):
    for k, v in store.items():
        if ALIASES.get(str(k).strip().replace(" ", ""), str(k).strip()) == name:
            return v
    die("{} 数据里找不到「{}」".format(label, name))


def columns_of(obj):
    """CSV 完整列顺序 = 明细机构 + 各 combo 的合计列。"""
    return list(obj["orgs_order"]) + [c["column"] for c in obj["combos"]]


def build_objects(data):
    """按 JSON 的 objects 组织待写入数据，返回 (日期, [object...])。"""
    objs = data.get("objects")
    if not objs:
        die("JSON 缺少 objects 数组：请用新版 fetch_positions.py 取数"
            "（合约/机构由 config.json 的 objects 配置）")

    d = data.get("date")
    if d:
        ds = csv_store.date_str(d)
        if not ds:
            die("date 应为 YYYY-MM-DD：{!r}".format(d))
        dd = datetime.strptime(ds, "%Y-%m-%d").date()
    else:
        dd = date.today()

    out = []
    for obj in objs:
        contract = str(obj.get("contract") or "").strip()
        if not contract:
            continue
        name = obj.get("name") or contract
        raw = obj.get("orgs") or {}
        orgs_order = [str(x).strip() for x in (obj.get("orgs_order") or []) if str(x).strip()]
        if not orgs_order:
            orgs_order = list(raw.keys())
        values = {}
        for org in orgs_order:
            values[org] = signed(pick(raw, org, name), "{}{}".format(name, org))
        combos = []
        for c in (obj.get("combos") or []):
            members = [str(x).strip() for x in (c.get("members") or []) if str(x).strip()]
            if len(members) < 2:
                continue
            combos.append({
                "members": members,
                "after": c.get("after"),
                "label": c.get("label") or " + ".join(members),
                "column": c.get("column") or "+".join(members),
            })
        out.append({
            "contract": contract,
            "name": name,
            "orgs_order": orgs_order,
            "combos": combos,
            "values": values,
        })
    if not out:
        die("objects 里没有可写入的合约")
    return dd, out


def backup_csvs(paths):
    paths = [p for p in paths if os.path.exists(p)]
    if not paths:
        return None
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    root = os.path.dirname(os.path.dirname(paths[0]))
    dst_dir = os.path.join(root, "备份", "csv_" + stamp)
    os.makedirs(dst_dir, exist_ok=True)
    for p in paths:
        shutil.copy2(p, os.path.join(dst_dir, os.path.basename(p)))
    return dst_dir


def main():
    ap = argparse.ArgumentParser(description="机构持仓 CSV：按 objects 追加当日一行")
    ap.add_argument("json", help="数据 JSON 路径，- 从标准输入读")
    ap.add_argument("--data-dir", default=None, help="CSV 目录，默认读 config.json 的 data_dir")
    ap.add_argument("--config", default=None, help="config.json 路径（默认项目根）")
    ap.add_argument("--dry-run", action="store_true", help="只打印将要写入的值，不落盘")
    ap.add_argument("--force", action="store_true", help="该日期已存在时仍然覆盖")
    ap.add_argument("--no-backup", action="store_true", help="跳过 CSV 备份")
    args = ap.parse_args()

    if args.json == "-":
        data = json.load(sys.stdin)
    else:
        with open(args.json, "r", encoding="utf-8") as f:
            data = json.load(f)

    dd, objects = build_objects(data)
    cfg = config_loader.load_config(args.config)
    data_dir = config_loader.data_dir(args.data_dir, cfg)
    os.makedirs(data_dir, exist_ok=True)

    print("📅 日期：{}（{}月{}日）".format(dd.isoformat(), dd.month, dd.day))
    print("📁 CSV 目录：{}".format(data_dir))

    targets = []
    for obj in objects:
        cols = columns_of(obj)
        path = os.path.join(data_dir, csv_store.contract_filename(obj["contract"]))
        print("   {}（{}）：".format(obj["name"], obj["contract"]))
        for org in obj["orgs_order"]:
            print("      {}={}".format(org, fmt_signed(obj["values"][org])))
        for c in obj["combos"]:
            total = sum(obj["values"].get(m, 0) for m in c["members"] if m in obj["values"])
            print("      {}={}".format(c["column"], fmt_signed(total)))
        targets.append((obj, path, cols))

    if args.dry_run:
        print("\n（--dry-run，未写文件）")
        return 0

    ds = dd.isoformat()
    if not args.force:
        all_exist = True
        for obj, path, cols in targets:
            dates, _ = csv_store.read_csv(path)
            if ds not in dates:
                all_exist = False
                break
        if all_exist:
            print("\n⚠️ 所有 CSV 里 {} 已存在，本次未改动；确需覆盖请加 --force".format(ds))
            return 5

    if not args.no_backup:
        bak = backup_csvs([p for _, p, _ in targets])
        if bak:
            print("💾 备份：{}".format(bak))

    results = []
    for obj, path, cols in targets:
        status, _ = csv_store.upsert_day(
            path, ds, obj["values"], cols, obj["combos"], force=args.force)
        results.append((obj, path, status))

    if all(r[2] == "skip" for r in results):
        print("\n⚠️ 所有 CSV 里 {} 已存在，本次未改动；确需覆盖请加 --force".format(ds))
        return 5

    print("\n✅ 完成")
    for obj, path, status in results:
        tag = {"added": "新增行", "updated": "覆盖行", "skip": "已存在跳过"}[status]
        print("   {}：{}（{}）".format(
            csv_store.contract_filename(obj["contract"]), tag, path))
    return 0


if __name__ == "__main__":
    sys.exit(main())
