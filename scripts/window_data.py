# -*- coding: utf-8 -*-
"""窗口数据管理：按「最近 5 个交易日」确保 CSV 完整，缺什么补什么。

规则（用户 2026-11-17 指定）：
  - 最新数据日期：执行时间 < 16:00 → 昨天；>= 16:00 → 今天；自动跳过周末；
  - 从最新日期往前取 5 个交易日（跳过周六周日）：
    如 11-17 14:36 → 11-16、11-13、11-12、11-11、11-10；
  - 先检查 data 中各合约 CSV 是否已有这些日期 × 配置机构的数据；
  - 完整则直接用；缺失则记录（合约/机构/日期）并调东财接口补齐，写入 CSV；
  - 最后供出图使用。
"""
import datetime
import os
import sys
from datetime import date, time, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import csv_store  # pyright: ignore[reportImplicitRelativeImport]
import fetch_positions as fp  # pyright: ignore[reportImplicitRelativeImport]

CUTOFF = time(16, 0)
WINDOW = 5


def latest_trade_date(now=None):
    """按 16:00 规则 + 跳过周末，返回「最新数据日期」。

    < 16:00 认为是昨天（收盘数据未出），>= 16:00 认为是今天；
    若落在周末则继续往前推到最近的工作日。
    """
    now = now or datetime.datetime.now()
    d = now.date()
    if now.time() < CUTOFF:
        d -= timedelta(days=1)
    while d.weekday() >= 5:      # 5=周六、6=周日
        d -= timedelta(days=1)
    return d


def trade_days_back(end, n=WINDOW):
    """从 end 往前取 n 个交易日（跳过周末），升序返回。"""
    days = []
    d = end
    while len(days) < n:
        if d.weekday() < 5:
            days.append(d)
        d -= timedelta(days=1)
    return list(reversed(days))


def window_dates(n=WINDOW, now=None):
    """返回最近 n 个交易日（date 列表，升序）。"""
    return trade_days_back(latest_trade_date(now), n)


def _as_str(d):
    return d.isoformat() if isinstance(d, date) else str(d)


def check_missing(data_dir, objects, dates):
    """检查窗口内数据完整性，返回 [(obj, 日期字符串, 缺失机构列表, 说明)]。"""
    ds_list = [_as_str(d) for d in dates]
    missing = []
    for obj in objects:
        path = os.path.join(data_dir, csv_store.contract_filename(obj["contract"]))
        have_dates, data = csv_store.read_csv(path)
        have = set(have_dates)
        for ds in ds_list:
            if ds not in have:
                missing.append((obj, ds, list(obj["orgs"]), "整行缺失"))
                continue
            lack = [org for org in obj["orgs"] if data.get(org, {}).get(ds) is None]
            if lack:
                missing.append((obj, ds, lack, "部分机构缺失"))
    return missing


def ensure(data_dir, objects, dates, verbose=True):
    """检查并补齐窗口数据；返回 (缺失记录, 补成功条数, 失败记录)。"""
    missing = check_missing(data_dir, objects, dates)
    if not missing:
        if verbose:
            print("✅ 窗口内 {} 个交易日数据完整，直接出图".format(len(dates)))
        return [], 0, []

    if verbose:
        print("⚠️ 窗口数据不完整，开始补取（{} 条）".format(len(missing)))
        for obj, ds, orgs, reason in missing:
            print("   - {} {} {}：{}".format(
                obj["name"], ds, reason, "、".join(orgs) if orgs else ""))

    filled = 0
    failed = []
    for obj, ds, _orgs, _reason in missing:
        path = os.path.join(data_dir, csv_store.contract_filename(obj["contract"]))
        try:
            rows, _real = fp.fetch_day(obj["contract"], ds)
        except Exception as e:
            failed.append((obj, ds, "接口异常：{}".format(e)))
            continue
        if not rows:
            failed.append((obj, ds, "接口无该日数据"))
            continue
        values = fp.net_map(rows, set(obj["orgs"]))
        if not values:
            failed.append((obj, ds, "接口无匹配机构"))
            continue
        columns = list(obj["orgs"]) + [c["column"] for c in obj["combos"]]
        csv_store.upsert_day(path, ds, values, columns, obj["combos"], force=True)
        filled += 1

    if verbose:
        print("补取完成：成功 {} 条，失败 {} 条".format(filled, len(failed)))
        for obj, ds, why in failed:
            print("   ✗ {} {}：{}".format(obj["name"], ds, why))
    return missing, filled, failed
