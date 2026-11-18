# -*- coding: utf-8 -*-
"""机构持仓 CSV 存储层（替代原 Excel「机构持仓表.xlsm」）。

布局（用户 2026-11-17 指定）：
  - 每个合约一个 CSV，文件名 = 合约代码小写，如 M2701 → m2701.csv；
  - **行 = 日期**（YYYY-MM-DD，升序，上旧下新），**列 = 机构持仓**；
  - 数值为**带符号累计净持仓**：净多为正、净空为负；缺失留空；
  - 合计列由 config.json 的 objects[].combos 决定（如「高盛+摩根」），
    本模块按 members 带符号相加维护。

CSV 示例（m2701.csv）：
    日期,高盛,摩根,中粮,国泰君安,东证,中信,高盛+摩根
    2026-11-12,32807,-120278,-164903,7570,-65024,74994,-87471
    2026-11-16,27592,-124903,-151961,14255,-66674,59477,-97311

编码统一 UTF-8 with BOM（utf-8-sig），Excel 直接双击打开中文不乱码。
"""
import csv
import os
import re
from datetime import date, datetime

DATE_HEADER = "日期"

DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def contract_filename(code):
    """合约代码 → CSV 文件名：M2701 → m2701.csv；RM2705 → rm2705.csv。"""
    return str(code).strip().lower() + ".csv"


def date_str(v):
    """把 date/datetime/字符串统一成 YYYY-MM-DD；无法识别返回 None。"""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = str(v).strip()
    if DATE_RE.match(s):
        return s
    # 兼容 2026/11/16、2026年11月16日
    s2 = s.replace("/", "-").replace("年", "-").replace("月", "-").replace("日", "")
    try:
        return datetime.strptime(s2, "%Y-%m-%d").date().isoformat()
    except ValueError:
        return None


def _to_int(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return int(round(v))
    s = str(v).strip().replace(",", "").replace("，", "")
    if s == "":
        return None
    try:
        return int(round(float(s)))
    except ValueError:
        return None


def read_csv(path):
    """读取一个 CSV（行 = 日期、列 = 机构），返回 (dates, data)。

    dates: 升序的日期字符串列表
    data:  {机构: {日期: 带符号整数}}
    """
    if not path or not os.path.exists(path):
        return [], {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    if not rows:
        return [], {}
    header = [h.strip() for h in rows[0]]
    orgs = [h for h in header[1:] if h]
    data = {org: {} for org in orgs}
    dates = []
    for row in rows[1:]:
        if not row:
            continue
        d = date_str(row[0])
        if not d:
            continue
        dates.append(d)
        for i, org in enumerate(orgs):
            if i + 1 >= len(row):
                continue
            v = _to_int(row[i + 1])
            if v is not None:
                data[org][d] = v
    return sorted(set(dates)), data


def read_header(path):
    """读 CSV 表头（去掉「日期」列）；文件不存在返回 []。"""
    if not path or not os.path.exists(path):
        return []
    with open(path, encoding="utf-8-sig", newline="") as f:
        row = next(csv.reader(f), [])
    return [h.strip() for h in row[1:] if h.strip()]


def merge_columns(existing, wanted, combos=None):
    """保留 CSV 已有列（历史机构），把 wanted 里新增的列插在合计列之前。

    用于「机构换人但历史不丢」：如棕榈油 CSV 老列为
    高盛/摩根/中粮/高盛+摩根，新配置为 高盛/摩根/中信，合并后
    高盛/摩根/中粮/中信/高盛+摩根，老「中粮」历史值原样保留。
    """
    combo_cols = {(c.get("column") or "+".join(c.get("members") or []))
                  for c in (combos or [])}
    cols = [c for c in (existing or []) if c]
    insert_at = len(cols)
    for i, c in enumerate(cols):
        if c in combo_cols:
            insert_at = i
            break
    for c in (wanted or []):
        if c and c not in cols:
            cols.insert(insert_at, c)
            insert_at += 1
    return cols


def write_csv(path, dates, data, orgs):
    """按「行 = 日期、列 = 机构」写出 CSV（dates 应已升序）。"""
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([DATE_HEADER] + list(orgs))
        for d in dates:
            row = [d]
            for org in orgs:
                v = data.get(org, {}).get(d)
                row.append("" if v is None else v)
            w.writerow(row)


def recompute_combos(data, combos, day):
    """按 combos 计算各合计列：members 带符号相加，写入 data[column][day]。"""
    for c in (combos or []):
        members = c.get("members") or []
        column = c.get("column") or "+".join(members)
        total = 0
        have = False
        for m in members:
            v = data.get(m, {}).get(day)
            if v is not None:
                total += v
                have = True
        if have:
            data.setdefault(column, {})[day] = total
        else:
            data.get(column, {}).pop(day, None)


def upsert_day(csv_path, day, values, columns, combos=None, force=False):
    """把某交易日的累计净持仓追加/更新进 CSV。

    values:  {机构: 带符号累计}（只含明细机构）
    columns: CSV 完整列顺序（明细 + 各合计列）
    combos:  合计组合，用于计算合计列
    返回 (status, detail)：status ∈ {"added","updated","skip"}
    """
    dates, data = read_csv(csv_path)
    ds = date_str(day)
    if not ds:
        raise ValueError("无法识别的日期：{!r}".format(day))

    existed = ds in dates
    if existed and not force:
        return "skip", {"need_create": False}

    if not existed:
        dates.append(ds)
        dates.sort()

    combo_cols = {(c.get("column") or "+".join(c.get("members") or []))
                  for c in (combos or [])}
    for org, v in values.items():
        if org in combo_cols:
            continue
        data.setdefault(org, {})[ds] = int(v)

    recompute_combos(data, combos or [], ds)
    # 保留 CSV 里已有但本次 config 未包含的机构列（历史不丢），新列插在合计列前
    columns = merge_columns(read_header(csv_path), list(columns), combos)
    write_csv(csv_path, dates, data, columns)
    return ("updated" if existed else "added"), {"need_create": not existed}


def latest_date(csv_path):
    dates, _ = read_csv(csv_path)
    return dates[-1] if dates else None


def read_window(csv_path, n_days, orgs):
    """取最近 n_days 个日期（升序）。返回 (dates, {org: [value|None]})。"""
    dates, data = read_csv(csv_path)
    if n_days:
        dates = dates[-n_days:]
    series = {org: [data.get(org, {}).get(d) for d in dates] for org in orgs}
    return dates, series
