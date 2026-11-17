# -*- coding: utf-8 -*-
"""从东方财富「期货持仓龙虎榜」接口直接取数，生成与截图版完全一致的 JSON。

合约与机构完全由 config.json 的 objects 决定（以合约为对象、机构为其子项），
不再在代码里硬编码。

口径（照抄 App，勿改）：
  净持仓 = 多单量 − 空单量
  但只统计进得了**前 20 名**的那一边：某家若没进前 20，接口该字段为 null，
  这一边按 0 计。增减同理。

用法：
    python fetch_positions.py                       # 每个合约取最新交易日
    python fetch_positions.py --date 2026-10-29
    python fetch_positions.py --out data.json       # 顺便落盘
返回码：0 成功；1 网络/接口异常；2 参数/配置错误；4 与 --check 目标不一致；
        6 --require-today 时当日数据未更新；7 --check-prev 严格模式发现增减修正。
"""
import argparse
import datetime
import json
import os
import sys
import urllib.parse
import urllib.request
from typing import Any

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import config_loader  # pyright: ignore[reportImplicitRelativeImport]

API = "https://datacenter-web.eastmoney.com/api/data/v1/get"
REPORT = "RPT_FUTU_DAILYPOSITION"
TOPN = 20

# 接口增减字段与「累计差」不一致的记录（某边上榜/掉榜导致），取数后打印给用户看
AUDIT = []

# 接口会员名 → 配置里的机构简称
ORG_ALIASES = {
    "高盛期货": "高盛", "高盛": "高盛",
    "摩根大通": "摩根", "摩根大通期货": "摩根", "JP摩根": "摩根", "摩根": "摩根",
    "中粮期货": "中粮", "中粮": "中粮",
    "国泰君安期货": "国泰君安", "国泰君安": "国泰君安", "国泰": "国泰君安",
    "东证期货": "东证", "东方证券": "东证", "东证": "东证",
    "中信期货": "中信", "中信": "中信",
    "永安期货": "永安", "永安": "永安",
    "华泰期货": "华泰", "华泰": "华泰",
    "银河期货": "银河", "银河": "银河",
}


def http_json(params):
    url = API + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0",
        "Referer": "https://data.eastmoney.com/",
    })
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read().decode("utf-8"))


def fetch_day(code, day):
    """取某合约某交易日的全部会员持仓；返回 (rows, actual_date)。"""
    flt = '(SECURITY_CODE="{}")(TRADE_DATE=\'{}\')'.format(code, day)
    data = http_json({
        "reportName": REPORT, "columns": "ALL", "filter": flt,
        "pageSize": 200, "pageNumber": 1, "source": "WEB", "client": "WEB",
    })
    if not data.get("success") or not data.get("result"):
        return [], None
    rows = data["result"]["data"] or []
    real = None
    for r in rows:
        d = str(r.get("TRADE_DATE") or "")[:10]
        if d:
            real = d
            break
    return rows, real


def latest_available(code, start=None, back=10):
    """从 start（默认今天）往前找最近一个有数据的交易日。"""
    d = start or datetime.date.today()
    for _ in range(back):
        rows, real = fetch_day(code, d.isoformat())
        if rows:
            return d, rows
        d -= datetime.timedelta(days=1)
    return None, []


def clean_name(s):
    return str(s or "").replace("（代客）", "").replace("(代客)", "").strip()


def norm_org(raw):
    """接口会员名归一化成配置里的简称。"""
    name = clean_name(raw)
    return ORG_ALIASES.get(name, name)


def ranked(row, side):
    """该边是否进了前 20（东财用 null / 9999 表示未进榜）。"""
    rank = row.get("LP_RANK" if side == "long" else "SP_RANK")
    return rank is not None and rank != 9999 and rank <= TOPN


def net_of(row):
    """返回 (net, chg)：净持仓、净增减，口径见文件头。"""
    lg = row.get("LONG_POSITION") if ranked(row, "long") else 0
    sh = row.get("SHORT_POSITION") if ranked(row, "short") else 0
    lc = row.get("LP_CHANGE") if ranked(row, "long") else 0
    sc = row.get("SP_CHANGE") if ranked(row, "short") else 0
    net = (lg or 0) - (sh or 0)
    chg = (lc or 0) - (sc or 0)
    return net, chg


def _skip(raw):
    return raw in ("本日合计", "上日合计", "总量增减", "")


def net_map(rows, wanted):
    """把某一天的接口行压成 {机构: 带符号净持仓}。"""
    agg = {}
    for r in rows:
        raw = clean_name(r.get("MEMBER_NAME_ABBR") or r.get("ORG_NAME_ABBR_NEW"))
        if _skip(raw):
            continue
        name = norm_org(raw)
        if name not in wanted:
            continue
        net, _ = net_of(r)
        agg[name] = agg.get(name, 0) + net
    return agg


def build(rows, wanted, prev_net=None):
    """把接口行整理成 机构 → {side, cum, chg}。

    prev_net 给出时，chg 用「今日累计 − 昨日累计」算，而不是接口的增减字段，
    避免某边刚掉出前 20 名时增减字段失真（详见 SKILL.md）。
    """
    agg = {}
    for r in rows:
        raw = clean_name(r.get("MEMBER_NAME_ABBR") or r.get("ORG_NAME_ABBR_NEW"))
        if _skip(raw):
            continue
        name = norm_org(raw)
        if name not in wanted:
            continue
        net, chg_api = net_of(r)
        cur = agg.setdefault(name, [0, 0])
        cur[0] += net
        cur[1] += chg_api
    out = {}
    for name in wanted:
        if name not in agg:
            continue
        net, chg_api = agg[name]
        if prev_net is not None:
            chg = net - prev_net.get(name, 0)
            if chg != chg_api:
                AUDIT.append((name, chg_api, chg))
        else:
            chg = chg_api
        out[name] = {
            "side": "long" if net >= 0 else "short",
            "cum": abs(net),
            "chg": chg if net >= 0 else -chg,
        }
    return out


def prev_trading_rows(code, day, back=10):
    """取 day 之前最近一个有数据的交易日；返回 (rows, 日期字符串)。"""
    d = datetime.date.fromisoformat(day) - datetime.timedelta(days=1)
    for _ in range(back):
        rows, _real = fetch_day(code, d.isoformat())
        if rows:
            return rows, d.isoformat()
        d -= datetime.timedelta(days=1)
    return [], None


def collect_objects(day, objects):
    """按 objects 逐个合约取数，返回 (payload, missing)。"""
    payload: dict[str, Any] = {"objects": []}
    missing = []
    used_date = None
    for obj in objects:
        code = obj["contract"]
        wanted = set(obj["orgs"])
        if day:
            rows, real = fetch_day(code, day)
            if not rows:
                missing.append("{}（{} {} 无数据）".format(obj["name"], code, day))
                continue
            dstr = real or day
        else:
            d, rows = latest_available(code)
            if not rows or d is None:
                missing.append("{}（{} 近 10 日无数据）".format(obj["name"], code))
                continue
            dstr = d.isoformat()
        used_date = used_date or dstr
        prev_rows, prev_date = prev_trading_rows(code, dstr)
        if prev_rows:
            prev_net = net_map(prev_rows, wanted)
        else:
            prev_net = None
            print("⚠️ {} 找不到 {} 之前的交易日，增减退回接口字段".format(
                obj["name"], dstr), file=sys.stderr)
        payload["objects"].append({
            "contract": code,
            "name": obj["name"],
            "primary": obj["primary"],
            "group": obj.get("_group"),
            "group_name": obj.get("_group_name"),
            "orgs_order": list(obj["orgs"]),
            "combos": list(obj["combos"]),
            "orgs": build(rows, obj["orgs"], prev_net),
            "prev_date": prev_date,
        })
    if used_date:
        payload["date"] = used_date
    return payload, missing


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default=None, help="YYYY-MM-DD，默认取每个合约最新交易日")
    ap.add_argument("--out", default=None, help="把 JSON 写到这个路径")
    ap.add_argument("--check", default=None, help="与这份 JSON 比对，不一致返回 4")
    ap.add_argument("--require-today", action="store_true",
                    help="接口返回日期必须是今天，否则返回 6（当日数据未更新）")
    ap.add_argument("--config", default=None, help="config.json 路径（默认项目根）")
    ap.add_argument("--check-prev", action="store_true",
                    help="严格模式：接口增减与累计差不一致时报错返回 7（有机构进出前20）")
    a = ap.parse_args()

    cfg = config_loader.load_config(a.config)
    objects = config_loader.get_objects(cfg)
    if not objects:
        print("✗ config.json 里没有配置 objects（合约+机构），无法取数", file=sys.stderr)
        return 2

    day = None
    if a.date:
        try:
            day = datetime.date.fromisoformat(a.date).isoformat()
        except ValueError:
            print("✗ --date 需要 YYYY-MM-DD", file=sys.stderr)
            return 2

    try:
        payload, missing = collect_objects(day, objects)
    except Exception as e:
        print("✗ 取数失败：{}".format(e), file=sys.stderr)
        return 1

    for m in missing:
        print("⚠️ 缺：{}".format(m), file=sys.stderr)
    if not payload.get("objects"):
        print("✗ 所有合约都没取到数，退出", file=sys.stderr)
        return 1

    txt = json.dumps(payload, ensure_ascii=False, indent=2)

    if AUDIT:
        print("ℹ️ 有 {} 家机构今日进/出前 20 名，增减已按累计差修正：".format(len(AUDIT)),
              file=sys.stderr)
        for name, api_v, real_v in AUDIT:
            print("   - {}：接口增减 {} → 按累计差 {}".format(name, api_v, real_v),
                  file=sys.stderr)
        if a.check_prev:
            print("   严格模式下视为异常，退出", file=sys.stderr)
            return 7

    if a.require_today:
        today = datetime.date.today().isoformat()
        got = str(payload.get("date") or "")
        if got != today:
            print("⚠️ 当日数据未更新：今天是 {}，接口最新只到 {}".format(today, got or "（空）"),
                  file=sys.stderr)
            print("   不落盘、不写表，按约定等下一次再取。", file=sys.stderr)
            return 6

    print(txt)
    if a.out:
        with open(a.out, "w", encoding="utf-8") as f:
            f.write(txt)
        print("\n→ 已写入 {}".format(a.out), file=sys.stderr)

    if a.check:
        with open(a.check, encoding="utf-8") as f:
            want = json.load(f)
        got_by_ct = {o.get("contract"): o for o in payload.get("objects", [])}
        diffs = []
        for wo in (want.get("objects") or []):
            ct = wo.get("contract")
            got = got_by_ct.get(ct)
            if got is None:
                diffs.append("{}：接口没取到".format(ct))
                continue
            for k, v in (wo.get("orgs") or {}).items():
                if got["orgs"].get(k) != v:
                    diffs.append("{} {}: 人工 {} ≠ 接口 {}".format(
                        ct, k, v, got["orgs"].get(k)))
        if diffs:
            print("\n✗ 与 {} 不一致：".format(a.check), file=sys.stderr)
            for d in diffs:
                print("   " + d, file=sys.stderr)
            return 4
        print("\n✅ 与 {} 逐项一致".format(a.check), file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
