# -*- coding: utf-8 -*-
"""从东方财富「期货持仓龙虎榜」接口直接取数，生成与截图版完全一致的 JSON。

为什么要有这个脚本：截图 → 手工抄数容易看错行，而且每天要截 4 张图。
接口是东方财富数据中心同一个数据源（RPT_FUTU_DAILYPOSITION），数值与 App
「净多单 / 净空单」页面逐项一致。

口径（照抄 App，勿改）：
  净持仓 = 多单量 − 空单量
  但只统计进得了**前 20 名**的那一边：某家若没进前 20，接口该字段为 null，
  这一边按 0 计。所以会出现「高盛 净多 109573（= 多单量，因为空单没进前 20）」
  这种看着像多单量的数——这是 App 的口径，本表历史数据一直如此。
  增减同理：净增减 = 多单增减 − 空单增减（未进前 20 的一边按 0）。

用法：
    python fetch_positions.py                       # 取最新交易日，输出到 stdout
    python fetch_positions.py --date 2026-10-29
    python fetch_positions.py --out data.json       # 顺便落盘
    python fetch_positions.py --check data.json     # 与已有人工取数比对
返回码：0 成功；1 网络/接口异常；2 参数错误；4 与 --check 目标不一致。
"""
import argparse
import datetime
import json
import sys
import urllib.parse
import urllib.request

API = "https://datacenter-web.eastmoney.com/api/data/v1/get"
REPORT = "RPT_FUTU_DAILYPOSITION"
TOPN = 20

# 接口增减字段与「累计差」不一致的记录（某边上榜/掉榜导致），取数后打印给用户看
AUDIT = []

# 品种 → 东财 SECURITY_CODE
# meal05 = 豆粕 05（次主力），2026-11-04 起纳入统计，写 K–Q 列
DEFAULT_CODES = {"meal": "M2701", "meal05": "M2705", "rapeseed": "RM2701"}
DEFAULT_NAMES = {"meal": "豆粕01", "meal05": "豆粕05", "rapeseed": "菜粕01"}

# 需要的机构（东财简称，去掉「（代客）」后匹配）
WANTED = {
    "meal": ["高盛期货", "摩根大通", "中粮期货", "国泰君安", "东证期货", "中信期货"],
    "meal05": ["高盛期货", "摩根大通", "中粮期货", "国泰君安", "东证期货", "中信期货"],
    "rapeseed": ["高盛期货", "摩根大通", "中粮期货"],
}
# 输出 JSON 里的键名
OUT_KEY = {
    "高盛期货": "高盛",
    "摩根大通": "摩根",
    "中粮期货": "中粮",
    "国泰君安": "国泰君安",
    "东证期货": "东证",
    "中信期货": "中信",
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


def net_map(rows, wanted):
    """把某一天的接口行压成 {机构: 带符号净持仓}。"""
    agg = {}
    for r in rows:
        name = clean_name(r.get("MEMBER_NAME_ABBR") or r.get("ORG_NAME_ABBR_NEW"))
        if name in ("本日合计", "上日合计", "总量增减", ""):
            continue
        if name not in wanted:
            continue
        net, _ = net_of(r)
        agg[name] = agg.get(name, 0) + net
    return agg


def build(rows, wanted, prev_net=None):
    """把接口行整理成 机构 → {side, cum, chg}。

    prev_net 给出时，chg 用「今日累计 − 昨日累计」算，而不是接口的增减字段。
    原因：某机构的多单（或空单）昨天刚够进前 20、今天掉出榜，接口的
    LP_CHANGE / SP_CHANGE 会直接变 null、按 0 计，于是「空头明明增加 19810」
    会显示成「减少 3914」。累计差没有这个毛病——两边都在榜时，它与增减列
    数学上完全等价（累计差 = 多增 − 空增 = 增减列），只在进榜状态翻转时才不同，
    而那正是增减列失真的情形。
    """
    agg = {}
    for r in rows:
        name = clean_name(r.get("MEMBER_NAME_ABBR") or r.get("ORG_NAME_ABBR_NEW"))
        if name in ("本日合计", "上日合计", "总量增减", ""):
            continue
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
                AUDIT.append((OUT_KEY.get(name, name), chg_api, chg))
        else:
            chg = chg_api
        out[OUT_KEY[name]] = {
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



def collect(day=None, codes=None):
    codes = codes or DEFAULT_CODES
    payload = {"contracts": {}}
    used_date = None
    missing = []
    for key, code in codes.items():
        if day:
            rows, real = fetch_day(code, day)
            if not rows:
                missing.append("{}（{} {} 无数据）".format(DEFAULT_NAMES[key], code, day))
                continue
            dstr = real or day
        else:
            d, rows = latest_available(code)
            if not rows:
                missing.append("{}（{} 近 10 日无数据）".format(DEFAULT_NAMES[key], code))
                continue
            dstr = d.isoformat()
        used_date = used_date or dstr
        # 上一交易日：增减用「累计差」算，避免上榜/掉榜时增减字段失真
        prev_rows, prev_date = prev_trading_rows(code, dstr)
        if prev_rows:
            prev_net = net_map(prev_rows, WANTED[key])
        else:
            prev_net = None
            print("⚠️ {} 找不到 {} 之前的交易日，增减退回接口字段".format(
                DEFAULT_NAMES[key], dstr), file=sys.stderr)
        payload[key] = build(rows, WANTED[key], prev_net)
        payload["contracts"][key] = code
        payload.setdefault("prev_date", {})[key] = prev_date

    if used_date:
        payload["date"] = used_date
    return payload, missing


def _flatten(obj):
    """摊平成 [(品种, 机构名, {side,cum,chg})]。"""
    out = []
    for grp in ("meal", "meal05", "rapeseed"):
        for k, v in (obj.get(grp) or {}).items():
            out.append((grp, k, v))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default=None, help="YYYY-MM-DD，默认取最新交易日")
    ap.add_argument("--out", default=None, help="把 JSON 写到这个路径")
    ap.add_argument("--check", default=None, help="与这份 JSON 比对，不一致返回 4")
    ap.add_argument("--require-today", action="store_true",
                    help="接口返回的日期必须是今天，否则返回 6（当日数据未更新）")
    ap.add_argument("--meal-code", default=None)
    ap.add_argument("--meal05-code", default=None, help="豆粕05 合约代码，默认 M2705")
    ap.add_argument("--rape-code", default=None)
    ap.add_argument("--check-prev", action="store_true",
                    help="严格模式：接口增减与累计差不一致时报错返回 7（有机构进出前20）")
    a = ap.parse_args()

    codes = dict(DEFAULT_CODES)
    if a.meal_code:
        codes["meal"] = a.meal_code
    if a.meal05_code:
        codes["meal05"] = a.meal05_code
    if a.rape_code:
        codes["rapeseed"] = a.rape_code

    day = None
    if a.date:
        try:
            day = datetime.date.fromisoformat(a.date).isoformat()
        except ValueError:
            print("✗ --date 需要 YYYY-MM-DD", file=sys.stderr)
            return 2

    try:
        payload, missing = collect(day, codes)
    except Exception as e:
        print("✗ 取数失败：{}".format(e), file=sys.stderr)
        return 1

    for m in missing:
        print("⚠️ 缺：{}".format(m), file=sys.stderr)
    if "meal" not in payload or "rapeseed" not in payload:
        print("✗ 两个品种没有都取到数，退出", file=sys.stderr)
        return 1
    if "meal05" not in payload:
        print("⚠️ 豆粕05 没取到，本次不写 K–Q 列", file=sys.stderr)

    txt = json.dumps(payload, ensure_ascii=False, indent=2)

    # 接口增减字段与「累计差」不一致 —— 说明某机构的多单或空单今天进/出了前 20 名。
    # 报告里用的是累计差（更符合 App 上看到的净持仓变化），这里把差异列出来备查。
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
        diffs = []
        for grp in ("meal", "meal05", "rapeseed"):
            for k, v in (want.get(grp) or {}).items():
                got = (payload.get(grp) or {}).get(k)
                if got != v:
                    diffs.append("{} {}: 人工 {} ≠ 接口 {}".format(grp, k, v, got))
        if diffs:
            print("\n✗ 与 {} 不一致：".format(a.check), file=sys.stderr)
            for d in diffs:
                print("   " + d, file=sys.stderr)
            return 4
        print("\n✅ 与 {} 逐项一致".format(a.check), file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
