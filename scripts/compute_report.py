#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""机构持仓报告生成器

输入：从「东方财富 - 机构净多/净空持仓」截图里抄下来的原始数字（JSON）。
输出：固定格式的持仓变化文本（可直接贴到聊天里）。

符号约定（务必保持一致，否则合计会反）：
  * side = "long"（净多图）：cum 填正数，chg 就是「增减」列的原始带符号数。
  * side = "short"（净空图）：cum 填正数（净空量），chg 也填「增减」列原始带符号数。
    净空表的「增减」正数 = 空头增加（看空），所以脚本内部会把 chg 取反得到「看多增量」。
  * 合计口径：累计 = 带符号相加（净多为正、净空为负）；
              当日增减 = 各自「看多增量」相加，方向按合计累计的方向写。
    例：高盛净多 +17124、摩根净空 -13842 → 当日净多 +30966、合计净多 39715。

用法：
  python compute_report.py data.json
  python compute_report.py data.json --date 2026-10-29 --flag-combined
  python compute_report.py --self-test
"""

import argparse
import json
import re
import sys
from datetime import date, datetime

# 机构顺序与合计组合不再硬编码，统一由 config.json 的 objects 决定（见 build）

ALIASES = {
    "高盛期货": "高盛", "高盛": "高盛",
    "摩根大通": "摩根", "摩根": "摩根", "JP摩根": "摩根", "摩根大通期货": "摩根",
    "中粮期货": "中粮", "中粮": "中粮",
    "国泰君安期货": "国泰君安", "国泰君安": "国泰君安", "国泰": "国泰君安",
    "东证期货": "东证", "东证": "东证", "东方证券": "东证",
    "中信期货": "中信", "中信": "中信",
}
LONG_WORDS = {"long", "多", "净多", "l", "buy", "多头"}
SHORT_WORDS = {"short", "空", "净空", "s", "sell", "空头"}
RATIO_LIMIT = 0.10


def die(msg, code=2):
    print("❌ " + msg, file=sys.stderr)
    sys.exit(code)


def norm_key(k):
    k = str(k).strip().replace(" ", "")
    return ALIASES.get(k, k)


def parse_side(v):
    s = str(v).strip().lower()
    if s in LONG_WORDS:
        return "long"
    if s in SHORT_WORDS:
        return "short"
    die(f"无法识别多空方向：{v!r}（应为 long/short 或 多/空）")


def to_int(v, what):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return int(round(v))
    s = str(v).replace(",", "").replace("，", "").replace("+", "").strip()
    s = s.replace("手", "").strip()
    try:
        return int(round(float(s)))
    except ValueError:
        die(f"{what} 不是数字：{v!r}")


class Rec(object):
    def __init__(self, name, side, cum, chg):
        self.name = name
        self.side = side
        self.cum = abs(cum)
        self.chg = chg
        self.cum_s = self.cum if side == "long" else -self.cum
        self.dlong = chg if side == "long" else -chg
        self.prev = self.cum_s - self.dlong
        if self.prev != 0:
            self.ratio = abs(self.dlong) / float(abs(self.prev))
        else:
            self.ratio = None
        # 多空翻向：前一日与当日分列符号两侧。此时比例没有意义（基数太小或跨零），
        # 报告里改为标注「当日净空转净多 / 净多转净空」，不再写占比。
        self.flip = None
        if self.prev < 0 < self.cum_s:
            self.flip = "当日净空转净多"
        elif self.cum_s < 0 < self.prev:
            self.flip = "当日净多转净空"

    @property
    def dirword(self):
        return "多" if self.cum_s > 0 else "空"


def load_section(raw, label, required):
    out = {}
    if not isinstance(raw, dict):
        die(f"{label} 数据缺失或格式不对")
    for k, v in raw.items():
        key = norm_key(k)
        if key not in required:
            continue
        if v is None:
            continue
        if not isinstance(v, dict):
            die(f"{label}「{key}」的值应是 {{side, cum, chg}}")
        side = parse_side(v.get("side"))
        cum = to_int(v.get("cum"), f"{label}{key} 累计")
        chg = to_int(v.get("chg"), f"{label}{key} 增减")
        if cum is None:
            die(f"{label}「{key}」缺累计持仓")
        if chg is None:
            die(f"{label}「{key}」缺当日增减")
        out[key] = Rec(key, side, cum, chg)
    return out


def fmt_signed(v):
    if v == 0:
        return "0"
    return ("+" if v > 0 else "-") + str(abs(v))


def single_line(r, flag):
    s = "{}净{} {}（累计净{} {}）".format(
        r.name, r.dirword, fmt_signed(r.chg), r.dirword, abs(r.cum_s))
    if r.flip:
        # 翻向优先：不再叠加占比标注
        s += "（{}）".format(r.flip)
    elif flag and r.ratio is not None and r.ratio > RATIO_LIMIT:
        s += "（当日增减占比 {:.1f}%，超10%）".format(r.ratio * 100)
    return s


def combined(recs):
    cum = sum(r.cum_s for r in recs)
    d = sum(r.dlong for r in recs)
    prev = cum - d
    ratio = abs(d) / float(abs(prev)) if prev else None
    return cum, d, ratio


def combined_line(label, recs, flag):
    cum, d, ratio = combined(recs)
    dirword = "多" if cum > 0 else "空"
    # 合计的增减也按「合计方向」表达：净多写多增/多减，净空写空增/空减
    shown = d if cum > 0 else -d
    s = "{}当日净{} {}，合计净{}：{}".format(label, dirword, fmt_signed(shown),
                                          dirword, abs(cum))
    if flag and ratio is not None and ratio > RATIO_LIMIT:
        s += "（当日增减占比 {:.1f}%，超10%）".format(ratio * 100)
    return s


def ct_label(v, default):
    """合约号统一成纯数字：M2701 → 2701、RM2701 → 2701。"""
    s = str(v or "").strip()
    if not s:
        return default
    m = re.match(r"^[A-Za-z]*(\d{3,4})$", s)
    return m.group(1) if m else s


def variety_of(obj):
    """对象显示名去掉末尾数字：豆粕01 → 豆粕。"""
    name = str(obj.get("name") or "")
    return re.sub(r"\d+$", "", name).strip() or name


def block(store, orgs_order, combos, flag, flag_comb):
    """一个合约的正文：按 orgs_order 逐行，combo 在 after 之后插合计行，无 after 的放段末。"""
    after_map = {}
    tail_combos = []
    for c in combos:
        members = [m for m in (c.get("members") or []) if m]
        if len(members) < 2:
            continue
        label = c.get("label") or " + ".join(members)
        after = c.get("after")
        if after:
            after_map.setdefault(after, []).append((label, members))
        else:
            tail_combos.append((label, members))

    def need(name):
        if name not in store:
            die("数据缺少「{}」，无法生成报告".format(name))
        return store[name]

    out = []
    for org in orgs_order:
        out.append(single_line(need(org), flag))
        for label, members in after_map.get(org, []):
            out.append(combined_line(label, [need(m) for m in members], flag_comb))
    for label, members in tail_combos:
        out.append(combined_line(label, [need(m) for m in members], flag_comb))
    return out


def build(data, group=None):
    objs = data.get("objects")
    if group:
        key = str(group).strip()
        objs = [o for o in (objs or [])
                if key in (o.get("group"), o.get("group_name"))]
    if not objs:
        if group:
            die("分组 {} 在 JSON 里没有合约数据".format(group))
        die("JSON 缺少 objects 数组：请用新版 fetch_positions.py 取数（合约/机构由 config.json 配置）")

    d = data.get("date")
    if d:
        try:
            dd = datetime.strptime(str(d), "%Y-%m-%d").date()
        except ValueError:
            die(f"date 应为 YYYY-MM-DD：{d!r}")
    else:
        dd = date.today()

    flag = bool(data.get("flag_over_10", True))
    flag_comb = bool(data.get("flag_combined", False))
    allow_missing = bool(data.get("allow_missing"))

    lines = ["{} 月 {} 日持仓变化".format(dd.month, dd.day)]
    numerals = "一二三四五六七八九十"
    for i, obj in enumerate(objs):
        name = obj.get("name") or obj.get("contract") or "合约"
        ct = ct_label(obj.get("contract"), "")
        title = "{} {} 合约".format(variety_of(obj), ct) if ct else "{} 合约".format(variety_of(obj))
        if obj.get("primary"):
            title += "（主力）"
        orgs_order = obj.get("orgs_order") or list((obj.get("orgs") or {}).keys())
        store = load_section(obj.get("orgs") or {}, name, set(orgs_order))
        missing = [n for n in orgs_order if n not in store]
        if missing and not allow_missing:
            die("{} 缺这些机构：{}（如确实没有该机构，设 allow_missing=true）".format(
                name, "、".join(missing)))
        lines.append("{}、{}：".format(numerals[i] if i < len(numerals) else str(i + 1), title))
        lines.extend(block(store, orgs_order, obj.get("combos") or [], flag, flag_comb))
    return "\n".join(lines)


SELF_TEST_DATA = {
    "date": "2025-05-11",
    "objects": [
        {
            "contract": "M2601", "name": "豆粕01", "primary": True,
            "orgs_order": ["高盛", "摩根", "中粮", "国泰君安", "东证", "中信"],
            "combos": [
                {"members": ["高盛", "摩根"], "after": "中粮", "label": "高盛 + 摩根"},
                {"members": ["国泰君安", "东证", "中信"], "label": "国泰 + 东证 + 中信"},
            ],
            "orgs": {
                "高盛期货": {"side": "long", "cum": 66570, "chg": 17124},
                "摩根大通": {"side": "short", "cum": 26855, "chg": -13842},
                "中粮期货": {"side": "short", "cum": 463034, "chg": 6637},
                "国泰君安": {"side": "long", "cum": 145360, "chg": 11241},
                "东证期货": {"side": "long", "cum": 86913, "chg": -8536},
                "中信期货": {"side": "long", "cum": 26218, "chg": 18042},
            },
        },
        {
            "contract": "RM2701", "name": "菜粕01", "primary": False,
            "orgs_order": ["高盛", "摩根", "中粮"],
            "combos": [{"members": ["高盛", "摩根"], "label": "高盛 + 摩根"}],
            "orgs": {
                "高盛": {"side": "long", "cum": 38043, "chg": 3302},
                "摩根": {"side": "short", "cum": 17397, "chg": -3701},
                "中粮": {"side": "short", "cum": 48623, "chg": 1680},
            },
        },
    ],
}

SELF_TEST_EXPECT = [
    "5 月 11 日持仓变化",
    "一、豆粕 2601 合约（主力）：",
    "高盛净多 +17124（累计净多 66570）",
    "摩根净空 -13842（累计净空 26855）",
    "中粮净空 +6637（累计净空 463034）",
    "高盛 + 摩根当日净多 +30966，合计净多：39715",
    "国泰君安净多 +11241（累计净多 145360）",
    "东证净多 -8536（累计净多 86913）",
    "中信净多 +18042（累计净多 26218）",
    "国泰 + 东证 + 中信当日净多 +20747，合计净多：258491",
    "二、菜粕 2701 合约：",
    "高盛净多 +3302（累计净多 38043）",
    "摩根净空 -3701（累计净空 17397）",
    "中粮净空 +1680（累计净空 48623）",
    "高盛 + 摩根当日净多 +7003，合计净多：20646",
]


def self_test():
    txt = build(dict(SELF_TEST_DATA))
    # 去掉 10% 标注后再逐行对案例
    import re
    body = [re.sub(r"（当日增减占比 [\d.]+%，超10%）$", "", l) for l in txt.split("\n")]
    ok = True
    for exp in SELF_TEST_EXPECT:
        if exp not in body:
            ok = False
            print("✗ 缺少：", exp)
    # 10% 标注：高盛 17124/49446=34.6%，中粮 6637/456397=1.5%（不标），摩根 13842/40697=34.0%
    if "高盛净多 +17124（累计净多 66570）（当日增减占比 34.6%，超10%）" not in txt:
        ok = False
        print("✗ 高盛 10% 标注不对")
    if "中粮净空 +6637（累计净空 463034）（" in txt:
        ok = False
        print("✗ 中粮 1.5% 不应标注")

    # 合计为净空时，增减要按「空增/空减」表达（2026-10-28 实盘数据回归）
    net_short = {
        "date": "2026-10-28",
        "objects": [
            {
                "contract": "M2701", "name": "豆粕01", "primary": True,
                "orgs_order": ["高盛", "摩根", "中粮", "国泰君安", "东证", "中信"],
                "combos": [
                    {"members": ["高盛", "摩根"], "after": "中粮", "label": "高盛 + 摩根"},
                    {"members": ["国泰君安", "东证", "中信"], "label": "国泰 + 东证 + 中信"},
                ],
                "orgs": {
                    "高盛": {"side": "long", "cum": 114312, "chg": -5396},
                    "摩根": {"side": "short", "cum": 52571, "chg": 5494},
                    "中粮": {"side": "short", "cum": 331646, "chg": -5213},
                    "国泰君安": {"side": "long", "cum": 52142, "chg": -4963},
                    "东证": {"side": "short", "cum": 10666, "chg": -3518},
                    "中信": {"side": "long", "cum": 109258, "chg": 5990},
                },
            },
            {
                "contract": "RM2701", "name": "菜粕01",
                "orgs_order": ["高盛", "摩根", "中粮"],
                "combos": [{"members": ["高盛", "摩根"], "label": "高盛 + 摩根"}],
                "orgs": {
                    "高盛": {"side": "short", "cum": 62705, "chg": 5665},
                    "摩根": {"side": "short", "cum": 8461, "chg": 6914},
                    "中粮": {"side": "short", "cum": 28665, "chg": -8725},
                },
            },
        ],
    }
    txt2 = build(net_short)
    for exp in ["高盛 + 摩根当日净多 -10890，合计净多：61741",
                "国泰 + 东证 + 中信当日净多 +4545，合计净多：150734",
                "高盛 + 摩根当日净空 +12579，合计净空：71166"]:
        if exp not in txt2:
            ok = False
            print("✗ 净空合计回归失败，缺：", exp)
    # 加入豆粕05 后，段落编号要顺延：二、豆粕 2705，三、菜粕 2701
    import copy
    m5 = copy.deepcopy(SELF_TEST_DATA["objects"][0])
    m5["contract"] = "M2705"
    m5["name"] = "豆粕05"
    m5["primary"] = False
    with5 = {
        "date": SELF_TEST_DATA["date"],
        "objects": [
            copy.deepcopy(SELF_TEST_DATA["objects"][0]),
            m5,
            copy.deepcopy(SELF_TEST_DATA["objects"][1]),
        ],
    }
    body5 = build(with5).split("\n")
    for exp in ["一、豆粕 2601 合约（主力）：", "二、豆粕 2705 合约：",
                "三、菜粕 2701 合约："]:
        if exp not in body5:
            ok = False
            print("✗ 含豆粕05 时缺：", exp)
    if "二、菜粕 2701 合约：" in body5:
        ok = False
        print("✗ 含豆粕05 时菜粕不该再是「二、」")
    # 分组过滤：只保留目标分组的合约，编号从「一、」重新开始
    grp = copy.deepcopy(with5)
    for o in grp["objects"]:
        o["group"] = "meal_rape"
        o["group_name"] = "豆粕菜粕"
    grp["objects"][1]["group"] = "soy_only"
    grp["objects"][1]["group_name"] = "大豆"
    bodyg = build(grp, group="soy_only")
    if "一、豆粕 2705 合约：" not in bodyg:
        ok = False
        print("✗ 分组过滤后缺「一、豆粕 2705 合约：」")
    for bad in ["豆粕 2601 合约", "菜粕 2701 合约"]:
        if bad in bodyg:
            ok = False
            print("✗ 分组过滤后不该出现：", bad)
    bodygn = build(grp, group="豆粕菜粕")
    if "菜粕 2701 合约" not in bodygn or "豆粕 2705 合约" in bodygn:
        ok = False
        print("✗ 按分组名称过滤结果不对")

    if not ok:
        print("\n--- 实际输出 ---\n" + txt)
        sys.exit(1)
    print("✅ 自检通过（与案例数字逐行一致）\n")
    print(txt)


def main():
    ap = argparse.ArgumentParser(description="机构持仓报告生成器")
    ap.add_argument("json", nargs="?", help="数据 JSON 路径，- 表示从标准输入读")
    ap.add_argument("--date", help="覆盖日期 YYYY-MM-DD（默认取 JSON 或今天）")
    ap.add_argument("--group", help="只生成指定分组的报告（分组 id 或名称，如 meal_rape / 豆粕菜粕）")
    ap.add_argument("--flag-combined", action="store_true",
                    help="合计行也标注超 10%% 的增减占比")
    ap.add_argument("--no-flag", action="store_true", help="不做 10%% 标注")
    ap.add_argument("--self-test", action="store_true", help="用案例数字自检")
    args = ap.parse_args()

    if args.self_test:
        self_test()
        return

    if not args.json:
        ap.error("缺少数据 JSON")
    if args.json == "-":
        data = json.load(sys.stdin)
    else:
        with open(args.json, "r", encoding="utf-8") as f:
            data = json.load(f)

    if args.date:
        data["date"] = args.date
    if args.flag_combined:
        data["flag_combined"] = True
    if args.no_flag:
        data["flag_over_10"] = False

    print(build(data, group=args.group))


if __name__ == "__main__":
    # Windows 控制台默认 GBK，直接 reconfigure 避免 ✅/❌ 等字符报错
    for _s in (sys.stdout, sys.stderr):
        try:
            _s.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    main()
