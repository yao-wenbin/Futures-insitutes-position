# -*- coding: utf-8 -*-
"""统一配置加载：读 config.json 的 groups / objects。

config.json 结构（分组版，用户 2026-11-17 指定）：

    {
      "data_dir": "data",
      "out_dir": "../输出",
      "groups": [
        {
          "id": "meal_rape",
          "name": "豆粕菜粕",
          "objects": [
            {
              "contract": "M2701", "name": "豆粕01", "primary": true,
              "orgs": ["高盛","摩根","中粮","国泰君安","东证","中信"],
              "combos": [{ "members": ["高盛","摩根"], "after": "中粮", "label": "高盛 + 摩根" }],
              "chart_orgs": ["高盛","摩根","国泰君安","东证","中信"]
            }
          ]
        },
        { "id": "palm", "name": "棕榈油", "objects": [ ... ] }
      ]
    }

每个分组产出独立的报告与图片；组内 objects 即该组包含的合约。
兼容旧版：没有 groups、只有 objects 时，视作单个分组（id=all）。

字段说明：
  contract   合约代码（决定 CSV 文件名：M2701 → m2701.csv）
  name       显示名（报告/图标题用）
  primary    是否主力（标题加「（主力）」）
  orgs       抓取/展示机构（顺序 = 报告行序 = CSV 明细列序）
  combos     合计组合：members 参与合计、after 插在该机构之后、label 报告标签、column CSV 列名
  chart_orgs 图里画哪些机构（省略 = orgs 去掉「中粮」）
"""
import json
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CONFIG_PATH = os.path.join(ROOT, "config.json")


def load_config(path=None):
    """读取 config.json；不存在或解析失败时返回空 dict。"""
    p = path or CONFIG_PATH
    if not os.path.exists(p):
        return {}
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def _normalize_object(o):
    if not isinstance(o, dict):
        return None
    contract = str(o.get("contract") or "").strip()
    if not contract:
        return None
    orgs = [str(x).strip() for x in (o.get("orgs") or []) if str(x).strip()]
    combos = []
    for c in (o.get("combos") or []):
        if not isinstance(c, dict):
            continue
        members = [str(x).strip() for x in (c.get("members") or []) if str(x).strip()]
        if len(members) < 2:
            continue
        combos.append({
            "members": members,
            "after": (c.get("after") or None),
            "label": c.get("label") or " + ".join(members),
            "column": c.get("column") or "+".join(members),
        })
    chart_orgs = [str(x).strip() for x in (o.get("chart_orgs") or []) if str(x).strip()]
    if not chart_orgs:
        chart_orgs = [x for x in orgs if x != "中粮"]
    return {
        "contract": contract,
        "name": str(o.get("name") or contract).strip(),
        "primary": bool(o.get("primary", False)),
        "orgs": orgs,
        "combos": combos,
        "chart_orgs": chart_orgs,
    }


def _parse_objects(lst):
    out = []
    for o in (lst or []):
        n = _normalize_object(o)
        if n:
            out.append(n)
    return out


def get_groups(cfg):
    """返回分组列表 [{id, name, objects}]；每个 object 带 _group / _group_name。"""
    groups = []
    raw = cfg.get("groups")
    if raw:
        for i, g in enumerate(raw):
            if not isinstance(g, dict):
                continue
            objs = _parse_objects(g.get("objects"))
            if not objs:
                continue
            gid = str(g.get("id") or "group{}".format(i + 1)).strip()
            name = str(g.get("name") or gid).strip()
            for o in objs:
                o["_group"] = gid
                o["_group_name"] = name
            groups.append({"id": gid, "name": name, "objects": objs})
    else:
        objs = _parse_objects(cfg.get("objects"))
        if objs:
            for o in objs:
                o["_group"] = "all"
                o["_group_name"] = "全部"
            groups.append({"id": "all", "name": "全部", "objects": objs})
    return groups


def get_objects(cfg):
    """扁平化的所有合约（跨分组），用于取数与写 CSV。"""
    out = []
    for g in get_groups(cfg):
        out.extend(g["objects"])
    return out


def find_group(cfg, key):
    """按分组 id 或名称查分组，返回 {id, name, objects}；找不到返回 None。"""
    k = str(key or "").strip()
    if not k:
        return None
    for g in get_groups(cfg):
        if k in (g["id"], g["name"]):
            return g
    return None


def group_name(cfg, key):
    """分组显示名：id 也能解析成中文名；解析不到时原样返回。"""
    g = find_group(cfg, key)
    return g["name"] if g else str(key or "").strip()


def combo_columns(obj):
    """CSV 列 = 明细机构 + 各 combo 的合计列。"""
    return list(obj["orgs"]) + [c["column"] for c in obj["combos"]]


def contract_digits(contract):
    """合约号取纯数字：M2701 → 2701、RM2701 → 2701。"""
    s = str(contract or "").strip()
    m = re.match(r"^[A-Za-z]*(\d{3,4})$", s)
    return m.group(1) if m else s


def variety_name(obj):
    """对象显示名去掉末尾数字：豆粕01 → 豆粕。"""
    return re.sub(r"\d+$", "", str(obj.get("name") or "")).strip() or obj["contract"]


def resolve_path(cli, cfg_value, default):
    """优先级：命令行 > 配置值 > 默认值；相对路径按项目根解析。"""
    base = cli or cfg_value or default
    if not base:
        return base
    if not os.path.isabs(base):
        base = os.path.normpath(os.path.join(ROOT, base))
    return base


def data_dir(cli, cfg):
    return resolve_path(cli, cfg.get("data_dir"), "data")


def out_dir(cli, cfg):
    return resolve_path(cli, cfg.get("out_dir"), os.path.join("..", "输出"))
