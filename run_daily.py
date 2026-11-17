# -*- coding: utf-8 -*-
"""机构持仓日更 · 一键运行器（便携版）

用法：
    python run_daily.py                 # 取当天 → 写表 → 出两张图
    python run_daily.py --date 2026-11-11
    python run_daily.py --no-write      # 只取数出图，不动表格
    python run_daily.py --no-chart      # 只取数与写表，不出图
    python run_daily.py --dry-run       # 写表只预览不落盘
    python run_daily.py --group palm    # 只跑一个分组

按 config.json 的 groups 分组产出：每个分组一套「报告文字 txt + 报告图 PNG + 折线图 PNG」，
文件名带分组名；多分组时 CSV 仍只写一次（一个合约一个文件）。

返回码：
    0 成功  1 依赖缺失  2 配置问题  3 子脚本失败
    6 当日数据未更新（东财还没出当天龙虎榜，稍后再跑）
    7 表内第 2 行已是这一天（今天已经跑过）
"""
import argparse
import json
import os
import subprocess
import sys
from datetime import date, datetime

def _console_enc():
    """stdout 接到真实控制台时按代码页输出（cmd 默认 GBK）；
    被重定向/管道时一律 UTF-8，避免乱码。"""
    if os.name != "nt":
        return "utf-8"
    try:
        if not sys.stdout.isatty():
            return "utf-8"
        import ctypes
        return "utf-8" if ctypes.windll.kernel32.GetConsoleOutputCP() == 65001 else "gbk"
    except Exception:
        return "utf-8"


if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding=_console_enc(), errors="replace")
        sys.stderr.reconfigure(encoding=_console_enc(), errors="replace")
    except Exception:
        pass

APP = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(APP, "scripts")
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)

import config_loader  # pyright: ignore[reportImplicitRelativeImport]
import window_data  # pyright: ignore[reportImplicitRelativeImport]
CONFIG_PATH = os.path.join(APP, "config.json")

DEFAULT_CFG = {
    "data_dir": "data",
    "out_dir": "../输出",
    "write_csv": True,
}

REQUIRED = [
    ("matplotlib", "折线图"),
    ("PIL", "报告图 / 图片合并"),
]


def log(msg=""):
    print(msg, flush=True)


def step(n, total, title):
    log("")
    log("=" * 58)
    log("  [{}/{}] {}".format(n, total, title))
    log("=" * 58)


def load_cfg():
    cfg = dict(DEFAULT_CFG)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception as e:
            log("[警告] config.json 读取失败，用默认配置：{}".format(e))
    return cfg


def resolve(p):
    """相对路径按 app 目录解析，绝对路径原样返回。"""
    if not p:
        return p
    return p if os.path.isabs(p) else os.path.normpath(os.path.join(APP, p))


def check_deps():
    missing = []
    for mod, why in REQUIRED:
        try:
            __import__(mod)
        except ImportError:
            missing.append((mod, why))
    return missing


def call(script, args, total, idx, title, allow=()):
    """跑一个子脚本，返回其返回码。allow 里的非零码不算失败。

    子进程强制 UTF-8 输出并在此转码：子脚本里带 emoji 的提示在 GBK 控制台
    直接 print 会抛 UnicodeEncodeError，必须由父进程兜住。
    """
    step(idx, total, title)
    cmd = [sys.executable, os.path.join(SCRIPTS, script)] + args
    log("  > {}".format(" ".join(cmd[1:])))
    log("")
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    r = subprocess.run(cmd, cwd=SCRIPTS, env=env, capture_output=True)
    out = (r.stdout or b"").decode("utf-8", "replace") + (r.stderr or b"").decode("utf-8", "replace")
    if out.strip():
        log(out.rstrip())
    if r.returncode == 0:
        log("")
        log("  [OK] {} 完成".format(script))
    elif r.returncode in allow:
        log("")
        log("  [注意] {} 返回 {}（可接受）".format(script, r.returncode))
    else:
        log("")
        log("  [失败] {} 返回码 {}".format(script, r.returncode))
    return r.returncode


def main():
    ap = argparse.ArgumentParser(description="机构持仓日更一键运行")
    ap.add_argument("--date", help="数据日期 YYYY-MM-DD，默认最新交易日")
    ap.add_argument("--data-dir", help="CSV 目录，覆盖 config.json")
    ap.add_argument("--out-dir", help="图片输出目录，覆盖 config.json")
    ap.add_argument("--no-write", action="store_true", help="不写 CSV，只取数出图")
    ap.add_argument("--no-chart", action="store_true", help="不出图")
    ap.add_argument("--dry-run", action="store_true", help="写 CSV 只预览")
    ap.add_argument("--json", dest="json_out", help="取数 JSON 另存路径（默认放输出目录）")
    ap.add_argument("--group", help="只跑指定分组（config.json 的分组 id 或名称）")
    a = ap.parse_args()

    cfg = load_cfg()
    data_dir = resolve(a.data_dir or cfg.get("data_dir"))
    out_dir = resolve(a.out_dir or cfg.get("out_dir"))
    os.makedirs(data_dir, exist_ok=True)
    os.makedirs(out_dir, exist_ok=True)

    log("=" * 58)
    log("  机构持仓日更 · 一键运行")
    log("=" * 58)
    log("  CSV 目录: {}".format(data_dir))
    log("  输出目录: {}".format(out_dir))
    log("  时间    : {}".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))

    do_write = (not a.no_write) and bool(cfg.get("write_csv", True))

    # ---- 依赖 ----
    missing = check_deps()
    if missing:
        log("")
        log("  [失败] 缺少依赖：")
        for mod, why in missing:
            log("     - {}（{}）".format(mod, why))
        log("")
        log("  请双击「3-安装依赖.bat」，或手动执行：")
        log("     {} -m pip install -r requirements.txt".format(os.path.basename(sys.executable)))
        return 1

    # ---- 分组（每组一套报告文字 + 报告图 + 折线图）----
    all_groups = config_loader.get_groups(cfg)
    if not all_groups:
        log("")
        log("  [失败] config.json 里没有 groups / objects，无法确定要出哪些分组")
        return 2
    groups = all_groups
    if a.group:
        one = config_loader.find_group(cfg, a.group)
        if not one:
            log("")
            log("  [失败] config.json 里没有分组：{}".format(a.group))
            return 2
        groups = [one]
    split = len(all_groups) > 1

    def suffix_of(g):
        """多分组时文件名带分组名；单分组保持旧文件名。"""
        return "_" + str(g.get("name") or g.get("id")) if split else ""

    log("  分组    : {}".format(" / ".join(
        g["name"] + "（{} 个合约）".format(len(g["objects"])) for g in groups)))

    total = 1 + len(groups) + (1 if do_write else 0) + (2 * len(groups) if not a.no_chart else 0)
    idx = 0

    # ---- 1 取数 ----
    latest = window_data.latest_trade_date()
    stamp = a.date or latest.isoformat()
    json_path = a.json_out or os.path.join(out_dir, "positions_{}.json".format(stamp.replace("-", "")))
    fetch_args = ["--out", json_path]
    if a.date:
        fetch_args += ["--date", a.date]
    else:
        fetch_args += ["--date", latest.isoformat()]
        log("  最新数据日: {}（16:00 规则自动判定）".format(latest.isoformat()))

    idx += 1
    rc = call("fetch_positions.py", fetch_args, total, idx, "取数（东方财富龙虎榜）", allow=(6,))
    if rc == 6:
        log("")
        log("  [等待] 东财今天的数据还没更新。稍后（一般 16:00 后）再跑一次即可。")
        return 6
    if rc != 0:
        return 3

    # ---- 2 报告文字（按分组）----
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        log("  [失败] 读取取数 JSON 失败：{}".format(e))
        return 3

    for g in groups:
        idx += 1
        rep_args = [json_path] + (["--date", a.date] if a.date else [])
        if split or a.group:
            rep_args += ["--group", g["id"]]
        rc = call("compute_report.py", rep_args, total, idx,
                  "生成报告文字 · {}".format(g["name"]))
        if rc != 0:
            return 3

        # 报告落盘一份 txt，方便直接复制
        txt_path = os.path.join(out_dir, "机构持仓报告_{}{}.txt".format(
            stamp.replace("-", ""), suffix_of(g)))
        try:
            sys.path.insert(0, SCRIPTS)
            import compute_report
            text = compute_report.build(data, group=g["id"] if (split or a.group) else None)
            if a.date:
                text = text.replace(text.split("\n")[0], "{} 月 {} 日持仓变化".format(
                    int(a.date[5:7]), int(a.date[8:10])), 1)
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(text + "\n")
        except Exception as e:
            log("  （报告 txt 落盘失败，不影响后续：{}）".format(e))

    # ---- 3 写表 ----
    if do_write:
        idx += 1
        w_args = [json_path, "--data-dir", data_dir]
        if a.dry_run:
            w_args.append("--dry-run")
        rc = call("write_positions.py", w_args, total, idx, "写入 CSV", allow=(5,))
        if rc == 5:
            log("")
            log("  [提示] CSV 里这一天已存在，本次跳过写入（今天已经跑过）。")
        elif rc not in (0, 5):
            return 3

    # ---- 4 报告图（按分组）----
    if not a.no_chart:
        for g in groups:
            idx += 1
            img_args = [json_path, "--out", os.path.join(
                out_dir, "机构持仓报告_{}{}.png".format(stamp.replace("-", ""), suffix_of(g)))]
            if a.date:
                img_args += ["--date", a.date]
            if split or a.group:
                img_args += ["--group", g["id"]]
            call("render_report_image.py", img_args, total, idx,
                 "生成报告图（白底黑字）· {}".format(g["name"]))

    # ---- 5 三联折线图（按分组）----
    if not a.no_chart:
        for g in groups:
            idx += 1
            c_args = ["--data-dir", data_dir, "--out-dir", out_dir]
            if a.date:
                c_args += ["--date", a.date]
            if split or a.group:
                c_args += ["--group", g["id"]]
            call("render_charts.py", c_args, total, idx,
                 "生成三联折线图 · {}".format(g["name"]))

    # ---- 汇总 ----
    log("")
    log("=" * 58)
    log("  完成，产出在：{}".format(out_dir))
    log("=" * 58)
    for fn in sorted(os.listdir(out_dir)):
        if stamp.replace("-", "") in fn:
            log("  · {}".format(os.path.join(out_dir, fn)))
    log("")
    return 0


if __name__ == "__main__":
    sys.exit(main())
