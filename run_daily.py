# -*- coding: utf-8 -*-
"""机构持仓日更 · 一键运行器（便携版）

用法：
    python run_daily.py                 # 取当天 → 写表 → 出两张图
    python run_daily.py --date 2026-11-11
    python run_daily.py --no-write      # 只取数出图，不动表格
    python run_daily.py --no-chart      # 只取数与写表，不出图
    python run_daily.py --dry-run       # 写表只预览不落盘

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
CONFIG_PATH = os.path.join(APP, "config.json")

DEFAULT_CFG = {
    "xlsm": "机构持仓表.xlsm",
    "out_dir": "../输出",
    "write_sheet": True,
    "meal_code": "M2701",
    "meal05_code": "M2705",
    "rape_code": "RM2701",
}

REQUIRED = [
    ("openpyxl", "读表"),
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


def check_deps(need_win32):
    missing = []
    for mod, why in REQUIRED:
        try:
            __import__(mod)
        except ImportError:
            missing.append((mod, why))
    if need_win32:
        try:
            __import__("win32com.client")
        except ImportError:
            missing.append(("pywin32", "写表（Excel COM）"))
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
    ap.add_argument("--xlsm", help="工作簿路径，覆盖 config.json")
    ap.add_argument("--out-dir", help="图片输出目录，覆盖 config.json")
    ap.add_argument("--no-write", action="store_true", help="不写表，只取数出图")
    ap.add_argument("--no-chart", action="store_true", help="不出图")
    ap.add_argument("--dry-run", action="store_true", help="写表只预览")
    ap.add_argument("--json", dest="json_out", help="取数 JSON 另存路径（默认放输出目录）")
    a = ap.parse_args()

    cfg = load_cfg()
    xlsm = resolve(a.xlsm or cfg.get("xlsm"))
    out_dir = resolve(a.out_dir or cfg.get("out_dir"))
    os.makedirs(out_dir, exist_ok=True)

    log("=" * 58)
    log("  机构持仓日更 · 一键运行")
    log("=" * 58)
    log("  工作簿  : {}".format(xlsm))
    log("  输出目录: {}".format(out_dir))
    log("  时间    : {}".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))

    have_book = bool(xlsm) and os.path.exists(xlsm)
    do_write = (not a.no_write) and bool(cfg.get("write_sheet", True)) and have_book
    if not have_book:
        log("  [注意] 找不到工作簿，本次只输出报告文字 + 报告图（不写表、不出折线图）")

    # ---- 依赖 ----
    missing = check_deps(need_win32=do_write)
    if missing:
        log("")
        log("  [失败] 缺少依赖：")
        for mod, why in missing:
            log("     - {}（{}）".format(mod, why))
        log("")
        log("  请双击「3-安装依赖.bat」，或手动执行：")
        log("     {} -m pip install -r requirements.txt".format(os.path.basename(sys.executable)))
        return 1

    lock = os.path.join(os.path.dirname(xlsm), "~$" + os.path.basename(xlsm))
    if have_book and os.path.exists(lock):
        log("")
        log("  [失败] 工作簿正被 Excel 打开（存在 {}）".format(os.path.basename(lock)))
        log("     请先关闭 Excel 再运行。")
        return 2

    total = 5 if (do_write and not a.no_chart) else 4
    idx = 0

    # ---- 1 取数 ----
    stamp = a.date or date.today().isoformat()
    json_path = a.json_out or os.path.join(out_dir, "positions_{}.json".format(stamp.replace("-", "")))
    fetch_args = ["--out", json_path,
                  "--meal-code", cfg.get("meal_code", "M2701"),
                  "--meal05-code", cfg.get("meal05_code", "M2705"),
                  "--rape-code", cfg.get("rape_code", "RM2701")]
    if a.date:
        fetch_args += ["--date", a.date]
    else:
        fetch_args.append("--require-today")   # 定时场景：只要当天的

    idx += 1
    rc = call("fetch_positions.py", fetch_args, total, idx, "取数（东方财富龙虎榜）", allow=(6,))
    if rc == 6:
        log("")
        log("  [等待] 东财今天的数据还没更新。稍后（一般 16:00 后）再跑一次即可。")
        return 6
    if rc != 0:
        return 3

    # ---- 2 报告 ----
    idx += 1
    rep_args = [json_path] + (["--date", a.date] if a.date else [])
    rc = call("compute_report.py", rep_args, total, idx, "生成报告文字")
    if rc != 0:
        return 3

    # 报告落盘一份 txt，方便直接复制
    txt_path = os.path.join(out_dir, "机构持仓报告_{}.txt".format(stamp.replace("-", "")))
    try:
        sys.path.insert(0, SCRIPTS)
        import compute_report
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
        text = compute_report.build(data)
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
        w_args = [json_path, "--file", xlsm]
        if a.dry_run:
            w_args.append("--dry-run")
        rc = call("write_positions.py", w_args, total, idx, "写入表格第 2 行", allow=(5,))
        if rc == 5:
            log("")
            log("  [提示] 表里第 2 行已经是这一天，本次跳过写入（今天已经跑过）。")
        elif rc not in (0, 5):
            return 3

    # ---- 4 报告图 ----
    if not a.no_chart:
        idx += 1
        img_args = [json_path, "--out", os.path.join(
            out_dir, "机构持仓报告_{}.png".format(stamp.replace("-", "")))]
        if a.date:
            img_args += ["--date", a.date]
        call("render_report_image.py", img_args, total, idx, "生成报告图（白底黑字）")

    # ---- 5 三联折线图 ----
    if not a.no_chart and have_book:
        idx += 1
        c_args = ["--file", xlsm, "--out-dir", out_dir]
        if a.date:
            c_args += ["--date", a.date]
        call("render_charts.py", c_args, total, idx, "生成三联折线图")

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
