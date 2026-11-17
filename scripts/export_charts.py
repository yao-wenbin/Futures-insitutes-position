# -*- coding: utf-8 -*-
"""导出「机构持仓表.xlsm」Sheet1 的豆粕01 / 菜粕01 两张图表，并左右合并成一张 PNG。

依赖：Excel COM（pywin32）、Pillow。
只读操作：不修改工作簿；文件被占用时自动以只读方式打开。

用法：
    python export_charts.py                          # 图表 1 + 图表 6，输出到「图表数据」目录
    python export_charts.py --date 2026-10-28        # 指定文件名里的日期
    python export_charts.py --charts 1,6 --out-dir 图表数据
返回码：0 成功；1 导出/合并失败；2 参数或路径问题；3 文件被占用且无法只读打开。
"""
import argparse
import datetime
import os
import sys

DEFAULT_FILE = os.environ.get("INST_HOLDINGS_XLSM") or r"D:/WorkBuddyWorkSpace/鲁信工作/机构持仓表.xlsm"
DEFAULT_DIR = os.environ.get("INST_HOLDINGS_OUT") or r"D:/WorkBuddyWorkSpace/鲁信工作/图表数据"
# 默认按标题找，比写死索引稳（插行、复制图表后索引会变）
# 顺序即合并顺序：左 豆粕01 → 中 豆粕05 → 右 菜粕01
DEFAULT_CHARTS = "豆粕01持仓,豆粕05持仓（次主力）,菜粕01持仓"
GAP = 18                        # 图与图之间留白
MARGIN = 14                     # 四周留白
TARGET_H = 700                  # 合并后统一高度（等比缩放）


def pick_chart(ws, spec):
    """spec 可以是索引（"1"）或标题关键词（"豆粕01持仓"）。"""
    total = ws.ChartObjects().Count
    try:
        idx = int(str(spec))
    except ValueError:
        idx = None
    if idx is not None:
        if 1 <= idx <= total:
            return ws.ChartObjects(idx), None
        return None, "图表{} 不存在（共 {} 张）".format(idx, total)
    for i in range(1, total + 1):
        co = ws.ChartObjects(i)
        try:
            t = co.Chart.ChartTitle.Text if co.Chart.HasTitle else ""
        except Exception:
            t = ""
        if str(spec) in str(t):
            return co, None
    return None, "找不到标题含「{}」的图表（共 {} 张）".format(spec, total)



def export(xlsm, charts, out_dir):
    import win32com.client as win32

    paths = []
    xl = None
    wb = None
    try:
        xl = win32.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False
        xl.AutomationSecurity = 3
        try:
            wb = xl.Workbooks.Open(os.path.abspath(xlsm), UpdateLinks=0)
        except Exception:
            wb = xl.Workbooks.Open(os.path.abspath(xlsm), UpdateLinks=0, ReadOnly=True)

        ws = None
        for s in wb.Worksheets:
            try:
                if str(s.Range("A1").Value or "").strip() == "豆粕01":
                    ws = s
                    break
            except Exception:
                continue
        if ws is None:
            print("✗ 找不到 A1=豆粕01 的工作表", file=sys.stderr)
            return None, 2

        total = ws.ChartObjects().Count
        for spec in charts:
            co, err = pick_chart(ws, spec)
            if err:
                print("✗ " + err, file=sys.stderr)
                return None, 2
            try:
                title = co.Chart.ChartTitle.Text
            except Exception:
                title = "图表"
            safe = str(title).replace("/", "_").replace("\\", "_")
            f = os.path.join(out_dir, "{}.png".format(safe))
            co.Chart.Export(f, "PNG")
            paths.append(f)
            print("  《{}》 → {}".format(title, os.path.basename(f)))
        return paths, 0
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass


def merge(paths, out_path):
    from PIL import Image

    ims = []
    for p in paths:
        im = Image.open(p).convert("RGBA")
        # 去掉 PNG 的透明边（导出图是白底，直接铺白更稳）
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        bg.paste(im, (0, 0), im)
        ims.append(bg.convert("RGB"))

    scaled = []
    for im in ims:
        w, h = im.size
        nw = max(1, int(round(w * TARGET_H / h)))
        scaled.append(im.resize((nw, TARGET_H), Image.LANCZOS))

    total_w = sum(i.size[0] for i in scaled) + GAP * (len(scaled) - 1) + MARGIN * 2
    total_h = TARGET_H + MARGIN * 2
    canvas = Image.new("RGB", (total_w, total_h), (255, 255, 255))
    x = MARGIN
    for im in scaled:
        canvas.paste(im, (x, MARGIN))
        x += im.size[0] + GAP
    canvas.save(out_path, "PNG", optimize=True)
    return canvas.size


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", default=DEFAULT_FILE)
    ap.add_argument("--out-dir", default=DEFAULT_DIR)
    ap.add_argument("--charts", default=DEFAULT_CHARTS,
                    help="逗号分隔的图表索引或标题关键词，按顺序合并")
    ap.add_argument("--date", default=None, help="输出文件名日期，默认今天")
    ap.add_argument("--keep-parts", action="store_true", help="保留单张图片")
    a = ap.parse_args()

    if not os.path.isfile(a.file):
        print("✗ 找不到工作簿：{}".format(a.file), file=sys.stderr)
        return 2
    os.makedirs(a.out_dir, exist_ok=True)

    charts = [x.strip() for x in str(a.charts).split(",") if x.strip()]
    if not charts:
        print("✗ --charts 不能为空", file=sys.stderr)
        return 2

    dd = datetime.date.today()
    if a.date:
        try:
            dd = datetime.date.fromisoformat(a.date)
        except ValueError:
            print("✗ --date 需要 YYYY-MM-DD", file=sys.stderr)
            return 2
    tag = dd.strftime("%Y%m%d")

    print("导出图表 …")
    try:
        paths, code = export(a.file, charts, a.out_dir)
    except Exception as e:
        print("✗ 导出失败：{}".format(e), file=sys.stderr)
        return 1
    if code:
        return code

    out = os.path.join(a.out_dir, "机构持仓_{}.png".format(tag))
    try:
        size = merge(paths, out)
    except Exception as e:
        print("✗ 合并失败：{}".format(e), file=sys.stderr)
        return 1
    print("\n✅ 合并完成：{}  （{}×{}）".format(out, size[0], size[1]))

    if not a.keep_parts:
        for p in paths:
            try:
                os.remove(p)
            except Exception:
                pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
