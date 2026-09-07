#!/usr/bin/env python3
"""在执行合同清单报告生成器。

用法:
    python make_report.py --input <guancli导出的JSON> --base-date 2026-08-31 \
        --store 弘阳家居南京江北店 --output <输出.xlsx> [--snapshot-time "2026-09-04 02:01"]

输入 JSON 为 `guancli ds preview k3a8a2772ee4143d694a2ecd --filter ... -f json` 的输出
（行记录数组，一行一合同）。输出 xlsx 含"汇总"与"合同明细"两个 sheet。
口径:在执行 = 起租日 <= 基准日 <= 到期日(日期窗口判定),含正铺+多经;合同状态仅展示。
"""
import argparse
import json
import sys
from collections import Counter
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DETAIL_COLS = [
    "合同号", "铺位号", "铺位类型", "合同状态", "租户名称", "品牌", "业态", "主品类",
    "楼层", "起租日", "到期日", "计租方式", "计租面积", "租金单价", "租金净单价",
    "合同总金额", "应收金额（含税）", "已收金额（含税）", "欠缴金额（含税）",
    "应收保证金", "实收保证金", "保证金余额", "招商人员姓名", "开业日期",
]
NUM_COLS = {
    "计租面积", "租金单价", "租金净单价", "合同总金额", "应收金额（含税）",
    "已收金额（含税）", "欠缴金额（含税）", "应收保证金", "实收保证金", "保证金余额",
}
WIDTHS = {
    "合同号": 15, "铺位号": 13, "铺位类型": 10, "合同状态": 10, "租户名称": 26,
    "品牌": 16, "业态": 10, "主品类": 12, "楼层": 8, "起租日": 12, "到期日": 12,
    "计租方式": 10, "计租面积": 11, "租金单价": 10, "租金净单价": 11,
    "合同总金额": 13, "应收金额（含税）": 14, "已收金额（含税）": 14,
    "欠缴金额（含税）": 14, "应收保证金": 12, "实收保证金": 12, "保证金余额": 12,
    "招商人员姓名": 12, "开业日期": 12,
}

H_FONT = Font(name="PingFang SC", size=11, bold=True, color="FFFFFF")
H_FILL = PatternFill("solid", fgColor="1F4E79")
B_FONT = Font(name="PingFang SC", size=11)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def num(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def clean(v, col):
    if col in ("起租日", "到期日") and v:
        return str(v)[:10]
    if col == "开业日期" and v:
        return str(v)[:10]
    if col in NUM_COLS:
        return num(v)
    return v


def expire_bucket(row, base):
    d = str(row.get("到期日") or "")[:10]
    if not d:
        return "无到期日"
    try:
        days = (date.fromisoformat(d) - base).days
    except ValueError:
        return "无到期日"
    if days < 0:
        return "已早于基准日"
    if days <= 30:
        return "基准日后30天内到期"
    if days <= 90:
        return "基准日后31-90天到期"
    if days <= 365:
        return "基准日后91-365天到期"
    return "一年以上后到期"


def build(input_path, base_date_str, store, output, snapshot_time, pos_note):
    with open(input_path, encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, list):
        sys.exit("输入 JSON 不是行记录数组，请确认是 guancli ds preview -f json 的输出")
    base = date.fromisoformat(base_date_str)

    n = len(data)
    contracts = {str(r.get("合同号") or "").strip() for r in data if r.get("合同号")}
    area = sum(num(r.get("计租面积")) or 0 for r in data)
    receivable = sum(num(r.get("应收金额（含税）")) or 0 for r in data)
    received = sum(num(r.get("已收金额（含税）")) or 0 for r in data)
    owed = sum(num(r.get("欠缴金额（含税）")) or 0 for r in data)
    dep = sum(num(r.get("保证金余额")) or 0 for r in data)

    wb = Workbook()
    ws = wb.active
    ws.title = "汇总"
    ws.sheet_view.showGridLines = False

    ws["B2"] = f"{store} — 在执行合同清单（基准日 {base_date_str}）"
    ws["B2"].font = Font(name="PingFang SC", size=14, bold=True)
    ws["B3"] = (
        "口径：起租日≤基准日≤到期日（日期窗口判定，合同状态仅展示不筛选）；"
        f"铺位类型：{pos_note}；数据来源：ADS-租费分析-租赁合同台账明细"
        + (f"（快照更新 {snapshot_time}）" if snapshot_time else "")
    )
    ws["B3"].font = Font(name="PingFang SC", size=9, color="808080")

    metrics = [
        ("在执行合同数（份）", n, "#,##0"),
        ("唯一合同号数", len(contracts), "#,##0"),
        ("计租面积合计（㎡）", area, "#,##0.00"),
        ("应收金额（含税）合计（元）", receivable, "#,##0.00"),
        ("已收金额（含税）合计（元）", received, "#,##0.00"),
        ("欠缴金额（含税）合计（元）", owed, "#,##0.00"),
        ("保证金余额合计（元）", dep, "#,##0.00"),
    ]
    r0 = 5
    for c, t in ((2, "指标"), (3, "数值")):
        cell = ws.cell(r0, c, t)
        cell.font = H_FONT; cell.fill = H_FILL; cell.border = BORDER
    for i, (k, v, fmt) in enumerate(metrics):
        ws.cell(r0 + 1 + i, 2, k).font = B_FONT
        cell = ws.cell(r0 + 1 + i, 3, v)
        cell.font = B_FONT; cell.number_format = fmt
        for c in (2, 3):
            ws.cell(r0 + 1 + i, c).border = BORDER

    def dist_table(r_start, title, counter, extra_area=None):
        ws.cell(r_start, 2, title).font = Font(name="PingFang SC", size=12, bold=True)
        heads = ["分类", "合同数"] + (["计租面积(㎡)"] if extra_area else [])
        for j, h in enumerate(heads):
            cell = ws.cell(r_start + 1, 2 + j, h)
            cell.font = H_FONT; cell.fill = H_FILL; cell.border = BORDER
        for i, (k, v) in enumerate(counter.most_common()):
            ws.cell(r_start + 2 + i, 2, k).font = B_FONT
            cell = ws.cell(r_start + 2 + i, 3, v)
            cell.font = B_FONT; cell.number_format = "#,##0"
            if extra_area:
                cell2 = ws.cell(r_start + 2 + i, 4, extra_area.get(k, 0))
                cell2.font = B_FONT; cell2.number_format = "#,##0.00"
            for j in range(len(heads)):
                ws.cell(r_start + 2 + i, 2 + j).border = BORDER
        return r_start + 2 + len(counter) + 2

    r = r0 + len(metrics) + 3
    pt = Counter(r.get("铺位类型") or "未分类" for r in data)
    pt_area = {
        k: sum(num(x.get("计租面积")) or 0 for x in data if (x.get("铺位类型") or "未分类") == k)
        for k in pt
    }
    r = dist_table(r, "按铺位类型分布", pt, pt_area)

    st = Counter(r.get("合同状态") or "未分类" for r in data)
    r = dist_table(r, "按合同状态分布（口径展示，不参与筛选）", st)

    biz = Counter(r.get("业态") or "未分类" for r in data)
    biz_area = {
        k: sum(num(x.get("计租面积")) or 0 for x in data if (x.get("业态") or "未分类") == k)
        for k in biz
    }
    r = dist_table(r, "按业态分布", biz, biz_area)

    exp = Counter(expire_bucket(x, base) for x in data)
    exp_order = ["基准日后30天内到期", "基准日后31-90天到期", "基准日后91-365天到期",
                 "一年以上后到期", "已早于基准日", "无到期日"]
    exp_sorted = [(k, exp[k]) for k in exp_order if k in exp] + \
                 [(k, v) for k, v in exp.most_common() if k not in exp_order]
    r = dist_table(r, "按到期时间分布（相对基准日）", Counter(dict(exp_sorted)))

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 16

    ws2 = wb.create_sheet("合同明细")
    ws2.freeze_panes = "A2"
    for j, col in enumerate(DETAIL_COLS, 1):
        cell = ws2.cell(1, j, col)
        cell.font = H_FONT; cell.fill = H_FILL; cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for i, row in enumerate(data):
        for j, col in enumerate(DETAIL_COLS, 1):
            v = clean(row.get(col), col)
            cell = ws2.cell(i + 2, j, v)
            cell.font = B_FONT; cell.border = BORDER
            if col in NUM_COLS and v is not None:
                cell.number_format = "#,##0.00"
            elif col == "合同号":
                cell.number_format = "@"
    for j, col in enumerate(DETAIL_COLS, 1):
        ws2.column_dimensions[get_column_letter(j)].width = WIDTHS.get(col, 12)
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(DETAIL_COLS))}{n + 1}"

    wb.save(output)
    print(f"已保存: {output}")
    print(f"汇总: 合同{n}份(唯一合同号{len(contracts)}), 面积{area:,.2f}㎡, "
          f"应收{receivable:,.2f}, 已收{received:,.2f}, 欠缴{owed:,.2f}, 保证金余额{dep:,.2f}")
    print("铺位类型分布:", dict(pt))
    print("合同状态分布:", dict(st))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="guancli ds preview -f json 输出文件")
    ap.add_argument("--base-date", required=True, help="基准日 YYYY-MM-DD")
    ap.add_argument("--store", required=True, help="门店全称（用于标题）")
    ap.add_argument("--output", required=True, help="输出 xlsx 路径")
    ap.add_argument("--snapshot-time", default="", help="数据集快照更新时间（可选）")
    ap.add_argument("--pos-note", default="正铺+多经", help="铺位类型范围说明")
    args = ap.parse_args()
    try:
        date.fromisoformat(args.base_date)
    except ValueError:
        sys.exit(f"基准日无效: {args.base_date}，需要 YYYY-MM-DD 格式的真实日历日期")
    build(args.input, args.base_date, args.store, args.output, args.snapshot_time, args.pos_note)


if __name__ == "__main__":
    main()
