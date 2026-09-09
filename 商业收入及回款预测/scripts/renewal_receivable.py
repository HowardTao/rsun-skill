#!/usr/bin/env python3
"""续签合同应收明细临时表生成器（预测年度内到期合同假定正常续签，中间临时表）。

用法:
    python renewal_receivable.py \
        --merged <工作目录>/output/执行+未执行合同清单_<基准日>.xlsx \
        --forecast-year <预测年度YYYY> --base-date <基准日YYYY-MM-DD> --store <门店全称> \
        --output <工作目录>/output/<门店简称>续签合同应收明细_<预测年度>年_<基准日>.xlsx \
        [--snapshot-time "..."]

口径（2026-09-08 用户确认）:
- 筛选: 合并清单中 **到期日落在预测年度内** 且 **是否续签=否** 的合同（在执行+未执行均含；
  是否续签=是 的合同已有真实续签合同，不生成假定续签明细）。
- 续签期滚动: 从原合同到期日+1 起，按合同支付周期连续切账期（账期长度=支付周期天数，
  不截断跨年）；**应收日 = 费用开始日 - 提前收款天数**，应收日落在预测年度内的账期才生成；
  应收日早于预测年度的账期跳过（回款属上一年度）但滚动继续；应收日晚于预测年度即停止。
- 费项: 租金/物业费/营销推广费三类，各用清单中的日单价/日净单价；日单价缺失的费项跳过。
- 金额（含税）: 应收金额 = 日单价×计租面积×账期天数；优惠后应收金额 = 日净单价×计租面积×账期天数；
  优惠金额 = 应收金额-优惠后应收金额；已收金额 = 0；未收金额 = 优惠后应收金额（预测行视为全部未回款）。
- 税率（2026-09-08 调整为固定费项税率，不再取历史账单税率）: **租金=9%、物业费=6%、营销推广费=6%**。
- 不含税: 不含税应收金额 = 应收金额/(1+税率)、不含税优惠后应收金额 = 优惠后应收金额/(1+税率)。
- 账单标识: 应收编号留空；账单主键 = 虚拟 ID `XC-<合同号>-<4位序号>`（与真实账单区分）；
  费用名 = 费用类型同名（租金/物业费/营销推广费）。
- 跳过并汇报: 支付周期或提前收款天数缺失的合同、三费项日单价全缺的合同、到期日缺失的合同。

输出: 单 sheet「续签合同应收明细」纯明细临时表，31 列与合同全周期应收明细格式完全一致
（含 12 个合同属性列：区域/楼栋/品牌/业态/主品类/楼层/计租面积/起租日/到期日/铺位类型/
合同状态/是否续签，按合同号取自合并清单；起租日/到期日/合同状态为原合同值），
按 合同号/费用类型/费用开始日/账单主键 排序；首行冻结 + 自动筛选；非最终输出表。
"""
import argparse
import os
import sys
from collections import defaultdict
from datetime import date, timedelta

from openpyxl import Workbook, load_workbook

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from make_report import num  # noqa: E402
from receivable_detail import ATTR_COLS, DETAIL_COLS, render_sheet  # noqa: E402

SHEET_NAME = "续签合同应收明细"
# 固定费项税率（2026-09-08 用户指定）: 租金 9%、物业费 6%、营销推广费 6%
FEE_TYPES = [
    # (费用类型/费用名, 日单价列, 日净单价列, 固定税率)
    ("租金", "租金日单价", "租金日净单价", 0.09),
    ("物业费", "物业费日单价", "物业费日净单价", 0.06),
    ("营销推广费", "营销推广费日单价", "营销推广费日净单价", 0.06),
]


def parse_date(v):
    s = str(v or "")[:10]
    if not s:
        return None
    try:
        return date.fromisoformat(s)
    except ValueError:
        return None


def load_merged(path):
    wb = load_workbook(path, read_only=True)
    ws = wb.active
    hdr = [c.value for c in next(ws.iter_rows(max_row=1))]
    rows = [dict(zip(hdr, r)) for r in ws.iter_rows(min_row=2, values_only=True)]
    wb.close()
    return rows


def build(merged_path, forecast_year, base_date_str, store, output, snapshot_time):
    try:
        year = int(forecast_year)
        assert 1900 < year < 2200
    except (ValueError, AssertionError):
        sys.exit(f"预测年度无效: {forecast_year}，需要 YYYY 格式（如 2026）")
    try:
        date.fromisoformat(base_date_str)
    except ValueError:
        sys.exit(f"基准日无效: {base_date_str}，需要 YYYY-MM-DD 格式的真实日历日期")

    merged = load_merged(merged_path)
    year_start, year_end = date(year, 1, 1), date(year, 12, 31)

    # 候选: 到期日在预测年度内 且 是否续签=否
    skipped_no_date = [r for r in merged
                       if r.get("是否续签") == "否" and parse_date(r.get("到期日")) is None
                       and str(r.get("起租日") or "")[:4] not in ("", str(year))]
    cands = [r for r in merged
             if r.get("是否续签") == "否"
             and (parse_date(r.get("到期日")) or date.min).year == year]
    renewed_excluded = [r for r in merged
                        if (parse_date(r.get("到期日")) or date.min).year == year
                        and r.get("是否续签") == "是"]

    rows = []
    skipped_period = []   # 支付周期/提前收款缺失
    skipped_price = []    # 三费项日单价全缺
    seq_by_cid = defaultdict(int)

    for r in cands:
        cid = str(r.get("合同号") or "").strip()
        end_old = parse_date(r.get("到期日"))
        period = r.get("支付周期")
        advance = r.get("提前收款天数")
        area = num(r.get("计租面积"))
        if period is None or advance is None or int(period) <= 0 or end_old is None:
            skipped_period.append((cid, str(r.get("到期日") or "")[:10]))
            continue
        period, advance = int(period), int(advance)
        fees = [(ft, num(r.get(p)), num(r.get(np_)), tax_rate)
                for ft, p, np_, tax_rate in FEE_TYPES if num(r.get(p)) is not None]
        if not fees or area is None or area <= 0:
            skipped_price.append(cid)
            continue

        s = end_old + timedelta(days=1)
        while True:
            e = s + timedelta(days=period - 1)
            recv = s - timedelta(days=advance)
            if recv > year_end:
                break
            if recv >= year_start:
                days = (e - s).days + 1
                attrs = {col: r.get(col) for col in ATTR_COLS}
                for ftype, price_day, net_day, tax_rate in fees:
                    seq_by_cid[cid] += 1
                    gross = round(price_day * area * days, 2)
                    net = round(net_day * area * days, 2) if net_day is not None else gross
                    rows.append({
                        "门店名称": r.get("门店名称"), "合同号": cid,
                        "铺位号": r.get("铺位号"), "租户名称": r.get("租户名称"),
                        **attrs,
                        "应收编号": None,
                        "账单主键": f"XC-{cid}-{seq_by_cid[cid]:04d}",
                        "费用名": ftype, "费用类型": ftype, "税率": tax_rate,
                        "费用开始日": s.isoformat(), "费用截止日": e.isoformat(),
                        "应收日": recv.isoformat(),
                        "应收金额": gross, "优惠金额": round(gross - net, 2),
                        "优惠后应收金额": net, "已收金额": 0.0, "未收金额": net,
                        "不含税应收金额": round(gross / (1 + tax_rate), 2)
                                        if tax_rate is not None else None,
                        "不含税优惠后应收金额": round(net / (1 + tax_rate), 2)
                                              if tax_rate is not None else None,
                    })
            s = e + timedelta(days=1)

    rows.sort(key=lambda x: (str(x.get("合同号") or ""), str(x.get("费用类型") or ""),
                             str(x.get("费用开始日") or ""), str(x.get("账单主键") or "")))

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    render_sheet(ws, rows, DETAIL_COLS)
    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    wb.save(output)

    # --- 校验与摘要 ---
    n_contracts = len({str(x.get("合同号")) for x in rows})
    print(f"已保存: {output}")
    print(f"门店: {store}, 基准日: {base_date_str}, 预测年度: {year}"
          + (f", 快照更新: {snapshot_time}" if snapshot_time else ""))
    print(f"候选合同: {len(cands)}份（预测年度内到期且是否续签=否；"
          f"在执行{sum(1 for r in cands if str(r.get('起租日') or '')[:10] <= base_date_str)}"
          f" + 未执行{sum(1 for r in cands if str(r.get('起租日') or '')[:10] > base_date_str)}）；"
          f"已续签不生成: {len(renewed_excluded)}份")
    print(f"生成: {len(rows)}行, 覆盖{n_contracts}份合同, "
          f"应收日范围 {min((x['应收日'] for x in rows), default='-')} ~ "
          f"{max((x['应收日'] for x in rows), default='-')}")
    if skipped_period:
        print(f"跳过（支付周期/提前收款天数缺失或到期日缺失）{len(skipped_period)}份: "
              + ", ".join(f"{c}({d})" for c, d in skipped_period[:10])
              + ("..." if len(skipped_period) > 10 else ""))
    if skipped_price:
        print(f"跳过（三费项日单价/计租面积缺失）{len(skipped_price)}份: "
              + ", ".join(skipped_price[:10]) + ("..." if len(skipped_price) > 10 else ""))
    if skipped_no_date:
        print(f"注意: {len(skipped_no_date)}份 是否续签=否 但到期日缺失的合同未参与判定")
    ft = {}
    for x in rows:
        ft[x["费用类型"]] = ft.get(x["费用类型"], 0) + 1
    print("费用类型行数分布:", ft)
    sums = {c: sum(num(x.get(c)) or 0 for x in rows) for c in
            ("应收金额", "优惠金额", "优惠后应收金额", "不含税应收金额",
             "不含税优惠后应收金额", "已收金额", "未收金额")}
    print("金额合计(元): " + ", ".join(f"{k}={v:,.2f}" for k, v in sums.items()))
    print("税率规则: 租金9%、物业费6%、营销推广费6%（固定费项税率）")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--merged", required=True, help="34 列合并清单 xlsx（output/ 下）")
    ap.add_argument("--forecast-year", required=True, help="预测年度 YYYY（必填，D14）")
    ap.add_argument("--base-date", required=True, help="基准日 YYYY-MM-DD")
    ap.add_argument("--store", required=True, help="门店全称（用于摘要打印）")
    ap.add_argument("--output", required=True, help="输出 xlsx 路径")
    ap.add_argument("--snapshot-time", default="", help="数据集快照更新时间（可选）")
    args = ap.parse_args()
    build(args.merged, args.forecast_year, args.base_date,
          args.store, args.output, args.snapshot_time)


if __name__ == "__main__":
    main()
