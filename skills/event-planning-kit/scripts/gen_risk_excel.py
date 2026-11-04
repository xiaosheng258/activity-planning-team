#!/usr/bin/env python3
"""风险预案 Excel 生成器

用法：
    python3 gen_risk_excel.py <data.json> [输出路径]

data.json 结构：
{
  "activity": "星海初见·迎新破冰",
  "type": "内部建设",
  "date": "2026-11-13",
  "scale": "60人",
  "risks": [
    {"item": "骨干临时请假", "type": "人员", "level": "高",
     "prevent": "每环节指定替补", "handle": "替补顶上", "owner": "副社长", "backup": "合并环节"}
  ],
  "contacts": [
    {"role": "活动总负责人", "owner": "张三", "phone": "由本人提供", "when": "立项时"}
  ],
  "evacuation": "主出口：活动室正门　备用出口：侧门　通道保留 1.5 米"
}

输出：两 Sheet
  Sheet1 风险预案（按等级排序）
  Sheet2 应急联络
"""
import json
import sys
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

MAIN = "1E5B96"
LIGHT = "E8F0F8"
BORDER_C = "D0DBE8"

LEVEL_STYLE = {
    "高": ("C0392B", "FDECEA"),
    "中": ("B9770E", "FDF6E8"),
    "低": ("1E8449", "EEF7F1"),
}

thin = Side(style="thin", color=BORDER_C)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD_FONT = Font(bold=True, color=MAIN, size=10)
HEAD_FILL = PatternFill("solid", fgColor=LIGHT)


def build(data, out_path):
    wb = Workbook()

    # ---------- Sheet1 风险预案 ----------
    ws = wb.active
    ws.title = "风险预案"

    ws["A1"] = f"{data.get('activity', '活动')} · 风险预案"
    ws["A1"].font = Font(bold=True, size=14, color=MAIN)
    ws.merge_cells("A1:G1")
    ws.row_dimensions[1].height = 28

    meta = (f"活动类型：{data.get('type', '—')}　规模：{data.get('scale', '—')}　"
            f"日期：{data.get('date', '—')}　编制：{datetime.now():%Y-%m-%d}")
    ws["A2"] = meta
    ws["A2"].font = Font(size=9, color="6B7C93")
    ws.merge_cells("A2:G2")

    headers = ["风险项", "类型", "等级", "预防措施", "事发处置", "责任人", "备用方案"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 22

    # 按等级排序：高 > 中 > 低
    order = {"高": 0, "中": 1, "低": 2}
    risks = sorted(data.get("risks", []), key=lambda r: order.get(r.get("level", "低"), 3))

    row = 4
    stats = {"高": 0, "中": 0, "低": 0}
    for r in risks:
        lv = r.get("level", "低")
        stats[lv] = stats.get(lv, 0) + 1
        color, bg = LEVEL_STYLE.get(lv, ("333333", "FFFFFF"))
        vals = [r.get("item", ""), r.get("type", ""), lv,
                r.get("prevent", ""), r.get("handle", ""),
                r.get("owner", ""), r.get("backup", "—")]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=row, column=c, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if c == 3:
                cell.font = Font(bold=True, color=color)
                cell.fill = PatternFill("solid", fgColor=bg)
                cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 34
        row += 1

    # 统计
    row += 1
    ws.cell(row=row, column=1, value="风险统计").font = Font(bold=True, color=MAIN)
    row += 1
    summary = f"高风险 {stats.get('高', 0)} 项 ｜ 中风险 {stats.get('中', 0)} 项 ｜ 低风险 {stats.get('低', 0)} 项 ｜ 合计 {len(risks)} 项"
    ws.cell(row=row, column=1, value=summary).font = Font(size=10)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)

    if stats.get("高", 0) == 0:
        row += 1
        ws.cell(row=row, column=1, value="⚠️ 未识别出高风险项，请复核是否回避了真问题（人员不足、经费未批、未报备）。")
        ws.cell(row=row, column=1).font = Font(size=9, color="C0392B")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)

    ev = data.get("evacuation")
    if ev:
        row += 2
        ws.cell(row=row, column=1, value="撤离路线").font = Font(bold=True, color=MAIN)
        row += 1
        ws.cell(row=row, column=1, value=ev).font = Font(size=9)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        ws.cell(row=row, column=1).alignment = Alignment(wrap_text=True, vertical="top")

    widths = [20, 10, 7, 26, 26, 12, 20]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"

    # ---------- Sheet2 应急联络 ----------
    ws2 = wb.create_sheet("应急联络")
    ws2["A1"] = "应急联络表"
    ws2["A1"].font = Font(bold=True, size=13, color=MAIN)
    ws2.merge_cells("A1:D1")

    ws2["A2"] = "社团无固定通讯录，本表需在各时点前逐一确认到人"
    ws2["A2"].font = Font(size=9, color="6B7C93")
    ws2.merge_cells("A2:D2")

    heads2 = ["角色", "责任人", "联系方式", "确认时机"]
    for c, h in enumerate(heads2, 1):
        cell = ws2.cell(row=4, column=c, value=h)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center")

    r = 5
    for ct in data.get("contacts", []):
        vals = [ct.get("role", ""), ct.get("owner", ""), ct.get("phone", ""), ct.get("when", "")]
        for c, v in enumerate(vals, 1):
            cell = ws2.cell(row=r, column=c, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws2.row_dimensions[r].height = 24
        r += 1

    for i, w in enumerate([20, 16, 22, 18], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    wb.save(out_path)
    return out_path, len(risks), stats


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    with open(sys.argv[1], encoding="utf-8") as f:
        data = json.load(f)
    name = data.get("activity", "活动").replace("/", "_")
    out = sys.argv[2] if len(sys.argv) > 2 else f"风险预案_{name}.xlsx"
    path, cnt, stats = build(data, out)
    print(f"✅ 已生成：{path}")
    print(f"   共 {cnt} 项风险：高 {stats.get('高',0)} / 中 {stats.get('中',0)} / 低 {stats.get('低',0)}")


if __name__ == "__main__":
    main()
