#!/usr/bin/env python3
"""预算表 Excel 生成器

用法：
    python3 gen_budget_excel.py <data.json> [输出路径]

data.json 结构：
{
  "activity": "星海初见·迎新破冰",
  "type": "内部建设",
  "date": "2026-11-13",
  "items": [
    {"name": "姓名贴", "qty": 75, "price": 0.5, "note": "60人×1.2含备用"},
    {"name": "矿泉水", "qty": 72, "price": 1.2, "note": "按人数×1.2"}
  ],
  "reserve_rate": 0.1
}

输出：两 Sheet
  Sheet1 预算明细（含合计行）
  Sheet2 汇总（各类占比）
"""
import json
import sys
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 样式常量
MAIN = "1E5B96"
LIGHT = "E8F0F8"
BORDER_C = "D0DBE8"
SUM_BG = "FFF8E6"

thin = Side(style="thin", color=BORDER_C)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD_FONT = Font(bold=True, color=MAIN, size=11)
HEAD_FILL = PatternFill("solid", fgColor=LIGHT)
SUM_FILL = PatternFill("solid", fgColor=SUM_BG)


def classify(name):
    """按项目名归类，用于汇总 Sheet"""
    if any(k in name for k in ["场地", "租金"]):
        return "场地"
    if any(k in name for k in ["打印", "横幅", "海报", "展架", "名片", "姓名贴",
                               "签到", "证书", "奖状", "卡片", "单页", "贴纸", "喷绘"]):
        return "印刷物料"
    if any(k in name for k in ["水", "茶", "零食", "餐", "蛋糕", "饮料", "奶茶", "小食"]):
        return "食品饮品"
    if any(k in name for k in ["奖", "礼品", "纪念", "钥匙扣", "书签", "文化衫", "徽章"]):
        return "奖品纪念品"
    if any(k in name for k in ["音响", "话筒", "投影", "灯光", "设备", "租赁", "摄影"]):
        return "设备"
    return "其他"


def build(data, out_path):
    wb = Workbook()

    # ---------- Sheet1 明细 ----------
    ws = wb.active
    ws.title = "预算明细"

    ws["A1"] = f"{data.get('activity', '活动')} · 预算明细"
    ws["A1"].font = Font(bold=True, size=14, color=MAIN)
    ws.merge_cells("A1:E1")
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    meta = f"活动类型：{data.get('type', '—')}　日期：{data.get('date', '—')}　编制：{datetime.now():%Y-%m-%d}"
    ws["A2"] = meta
    ws["A2"].font = Font(size=9, color="6B7C93")
    ws.merge_cells("A2:E2")

    headers = ["支出项目", "数量", "单价(元)", "总价(元)", "说明"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 22

    row = 4
    total = 0.0
    for it in data.get("items", []):
        qty = it.get("qty", 0)
        price = it.get("price", 0)
        subtotal = round(qty * price, 2)
        total += subtotal
        ws.cell(row=row, column=1, value=it.get("name", "")).border = BORDER
        c2 = ws.cell(row=row, column=2, value=qty)
        c2.border = BORDER
        c2.alignment = Alignment(horizontal="right")
        c3 = ws.cell(row=row, column=3, value=price)
        c3.border = BORDER
        c3.alignment = Alignment(horizontal="right")
        c3.number_format = "0.00"
        c4 = ws.cell(row=row, column=4, value=f"=B{row}*C{row}")
        c4.border = BORDER
        c4.alignment = Alignment(horizontal="right")
        c4.number_format = "#,##0.00"
        ws.cell(row=row, column=5, value=it.get("note", "")).border = BORDER
        row += 1

    # 小计
    ws.cell(row=row, column=1, value="小计").font = Font(bold=True)
    sub_cell = ws.cell(row=row, column=4, value=f"=SUM(D4:D{row-1})")
    sub_cell.font = Font(bold=True)
    sub_cell.number_format = "#,##0.00"
    sub_cell.alignment = Alignment(horizontal="right")
    for c in range(1, 6):
        ws.cell(row=row, column=c).border = BORDER
    sub_row = row
    row += 1

    # 备用金
    rate = data.get("reserve_rate", 0.1)
    ws.cell(row=row, column=1, value="机动备用金")
    ws.cell(row=row, column=5, value=f"小计 × {int(rate*100)}%")
    res_cell = ws.cell(row=row, column=4, value=f"=ROUND(D{sub_row}*{rate},2)")
    res_cell.number_format = "#,##0.00"
    res_cell.alignment = Alignment(horizontal="right")
    for c in range(1, 6):
        ws.cell(row=row, column=c).border = BORDER
    res_row = row
    row += 1

    # 合计
    ws.cell(row=row, column=1, value="合计").font = Font(bold=True, size=12, color=MAIN)
    tot_cell = ws.cell(row=row, column=4, value=f"=D{sub_row}+D{res_row}")
    tot_cell.font = Font(bold=True, size=12, color=MAIN)
    tot_cell.number_format = "#,##0.00"
    tot_cell.alignment = Alignment(horizontal="right")
    ws.cell(row=row, column=5, value="含机动备用金").font = Font(size=9, color="6B7C93")
    for c in range(1, 6):
        ws.cell(row=row, column=c).border = BORDER
        ws.cell(row=row, column=c).fill = SUM_FILL
    total_row = row

    # 备注
    note = data.get("note")
    if note:
        row += 2
        ws.cell(row=row, column=1, value="说明").font = Font(bold=True, color=MAIN)
        ws.cell(row=row + 1, column=1, value=note).font = Font(size=9)
        ws.merge_cells(start_row=row + 1, start_column=1, end_row=row + 1, end_column=5)
        ws.cell(row=row + 1, column=1).alignment = Alignment(wrap_text=True, vertical="top")

    widths = [22, 10, 12, 14, 34]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"

    # ---------- Sheet2 汇总 ----------
    ws2 = wb.create_sheet("分类汇总")
    ws2["A1"] = "支出结构汇总"
    ws2["A1"].font = Font(bold=True, size=13, color=MAIN)
    ws2.merge_cells("A1:C1")

    heads2 = ["类别", "金额(元)", "占比"]
    for c, h in enumerate(heads2, 1):
        cell = ws2.cell(row=3, column=c, value=h)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center")

    groups = {}
    for it in data.get("items", []):
        g = classify(it.get("name", ""))
        groups[g] = groups.get(g, 0) + it.get("qty", 0) * it.get("price", 0)

    r = 4
    for g, amt in sorted(groups.items(), key=lambda x: -x[1]):
        ws2.cell(row=r, column=1, value=g).border = BORDER
        c2 = ws2.cell(row=r, column=2, value=round(amt, 2))
        c2.border = BORDER
        c2.number_format = "#,##0.00"
        c2.alignment = Alignment(horizontal="right")
        pct = amt / total if total else 0
        c3 = ws2.cell(row=r, column=3, value=pct)
        c3.border = BORDER
        c3.number_format = "0.0%"
        c3.alignment = Alignment(horizontal="right")
        r += 1

    ws2.cell(row=r, column=1, value="合计").font = Font(bold=True)
    c = ws2.cell(row=r, column=2, value=round(total, 2))
    c.font = Font(bold=True)
    c.number_format = "#,##0.00"
    c.alignment = Alignment(horizontal="right")
    ws2.cell(row=r, column=3, value=1.0).number_format = "0.0%"
    for cc in range(1, 4):
        ws2.cell(row=r, column=cc).border = BORDER
        ws2.cell(row=r, column=cc).fill = SUM_FILL

    for i, w in enumerate([18, 16, 12], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    wb.save(out_path)
    return out_path, total, total * rate


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    with open(sys.argv[1], encoding="utf-8") as f:
        data = json.load(f)
    name = data.get("activity", "活动").replace("/", "_")
    out = sys.argv[2] if len(sys.argv) > 2 else f"预算表_{name}.xlsx"
    path, total, reserve = build(data, out)
    print(f"✅ 已生成：{path}")
    print(f"   小计 {total:,.2f} 元 + 备用金 {reserve:,.2f} 元 = 合计 {total + reserve:,.2f} 元")


if __name__ == "__main__":
    main()
