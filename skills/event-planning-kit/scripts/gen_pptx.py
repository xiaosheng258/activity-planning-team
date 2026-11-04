#!/usr/bin/env python3
"""宣传 / 汇报 PPTX 生成器（双模式）

用法：
    python3 gen_pptx.py <data.json> [输出路径]
    python3 gen_pptx.py <data.json> [输出路径] --mode promo|report

模式：
    promo  （默认）招新宣传版，4 页：封面 / 亮点 / 信息 / 行动号召
    report 　审批汇报版，8 页：封面 / 背景 / 目标 / 流程 / 预算 / 风险 / 效果 / 申请

data.json 通用字段：
{
  "mode": "promo",                  # 可省略，命令行 --mode 优先
  "activity": "星海初见·迎新破冰",
  "type": "内部建设",
  "date": "11月13日 周五 19:00",
  "venue": "大学生活动中心多功能厅",
  "target": "2026 级全体新成员",
  "org": "星海创新实践工作室",
  "color": "1E5B96"
}

promo 额外字段：
  "slogan": "...", "highlights": ["...","..."], "cta": "扫码报名..."

report 额外字段：
  "background": "背景一段话",
  "objectives": ["目标1", "目标2"],
  "flow": [{"time":"18:30","name":"签到","owner":"张三"}],
  "budget_total": "792 元",
  "budget_items": [{"name":"零食","amount":"396"}],
  "risks": [{"item":"人数不足","level":"高","plan":"提前宣传"}],
  "outcomes": ["预计覆盖 60 人", "形成 8 个分组"],
  "request": "申请经费 792 元，场地一间"
"""
import json
import sys
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

W, H = 13.333, 7.5
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x6B, 0x7C, 0x93)
DARK = RGBColor(0x2C, 0x3E, 0x50)
LIGHT_BG = RGBColor(0xF2, 0xF6, 0xFA)


def hex2rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_text(slide, text, left, top, width, height, size, bold=False,
             color=DARK, align=PP_ALIGN.LEFT, font="微软雅黑"):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = str(text)
    p.alignment = align
    if p.runs:
        r = p.runs[0]
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        r.font.name = font
    return tb


def add_rect(slide, left, top, width, height, color, line=False):
    shp = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(width), Inches(height))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    if not line:
        shp.line.fill.background()
    return shp


def add_bullets(slide, items, left, top, width, size=16, color=DARK, gap=0.55, accent=None):
    y = top
    for it in items:
        dot_color = accent or DARK
        add_text(slide, "·", left, y, 0.3, 0.4, size + 4, True, dot_color)
        add_text(slide, it, left + 0.32, y, width - 0.32, 0.4, size, False, color)
        y += gap
    return y


# ---------------- promo 模式 ----------------
def build_promo(prs, data, accent):
    blank = prs.slide_layouts[6]

    # 封面
    s = prs.slides.add_slide(blank)
    add_rect(s, 0, 0, W, H, accent)
    add_text(s, data.get("activity", "活动"), 1.2, 2.0, W - 2.4, 1.2, 44, True, WHITE)
    add_text(s, data.get("slogan", ""), 1.2, 3.25, W - 2.4, 0.8, 22, False,
             RGBColor(0xC9, 0xDE, 0xF2))
    add_text(s, f"{data.get('date','')}　|　{data.get('venue','')}", 1.2, 4.5,
             W - 2.4, 0.6, 16, False, WHITE)
    add_text(s, data.get("org", ""), 1.2, 6.4, W - 2.4, 0.5, 13, False,
             RGBColor(0xC9, 0xDE, 0xF2))

    # 亮点
    s = prs.slides.add_slide(blank)
    add_text(s, "你可以期待什么", 1.0, 0.7, 8, 0.8, 32, True, accent)
    y = 2.1
    for i, item in enumerate(data.get("highlights", [])[:4], 1):
        c = s.shapes.add_shape(9, Inches(1.0), Inches(y), Inches(0.55), Inches(0.55))
        c.fill.solid()
        c.fill.fore_color.rgb = accent
        c.line.fill.background()
        c.text_frame.text = str(i)
        p = c.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.runs[0].font.size = Pt(16)
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = WHITE
        add_text(s, item, 1.85, y - 0.08, W - 3.2, 0.72, 20, False, DARK)
        y += 1.15

    # 信息
    s = prs.slides.add_slide(blank)
    add_rect(s, 0, 0, 0.35, H, accent)
    add_text(s, "活动信息", 1.0, 0.7, 8, 0.8, 32, True, accent)
    rows = [("时间", data.get("date", "—")), ("地点", data.get("venue", "—")),
            ("对象", data.get("target", "—")), ("主办", data.get("org", "—"))]
    y = 2.0
    for k, v in rows:
        add_text(s, k, 1.0, y, 1.5, 0.6, 18, True, GRAY)
        add_text(s, v, 2.6, y, W - 4.0, 0.6, 18, False, DARK)
        y += 1.0

    # 行动
    s = prs.slides.add_slide(blank)
    add_rect(s, 0, 0, W, H, accent)
    add_text(s, data.get("cta", "扫码报名"), 1.2, 2.6, W - 2.4, 1.0, 34, True,
             WHITE, PP_ALIGN.CENTER)
    add_text(s, "名额有限，先到先得", 1.2, 3.9, W - 2.4, 0.6, 18, False,
             RGBColor(0xC9, 0xDE, 0xF2), PP_ALIGN.CENTER)
    add_text(s, data.get("org", ""), 1.2, 6.4, W - 2.4, 0.5, 13, False,
             RGBColor(0xC9, 0xDE, 0xF2), PP_ALIGN.CENTER)


# ---------------- report 模式 ----------------
def section_head(slide, num, title, accent):
    add_rect(slide, 0, 0, 0.28, H, accent)
    add_text(slide, f"{num:02d}", 0.7, 0.55, 1.2, 0.9, 40, True,
             RGBColor(0xD6, 0xE4, 0xF2))
    add_text(slide, title, 0.7, 0.62, W - 2.0, 0.8, 30, True, accent)


def build_report(prs, data, accent):
    blank = prs.slide_layouts[6]

    # 1 封面
    s = prs.slides.add_slide(blank)
    add_rect(s, 0, 0, W, H, accent)
    add_text(s, "活动策划汇报", 1.2, 1.6, W - 2.4, 0.6, 18, False,
             RGBColor(0xC9, 0xDE, 0xF2))
    add_text(s, data.get("activity", "活动"), 1.2, 2.3, W - 2.4, 1.2, 42, True, WHITE)
    add_text(s, f"{data.get('date','')}　|　{data.get('venue','')}", 1.2, 3.7,
             W - 2.4, 0.6, 16, False, WHITE)
    add_text(s, f"主办：{data.get('org','')}　规模：{data.get('target','')}", 1.2, 4.4,
             W - 2.4, 0.6, 14, False, RGBColor(0xC9, 0xDE, 0xF2))

    # 2 背景与意义
    s = prs.slides.add_slide(blank)
    section_head(s, 1, "活动背景与意义", accent)
    bg = data.get("background", "")
    add_text(s, bg, 0.7, 1.9, W - 1.8, 1.4, 16, False, DARK)
    add_text(s, "目标", 0.7, 3.4, 3, 0.5, 18, True, accent)
    add_bullets(s, data.get("objectives", [])[:4], 0.7, 3.95, W - 1.8, 15, DARK, 0.62, accent)

    # 3 流程
    s = prs.slides.add_slide(blank)
    section_head(s, 2, "活动流程", accent)
    flow = data.get("flow", [])[:9]
    y = 1.75
    for f in flow:
        add_rect(s, 0.7, y, 1.35, 0.42, RGBColor(0xE8, 0xF0, 0xF8))
        add_text(s, f.get("time", ""), 0.72, y, 1.3, 0.42, 13, True, accent, PP_ALIGN.CENTER)
        add_text(s, f.get("name", ""), 2.25, y, 4.2, 0.42, 14, False, DARK)
        add_text(s, f.get("owner", ""), 6.6, y, 2.4, 0.42, 13, False, GRAY)
        y += 0.55

    # 4 预算
    s = prs.slides.add_slide(blank)
    section_head(s, 3, "经费预算", accent)
    add_text(s, "预算总额", 0.7, 1.85, 3, 0.5, 16, False, GRAY)
    add_text(s, data.get("budget_total", "—"), 0.7, 2.25, 5, 0.9, 40, True, accent)
    items = data.get("budget_items", [])[:6]
    y = 3.5
    for it in items:
        add_text(s, it.get("name", ""), 0.7, y, 3.6, 0.42, 15, False, DARK)
        add_text(s, it.get("amount", ""), 4.4, y, 1.6, 0.42, 15, False, DARK, PP_ALIGN.RIGHT)
        y += 0.5

    # 5 风险
    s = prs.slides.add_slide(blank)
    section_head(s, 4, "风险与应对", accent)
    lv_color = {"高": RGBColor(0xC0, 0x39, 0x2B), "中": RGBColor(0xB9, 0x77, 0x0E),
                "低": RGBColor(0x1E, 0x84, 0x49)}
    y = 1.85
    for r in data.get("risks", [])[:5]:
        lv = r.get("level", "中")
        add_rect(s, 0.7, y + 0.06, 0.5, 0.34, lv_color.get(lv, GRAY))
        add_text(s, lv, 0.7, y + 0.06, 0.5, 0.34, 12, True, WHITE, PP_ALIGN.CENTER)
        add_text(s, r.get("item", ""), 1.4, y, 3.4, 0.42, 15, True, DARK)
        add_text(s, r.get("plan", ""), 4.9, y, W - 5.8, 0.42, 13, False, GRAY)
        y += 0.78

    # 6 预期效果
    s = prs.slides.add_slide(blank)
    section_head(s, 5, "预期效果", accent)
    add_bullets(s, data.get("outcomes", [])[:5], 0.7, 2.0, W - 1.8, 17, DARK, 0.75, accent)

    # 7 申请事项
    s = prs.slides.add_slide(blank)
    add_rect(s, 0, 0, W, H, accent)
    add_text(s, "申请事项", 1.2, 1.6, W - 2.4, 0.7, 24, True, RGBColor(0xC9, 0xDE, 0xF2))
    add_text(s, data.get("request", ""), 1.2, 2.5, W - 2.4, 1.6, 26, True, WHITE)
    add_text(s, f"{data.get('org','')}　|　{data.get('date','')}", 1.2, 6.4,
             W - 2.4, 0.5, 13, False, RGBColor(0xC9, 0xDE, 0xF2))


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    mode_arg = None
    for i, a in enumerate(sys.argv):
        if a == "--mode" and i + 1 < len(sys.argv):
            mode_arg = sys.argv[i + 1]
    if not args:
        print(__doc__)
        sys.exit(1)

    with open(args[0], encoding="utf-8") as f:
        data = json.load(f)

    mode = mode_arg or data.get("mode", "promo")
    if mode not in ("promo", "report"):
        print(f"未知模式：{mode}（可选 promo / report）")
        sys.exit(1)

    accent = hex2rgb(data.get("color", "1E5B96"))
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)

    if mode == "promo":
        build_promo(prs, data, accent)
    else:
        build_report(prs, data, accent)

    name = data.get("activity", "活动").replace("/", "_")
    suffix = "宣传" if mode == "promo" else "汇报"
    out = args[1] if len(args) > 1 else f"{suffix}_{name}.pptx"
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    prs.save(out)
    print(f"✅ 已生成（{mode} 模式）：{out}（{len(prs.slides)} 页）")


if __name__ == "__main__":
    main()
