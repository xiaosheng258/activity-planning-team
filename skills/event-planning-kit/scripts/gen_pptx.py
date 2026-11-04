#!/usr/bin/env python3
"""宣传 PPTX 生成器（3-5 页极简推广版）

用法：
    python3 gen_pptx.py <data.json> [输出路径]

data.json 结构：
{
  "activity": "星海初见·迎新破冰",
  "slogan": "先认识人，再谈项目。",
  "date": "11月13日 周五 19:00",
  "venue": "大学生活动中心多功能厅",
  "target": "2026 级全体新成员",
  "org": "星海创新实践工作室",
  "highlights": ["两小时，8 个小组", "你画我猜技术梗版", "来了送纪念钥匙扣"],
  "cta": "扫码报名，11月10日 24:00 截止",
  "color": "1E5B96"
}
"""
import json
import sys
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


def hex2rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_text(slide, text, left, top, width, height, size, bold=False,
             color=RGBColor(0x2C, 0x3E, 0x50), align=PP_ALIGN.LEFT, font="微软雅黑"):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    run = p.runs[0]
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    return tb


def build(data, out_path):
    accent = hex2rgb(data.get("color", "1E5B96"))
    white = RGBColor(0xFF, 0xFF, 0xFF)
    gray = RGBColor(0x6B, 0x7C, 0x93)
    dark = RGBColor(0x2C, 0x3E, 0x50)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    W, H = 13.333, 7.5

    # ---------- 第1页 封面 ----------
    s = prs.slides.add_slide(blank)
    bg = s.shapes.add_shape(1, 0, 0, Inches(W), Inches(H))
    bg.fill.solid()
    bg.fill.fore_color.rgb = accent
    bg.line.fill.background()

    add_text(s, data.get("activity", "活动"), 1.2, 2.0, W - 2.4, 1.2, 44, True, white)
    add_text(s, data.get("slogan", ""), 1.2, 3.25, W - 2.4, 0.8, 22, False, RGBColor(0xC9, 0xDE, 0xF2))
    add_text(s, f"{data.get('date','')}　|　{data.get('venue','')}", 1.2, 4.5, W - 2.4, 0.6, 16, False, white)
    add_text(s, data.get("org", ""), 1.2, 6.4, W - 2.4, 0.5, 13, False, RGBColor(0xC9, 0xDE, 0xF2))

    # ---------- 第2页 亮点 ----------
    s = prs.slides.add_slide(blank)
    add_text(s, "你可以期待什么", 1.0, 0.7, 8, 0.8, 32, True, accent)
    hl = data.get("highlights", [])[:4]
    y = 2.1
    for i, item in enumerate(hl, 1):
        circle = s.shapes.add_shape(9, Inches(1.0), Inches(y), Inches(0.55), Inches(0.55))
        circle.fill.solid()
        circle.fill.fore_color.rgb = accent
        circle.line.fill.background()
        tf = circle.text_frame
        tf.text = str(i)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.runs[0].font.size = Pt(16)
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = white
        add_text(s, item, 1.85, y - 0.08, W - 3.2, 0.72, 20, False, dark)
        y += 1.15

    # ---------- 第3页 信息 ----------
    s = prs.slides.add_slide(blank)
    bar = s.shapes.add_shape(1, 0, 0, Inches(0.35), Inches(H))
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent
    bar.line.fill.background()

    add_text(s, "活动信息", 1.0, 0.7, 8, 0.8, 32, True, accent)
    rows = [("时间", data.get("date", "—")),
            ("地点", data.get("venue", "—")),
            ("对象", data.get("target", "—")),
            ("主办", data.get("org", "—"))]
    y = 2.0
    for k, v in rows:
        add_text(s, k, 1.0, y, 1.5, 0.6, 18, True, gray)
        add_text(s, v, 2.6, y, W - 4.0, 0.6, 18, False, dark)
        y += 1.0

    # ---------- 第4页 行动 ----------
    s = prs.slides.add_slide(blank)
    bg = s.shapes.add_shape(1, 0, 0, Inches(W), Inches(H))
    bg.fill.solid()
    bg.fill.fore_color.rgb = accent
    bg.line.fill.background()

    add_text(s, data.get("cta", "扫码报名"), 1.2, 2.6, W - 2.4, 1.0, 34, True, white, PP_ALIGN.CENTER)
    add_text(s, "名额有限，先到先得", 1.2, 3.9, W - 2.4, 0.6, 18, False,
             RGBColor(0xC9, 0xDE, 0xF2), PP_ALIGN.CENTER)
    add_text(s, data.get("org", ""), 1.2, 6.4, W - 2.4, 0.5, 13, False,
             RGBColor(0xC9, 0xDE, 0xF2), PP_ALIGN.CENTER)

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    prs.save(out_path)
    return out_path, len(prs.slides)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    with open(sys.argv[1], encoding="utf-8") as f:
        data = json.load(f)
    name = data.get("activity", "活动").replace("/", "_")
    out = sys.argv[2] if len(sys.argv) > 2 else f"宣传_{name}.pptx"
    path, n = build(data, out)
    print(f"✅ 已生成：{path}（{n} 页）")


if __name__ == "__main__":
    main()
