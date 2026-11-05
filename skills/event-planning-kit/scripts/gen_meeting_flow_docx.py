#!/usr/bin/env python3
"""会议/活动流程单生成器（Word）

格式参考：11月5日会议主要流程.docx
  - A4 纵向，页边距 上下 2.54cm / 左右 3.17cm
  - 三级嵌套编号：L0 = 1.  L1 = (1)  L2 = ①
  - 正文小四（12pt），无额外段间距

用法：
    python3 gen_meeting_flow_docx.py <data.json> [输出路径]

data.json 结构：
{
  "title": "11月5日会议流程",
  "nodes": [
    {"lvl": 0, "text": "由胡一鸣进行人员清点。"},
    {"lvl": 0, "text": "清点后首先强调会议整体内容，1-2分钟简单介绍即可。"},
    {"lvl": 1, "text": "介绍工作室简介（刘硕）"},
    {"lvl": 1, "text": "宣读员工手册"},
    {"lvl": 0, "text": "部门划分："},
    {"lvl": 2, "text": "项目部：部长王淼鑫，代理部长王文豪，成员：刘硕、胡一鸣"},
    {"lvl": 0, "text": "散会"}
  ]
}

lvl: 0 = 一级（1.），1 = 二级（(1)），2 = 三级（①）
"""
import json
import sys
import os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# 三级编号：1. / (1) / ①
NUM_FORMATS = [
    ("decimal", "%1."),                            # 1.
    ("decimal", "(%2)"),                           # (1)
    ("decimalEnclosedCircleChinese", "%3"),        # ①
]
# 各级缩进（twips，1cm ≈ 567twips）：left, hanging
LEVEL_INDENT = [
    (840, 420),   # L0: 左缩进 1.48cm，悬挂 0.74cm
    (1260, 420),  # L1: 递进
    (1680, 420),  # L2: 再递进
]


def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(f"w:{k}"), str(v))
    return e


def add_numbering(doc):
    """在 numbering.xml 中注册一套三级编号，返回 numId"""
    numbering = doc.part.numbering_part.element

    # 找当前最大 abstractNumId / numId
    abs_ids = [int(a.get(qn("w:abstractNumId"))) for a in numbering.findall(qn("w:abstractNum"))]
    num_ids = [int(n.get(qn("w:numId"))) for n in numbering.findall(qn("w:num"))]
    new_abs_id = (max(abs_ids) + 1) if abs_ids else 0
    new_num_id = (max(num_ids) + 1) if num_ids else 1

    abstract = _el("w:abstractNum", abstractNumId=new_abs_id)
    # 多级列表标记
    abstract.append(_el("w:multiLevelType", val="multilevel"))

    for lvl, (fmt, text) in enumerate(NUM_FORMATS):
        left, hanging = LEVEL_INDENT[lvl]
        lvl_el = _el("w:lvl", ilvl=lvl)
        lvl_el.append(_el("w:start", val=1))
        lvl_el.append(_el("w:numFmt", val=fmt))
        lvl_el.append(_el("w:lvlText", val=text))
        lvl_el.append(_el("w:lvlJc", val="left"))
        ppr = _el("w:pPr")
        ppr.append(_el("w:ind", left=left, hanging=hanging))
        lvl_el.append(ppr)
        abstract.append(lvl_el)

    # abstractNum 必须排在 num 之前
    first_num = numbering.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abstract)
    else:
        numbering.append(abstract)

    num = _el("w:num", numId=new_num_id)
    num.append(_el("w:abstractNumId", val=new_abs_id))
    numbering.append(num)
    return new_num_id


def set_run_font(run, size_pt=12):
    run.font.size = Pt(size_pt)
    run.font.name = "宋体"
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = _el("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), "宋体")
    rfonts.set(qn("w:ascii"), "Times New Roman")


def build(data, out_path):
    doc = Document()

    # 页面：A4 纵向，页边距参考原文件
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.54)
    sec.bottom_margin = Cm(2.54)
    sec.left_margin = Cm(3.17)
    sec.right_margin = Cm(3.17)

    # 正文默认样式
    normal = doc.styles["Normal"]
    normal.font.name = "宋体"
    normal.font.size = Pt(12)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    pf = normal.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.5

    num_id = add_numbering(doc)

    # 标题（可选）
    title = data.get("title")
    if title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        set_run_font(r, 16)
        r.font.bold = True
        p.paragraph_format.space_after = Pt(12)
        p.paragraph_format.line_spacing = 1.5

    # 编号段落
    for node in data.get("nodes", []):
        lvl = int(node.get("lvl", 0))
        lvl = max(0, min(lvl, len(NUM_FORMATS) - 1))
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.5

        ppr = p._p.get_or_add_pPr()
        numpr = _el("w:numPr")
        numpr.append(_el("w:ilvl", val=lvl))
        numpr.append(_el("w:numId", val=num_id))
        ppr.append(numpr)

        r = p.add_run(str(node.get("text", "")))
        set_run_font(r, 12)

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    doc.save(out_path)
    return out_path, len(data.get("nodes", []))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    with open(sys.argv[1], encoding="utf-8") as f:
        data = json.load(f)
    title = data.get("title", "流程单").replace("/", "_")
    out = sys.argv[2] if len(sys.argv) > 2 else f"{title}.docx"
    path, n = build(data, out)
    print(f"✅ 已生成：{path}（{n} 个流程节点）")


if __name__ == "__main__":
    main()
