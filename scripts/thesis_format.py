#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
论文格式学习与应用工具
从参考文档学习格式规则，然后批量应用到其他文档。

用法:
  python thesis_format.py learn reference.docx          # 学习格式，输出 rules.json
  python thesis_format.py apply input.docx rules.json   # 应用格式，输出 _formatted.docx
  python thesis_format.py check input.docx rules.json   # 检查差异，不修改
"""
import json
import re
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH


# ============================================================
# 格式规则结构
# ============================================================

def make_rule():
    return {
        "font_cn": None,        # 中文字体
        "font_en": None,        # 英文字体
        "size_pt": None,        # 字号（磅）
        "bold": None,           # 加粗
        "italic": None,         # 斜体
        "alignment": None,      # 对齐: left/center/right/justify
        "line_spacing_pt": None,  # 固定行距（磅）
        "line_spacing_rule": None,  # 行距规则
        "space_before_pt": None,   # 段前间距
        "space_after_pt": None,    # 段后间距
        "first_line_indent_chars": None,  # 首行缩进（字符数）
        "left_indent_chars": None,        # 左侧缩进（字符数）
    }


# ============================================================
# 段落分类
# ============================================================

def classify_paragraph(para):
    """根据段落内容和样式判断类型。"""
    text = para.text.strip()
    if not text:
        return "empty"

    style_name = para.style.name if para.style else ""

    # TOC styles
    if style_name.startswith("toc"):
        return "toc"

    # Heading styles
    if style_name.startswith("Heading 1") or style_name == "heading 1":
        return "heading1"
    if style_name.startswith("Heading 2") or style_name == "heading 2":
        return "heading2"
    if style_name.startswith("Heading 3") or style_name == "heading 3":
        return "heading3"
    if style_name.startswith("Heading") or style_name == "heading":
        return "heading_other"

    # Content-based detection
    if re.match(r'^\d+\s+\S', text) and style_name.startswith("Heading"):
        return "heading1"

    # TOC title
    if text in ("目  录", "目 录", "目录"):
        return "toc_title"

    # Reference heading
    if text in ("参考文献", "参 考 文 献"):
        return "ref_heading"

    # Acknowledgment heading
    if text in ("致谢", "致  谢", "致　谢"):
        return "ack_heading"

    # References
    if re.match(r'^\[\d+\]', text):
        return "reference"

    # Figure/table captions
    if re.match(r'^(图|表|Fig|Table)\s*\d', text):
        return "caption"

    # Table notes
    if text.startswith("注：") or text.startswith("注:"):
        return "table_note"

    # Abstract labels
    if text.startswith("摘要") or text.startswith("摘 要"):
        return "abstract_cn"
    if text.startswith("关键词") or text.startswith("关键 词"):
        return "keywords_cn"
    if re.match(r'^Abstract', text, re.IGNORECASE):
        return "abstract_en"
    if re.match(r'^Key\s*words', text, re.IGNORECASE):
        return "keywords_en"

    # English title (all English, centered, large font)
    if all(ord(c) < 128 or c in ' \t' for c in text) and len(text) > 20:
        run = para.runs[0] if para.runs else None
        if run and run.font.size and run.font.size.pt >= 14:
            return "title_en"

    # Chinese title (first large centered text)
    run = para.runs[0] if para.runs else None
    if run and run.font.size and run.font.size.pt >= 14:
        align = para.alignment
        if align == WD_ALIGN_PARAGRAPH.CENTER or str(align) == "CENTER (1)":
            return "title_cn"

    return "body"


# ============================================================
# 学习格式
# ============================================================

def get_font_info(para):
    """从段落的第一个 run 提取字体信息。"""
    info = {}
    if not para.runs:
        return info
    run = para.runs[0]
    font = run.font
    if font.name:
        info["font_en"] = font.name
    if font.size:
        info["size_pt"] = round(font.size.pt, 1)
    if font.bold is not None:
        info["bold"] = font.bold
    if font.italic is not None:
        info["italic"] = font.italic

    # 尝试从 XML 获取中文字体
    from lxml import etree
    ns_w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    rpr = run._element.find(f"{ns_w}rPr")
    if rpr is not None:
        fonts = rpr.find(f"{ns_w}rFonts")
        if fonts is not None:
            east = fonts.get(f"{ns_w}eastAsia")
            if east:
                info["font_cn"] = east
    return info


def get_para_info(para):
    """提取段落格式信息。"""
    info = {}
    pf = para.paragraph_format

    if pf.alignment is not None:
        info["alignment"] = str(pf.alignment)

    if pf.line_spacing is not None:
        if isinstance(pf.line_spacing, float):
            info["line_spacing_multiple"] = pf.line_spacing
        else:
            info["line_spacing_pt"] = round(pf.line_spacing.pt, 1)

    if pf.line_spacing_rule is not None:
        info["line_spacing_rule"] = str(pf.line_spacing_rule)

    if pf.space_before is not None:
        info["space_before_pt"] = round(pf.space_before.pt, 1)
    if pf.space_after is not None:
        info["space_after_pt"] = round(pf.space_after.pt, 1)

    if pf.first_line_indent is not None:
        info["first_line_indent_pt"] = round(pf.first_line_indent.pt, 1)

    if pf.left_indent is not None:
        info["left_indent_pt"] = round(pf.left_indent.pt, 1)

    return info


def learn_format(docx_path):
    """从参考文档学习格式规则。"""
    doc = Document(str(docx_path))
    rules = {}

    # 页面设置
    section = doc.sections[0]
    rules["_page"] = {
        "top_cm": round(section.top_margin.cm, 2),
        "bottom_cm": round(section.bottom_margin.cm, 2),
        "left_cm": round(section.left_margin.cm, 2),
        "right_cm": round(section.right_margin.cm, 2),
        "header_cm": round(section.header_distance.cm, 2),
        "footer_cm": round(section.footer_distance.cm, 2),
    }

    # 按类型收集样本
    samples = {}
    for para in doc.paragraphs:
        ptype = classify_paragraph(para)
        if ptype in ("empty", "toc"):
            continue
        if ptype not in samples:
            samples[ptype] = []
        samples[ptype].append({
            "text": para.text[:50],
            "font": get_font_info(para),
            "para": get_para_info(para),
        })

    # 每种类型取最常见的格式作为规则
    for ptype, items in samples.items():
        if not items:
            continue

        # 合并字体信息（取第一个非空值）
        font_rule = {}
        for item in items:
            for k, v in item["font"].items():
                if k not in font_rule and v is not None:
                    font_rule[k] = v

        # 合并段落信息
        para_rule = {}
        for item in items:
            for k, v in item["para"].items():
                if k not in para_rule and v is not None:
                    para_rule[k] = v

        rules[ptype] = {**font_rule, **para_rule}

    return rules


# ============================================================
# 应用格式
# ============================================================

ALIGN_MAP = {
    "LEFT (0)": WD_ALIGN_PARAGRAPH.LEFT,
    "CENTER (1)": WD_ALIGN_PARAGRAPH.CENTER,
    "RIGHT (2)": WD_ALIGN_PARAGRAPH.RIGHT,
    "JUSTIFY (3)": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "DISTRIBUTE (4)": WD_ALIGN_PARAGRAPH.DISTRIBUTE,
}


def apply_font(run, rule):
    """应用字体规则到 run。"""
    if not rule:
        return
    font = run.font
    if "size_pt" in rule and rule["size_pt"]:
        font.size = Pt(rule["size_pt"])
    if "bold" in rule and rule["bold"] is not None:
        font.bold = rule["bold"]
    if "italic" in rule and rule["italic"] is not None:
        font.italic = rule["italic"]
    if "font_en" in rule and rule["font_en"]:
        font.name = rule["font_en"]

    # 中文字体需要通过 XML 设置
    if "font_cn" in rule and rule["font_cn"]:
        from lxml import etree
        ns_w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        rpr = run._element.find(f"{ns_w}rPr")
        if rpr is None:
            rpr = etree.SubElement(run._element, f"{ns_w}rPr")
        fonts = rpr.find(f"{ns_w}rFonts")
        if fonts is None:
            fonts = etree.SubElement(rpr, f"{ns_w}rFonts")
        fonts.set(f"{ns_w}eastAsia", rule["font_cn"])


def apply_para_format(para, rule):
    """应用段落格式规则。"""
    if not rule:
        return
    pf = para.paragraph_format

    if "alignment" in rule and rule["alignment"]:
        pf.alignment = ALIGN_MAP.get(rule["alignment"])

    if "line_spacing_pt" in rule and rule["line_spacing_pt"]:
        from docx.shared import Pt
        pf.line_spacing = Pt(rule["line_spacing_pt"])

    if "space_before_pt" in rule and rule["space_before_pt"] is not None:
        pf.space_before = Pt(rule["space_before_pt"])
    if "space_after_pt" in rule and rule["space_after_pt"] is not None:
        pf.space_after = Pt(rule["space_after_pt"])

    if "first_line_indent_pt" in rule and rule["first_line_indent_pt"]:
        pf.first_line_indent = Pt(rule["first_line_indent_pt"])
    if "left_indent_pt" in rule and rule["left_indent_pt"]:
        pf.left_indent = Pt(rule["left_indent_pt"])


def apply_page_setup(doc, page_rule):
    """应用页面设置。"""
    if not page_rule:
        return
    for section in doc.sections:
        section.top_margin = Cm(page_rule.get("top_cm", 2.7))
        section.bottom_margin = Cm(page_rule.get("bottom_cm", 2.7))
        section.left_margin = Cm(page_rule.get("left_cm", 2.7))
        section.right_margin = Cm(page_rule.get("right_cm", 2.7))
        if "header_cm" in page_rule:
            section.header_distance = Cm(page_rule["header_cm"])
        if "footer_cm" in page_rule:
            section.footer_distance = Cm(page_rule["footer_cm"])


def apply_format(docx_path, rules, output_path=None):
    """将格式规则应用到文档。"""
    doc = Document(str(docx_path))

    # 页面设置
    apply_page_setup(doc, rules.get("_page", {}))

    stats = {k: 0 for k in rules if not k.startswith("_")}

    for para in doc.paragraphs:
        ptype = classify_paragraph(para)
        if ptype == "empty":
            continue

        rule = rules.get(ptype)
        if not rule:
            continue

        # 应用段落格式
        apply_para_format(para, rule)

        # 应用字体格式到所有 runs
        for run in para.runs:
            apply_font(run, rule)

        stats[ptype] = stats.get(ptype, 0) + 1

    # 保存
    if not output_path:
        p = Path(docx_path)
        output_path = p.parent / f"{p.stem}_formatted{p.suffix}"
    doc.save(str(output_path))
    return output_path, stats


# ============================================================
# 检查差异
# ============================================================

def check_format(docx_path, rules):
    """检查文档格式与规则的差异。"""
    doc = Document(str(docx_path))
    issues = []

    for i, para in enumerate(doc.paragraphs):
        ptype = classify_paragraph(para)
        if ptype in ("empty", "toc"):
            continue

        rule = rules.get(ptype)
        if not rule:
            continue

        font = get_font_info(para)
        para_info = get_para_info(para)
        text = para.text[:40]

        # 检查字号
        if "size_pt" in rule and rule["size_pt"]:
            actual = font.get("size_pt")
            if actual and abs(actual - rule["size_pt"]) > 0.5:
                issues.append(f"[{ptype}] para {i}: 字号 {actual}pt != {rule['size_pt']}pt | {text}")

        # 检查加粗
        if "bold" in rule and rule["bold"] is not None:
            actual = font.get("bold", False)
            if actual != rule["bold"]:
                issues.append(f"[{ptype}] para {i}: bold={actual} != {rule['bold']} | {text}")

        # 检查对齐
        if "alignment" in rule and rule["alignment"]:
            actual = para_info.get("alignment", "")
            if actual and actual != rule["alignment"]:
                issues.append(f"[{ptype}] para {i}: align={actual} != {rule['alignment']} | {text}")

    return issues


# ============================================================
# CLI
# ============================================================

def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    cmd = sys.argv[1]

    if cmd == "learn":
        ref_path = sys.argv[2]
        rules = learn_format(ref_path)
        output = sys.argv[3] if len(sys.argv) > 3 else "rules.json"
        with open(output, "w", encoding="utf-8") as f:
            json.dump(rules, f, ensure_ascii=False, indent=2)
        print(f"Learned {len([k for k in rules if not k.startswith('_')])} format rules")
        print(f"Page setup: {rules.get('_page', {})}")
        for k, v in rules.items():
            if not k.startswith("_"):
                print(f"  {k}: {v}")
        print(f"Saved to: {output}")

    elif cmd == "apply":
        docx_path = sys.argv[2]
        rules_path = sys.argv[3]
        with open(rules_path, "r", encoding="utf-8") as f:
            rules = json.load(f)
        output_path, stats = apply_format(docx_path, rules)
        print(f"Applied formatting to: {output_path}")
        for k, count in stats.items():
            if count > 0:
                print(f"  {k}: {count} paragraphs")

    elif cmd == "check":
        docx_path = sys.argv[2]
        rules_path = sys.argv[3]
        with open(rules_path, "r", encoding="utf-8") as f:
            rules = json.load(f)
        issues = check_format(docx_path, rules)
        if issues:
            print(f"Found {len(issues)} issues:")
            for issue in issues:
                print(f"  {issue}")
        else:
            print("No issues found!")

    else:
        print(f"Unknown command: {cmd}")
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
