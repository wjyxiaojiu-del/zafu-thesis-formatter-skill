#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
浙农林大学毕业论文格式修复脚本
用法:
  python fix_zafu_thesis.py input.docx [output.docx]
  python fix_zafu_thesis.py <unpacked_dir>
"""
import sys
import os
import re
import copy
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET

ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ns_mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

# Register all namespaces to preserve them on write
namespaces = {
    'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
    'cx1': 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
    'cx2': 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
    'cx3': 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex',
    'cx4': 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex',
    'cx5': 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex',
    'cx6': 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex',
    'cx7': 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex',
    'cx8': 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex',
    'mc': ns_mc,
    'aink': 'http://schemas.microsoft.com/office/drawing/2016/ink',
    'am3d': 'http://schemas.microsoft.com/office/drawing/2017/model3d',
    'o': 'urn:schemas-microsoft-com:office:office',
    'oel': 'http://schemas.microsoft.com/office/2019/extlst',
    'r': ns_r,
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v': 'urn:schemas-microsoft-com:vml',
    'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'w': ns_w,
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'w16cex': 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
    'w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
    'w16': 'http://schemas.microsoft.com/office/word/2018/wordml',
    'w16du': 'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
    'w16sdtdh': 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
    'w16sdtfl': 'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock',
    'w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
}
for prefix, uri in namespaces.items():
    ET.register_namespace(prefix, uri)


def w(tag):
    return f'{{{ns_w}}}{tag}'


def get_all_text(elem):
    texts = []
    for t in elem.iter(w('t')):
        if t.text:
            texts.append(t.text)
    return ''.join(texts).strip()


def set_run_font(run, east_asia, ascii_font='Times New Roman', h_ansi='Times New Roman', size=None, bold=None, bold_cs=None):
    """Set font properties on a w:r element."""
    rpr = run.find(w('rPr'))
    if rpr is None:
        rpr = ET.SubElement(run, w('rPr'))

    # Set fonts
    fonts = rpr.find(w('rFonts'))
    if fonts is None:
        fonts = ET.SubElement(rpr, w('rFonts'))
    fonts.set(w('eastAsia'), east_asia)
    fonts.set(w('ascii'), ascii_font)
    fonts.set(w('hAnsi'), h_ansi)

    # Set size
    if size is not None:
        sz = rpr.find(w('sz'))
        if sz is None:
            sz = ET.SubElement(rpr, w('sz'))
        sz.set(w('val'), str(size))
        szCs = rpr.find(w('szCs'))
        if szCs is None:
            szCs = ET.SubElement(rpr, w('szCs'))
        szCs.set(w('val'), str(size))

    # Set bold
    if bold is not None:
        b = rpr.find(w('b'))
        if bold:
            if b is None:
                b = ET.SubElement(rpr, w('b'))
            if w('val') in b.attrib:
                del b.attrib[w('val')]
        else:
            if b is not None:
                rpr.remove(b)
            b = ET.SubElement(rpr, w('b'))
            b.set(w('val'), '0')

    if bold_cs is not None:
        bcs = rpr.find(w('bCs'))
        if bold_cs:
            if bcs is None:
                bcs = ET.SubElement(rpr, w('bCs'))
            if w('val') in bcs.attrib:
                del bcs.attrib[w('val')]
        else:
            if bcs is not None:
                rpr.remove(bcs)
            bcs = ET.SubElement(rpr, w('bCs'))
            bcs.set(w('val'), '0')


def set_para_alignment(para, align):
    """Set paragraph alignment: center, both, left, right."""
    ppr = para.find(w('pPr'))
    if ppr is None:
        ppr = ET.SubElement(para, w('pPr'))
    jc = ppr.find(w('jc'))
    if jc is None:
        jc = ET.SubElement(ppr, w('jc'))
    jc.set(w('val'), align)


def set_para_spacing(para, before=None, after=None, line=None, line_rule=None):
    """Set paragraph spacing."""
    ppr = para.find(w('pPr'))
    if ppr is None:
        ppr = ET.SubElement(para, w('pPr'))
    spacing = ppr.find(w('spacing'))
    if spacing is None:
        spacing = ET.SubElement(ppr, w('spacing'))
    if before is not None:
        spacing.set(w('before'), str(before))
    if after is not None:
        spacing.set(w('after'), str(after))
    if line is not None:
        spacing.set(w('line'), str(line))
    if line_rule is not None:
        spacing.set(w('lineRule'), line_rule)


def set_para_indent(para, first_line=None, first_line_chars=None, left=None, left_chars=None, hanging=None, hanging_chars=None):
    """Set paragraph indent. When setting left/hanging indent, clear conflicting firstLine; when setting firstLine, clear left/hanging."""
    ppr = para.find(w('pPr'))
    if ppr is None:
        ppr = ET.SubElement(para, w('pPr'))
    ind = ppr.find(w('ind'))
    if ind is None:
        ind = ET.SubElement(ppr, w('ind'))

    # Clear conflicting indent properties
    if left is not None or left_chars is not None or hanging is not None or hanging_chars is not None:
        # Setting left or hanging indent: clear firstLine
        if w('firstLine') in ind.attrib:
            del ind.attrib[w('firstLine')]
        if w('firstLineChars') in ind.attrib:
            del ind.attrib[w('firstLineChars')]
    if first_line is not None or first_line_chars is not None:
        # Setting firstLine indent: clear left and hanging
        if w('left') in ind.attrib:
            del ind.attrib[w('left')]
        if w('leftChars') in ind.attrib:
            del ind.attrib[w('leftChars')]
        if w('hanging') in ind.attrib:
            del ind.attrib[w('hanging')]
        if w('hangingChars') in ind.attrib:
            del ind.attrib[w('hangingChars')]

    if first_line is not None:
        ind.set(w('firstLine'), str(first_line))
    if first_line_chars is not None:
        ind.set(w('firstLineChars'), str(first_line_chars))
    if left is not None:
        ind.set(w('left'), str(left))
    if left_chars is not None:
        ind.set(w('leftChars'), str(left_chars))
    if hanging is not None:
        ind.set(w('hanging'), str(hanging))
    if hanging_chars is not None:
        ind.set(w('hangingChars'), str(hanging_chars))


def disable_para_numbering(para):
    """Disable inherited Word auto-numbering so explicit thesis numbers are used."""
    ppr = para.find(w('pPr'))
    if ppr is None:
        ppr = ET.SubElement(para, w('pPr'))
    num_pr = ppr.find(w('numPr'))
    if num_pr is None:
        num_pr = ET.SubElement(ppr, w('numPr'))
    for child in list(num_pr):
        num_pr.remove(child)
    num_id = ET.SubElement(num_pr, w('numId'))
    num_id.set(w('val'), '0')


def load_style_names(styles_path):
    """Return {styleId: styleName} for safer style-based classification."""
    if not os.path.exists(styles_path):
        return {}
    try:
        tree = ET.parse(styles_path)
    except ET.ParseError:
        return {}
    names = {}
    for style in tree.getroot().findall(w('style')):
        sid = style.get(w('styleId'), '')
        name_elem = style.find(w('name'))
        if sid and name_elem is not None:
            names[sid] = name_elem.get(w('val'), '')
    return names


def classify_paragraph(text, style_id, style_name=''):
    """Classify paragraph type based on text content and style."""
    t = text.strip()
    style_name_lower = style_name.lower()

    if not t:
        return 'empty'

    # TOC entries
    if style_id in ('TOC1', 'TOC2', 'TOC3', 'toc 1', 'toc 2', 'toc 3') or 'toc' in style_name_lower:
        return 'toc'

    # References
    if re.match(r'^\[\d+\]', t):
        return 'reference'

    # Heading patterns. Check deeper dotted levels first so 1.1.1.1 is not
    # accidentally classified as 1.1.1.
    if re.match(r'^[1-7]\.\d+\.\d+\.\d+\s*\S*', t):
        return 'h4'

    if re.match(r'^[1-7]\.\d+\.\d+\s*\S*', t):
        return 'h3'

    if re.match(r'^[1-7]\.\d+\s*\S*', t):
        return 'h2'

    # H1: "N.文字" format like "1.文献综述" (numbered chapters with dot)
    if re.match(r'^[1-7]\.[^\d]', t) and len(t) > 3:
        return 'h1'

    # H1: "N 文字" format like "1 文献综述" "4 讨论" (numbered with space)
    if re.match(r'^[1-7]\s+\S', t) and len(t) > 2:
        return 'h1'

    # H1: "N文字" format like "1文献综述" (numbered without separator)
    if re.match(r'^[1-7][\u4e00-\u9fff]', t) and len(t) > 2:
        return 'h1'

    # Special headings: "引言", "致谢" etc. use H1 format (楷体加粗四号居中)
    if t == '致谢':
        return 'ack_heading'
    if t in ('引言', '附录', '结论'):
        return 'h1'
    # "参考文献" heading: 楷体加粗四号居中，段前段后6磅
    if t == '参考文献':
        return 'ref_heading'

    # Some Word files use localized built-in heading style ids where the visible
    # heading number comes from style numbering rather than paragraph text.
    if style_id in ('2', 'Heading1', 'heading 1') or style_name_lower == 'heading 1':
        return 'h1'
    if style_id in ('3', 'Heading2', 'heading 2') or style_name_lower == 'heading 2':
        return 'h2'
    if style_id in ('4', 'Heading3', 'heading 3') or style_name_lower == 'heading 3':
        return 'h3'
    if style_id in ('Heading4', 'heading 4') or style_name_lower == 'heading 4':
        return 'h4'

    # Table/figure captions
    if re.match(r'^表\s*\d+', t) or re.match(r'^图\s*\d+', t):
        return 'caption'

    # Table notes below tables, usually "注：...".
    if t.startswith('注：') or t.startswith('注:'):
        return 'table_note'

    # Abstract labels
    if t.startswith('摘要') and ('：' in t or ':' in t or t == '摘要' or t == '摘要：'):
        return 'abstract_cn'
    if t.startswith('关键词') and ('：' in t or ':' in t or t == '关键词' or t == '关键词：'):
        return 'keywords_cn'
    if t == 'Abstract' or t.startswith('Abstract:') or t.startswith('Abstract：'):
        return 'abstract_en'
    if (t in ('Keywords', 'Key words', 'Key Words') or
            t.startswith('Keywords:') or
            t.startswith('Keywords：') or
            t.startswith('Key words:') or
            t.startswith('Key words：') or
            t.startswith('Key Words：') or
            t.startswith('Key Words:')):
        return 'keywords_en'

    # English title: all ASCII/Latin text, no Chinese chars, length > 10, not a label
    if len(t) > 10 and not any('\u4e00' <= c <= '\u9fff' for c in t):
        # Check if it looks like a title (not starting with number like a heading)
        if not re.match(r'^\d', t) and ':' not in t and 'Abstract' not in t:
            return 'en_title'

    # Formula lines
    if t.startswith('式中') or t.startswith('公式') or re.match(r'^[A-Z]', t) and '=' in t:
        return 'formula'

    return 'body'


def fix_abstract_paragraph(para, text, label_font, content_font, content_size):
    """Fix abstract/keywords paragraph: adjust fonts on existing runs without destroying structure."""
    runs = para.findall(w('r'))
    if not runs:
        return

    for i, run in enumerate(runs):
        rt = run.find(w('t'))
        if rt is None or not rt.text:
            continue

        run_text = rt.text.strip()
        if i == 0:
            # First run = label ("摘要"/"关键词"/"Abstract"/"Keywords")
            set_run_font(run, label_font, ascii_font=label_font, h_ansi=label_font,
                         size=content_size, bold=True, bold_cs=True)
        elif ':' in run_text or '\uff1a' in run_text:
            # Separator run - make it part of label style
            set_run_font(run, label_font, ascii_font=label_font, h_ansi=label_font,
                         size=content_size, bold=True, bold_cs=True)
        else:
            # Content runs
            set_run_font(run, content_font, size=content_size, bold=False, bold_cs=False)


def normalize_keyword_separators(para, language):
    """Normalize keyword labels and separators from detector feedback."""
    seen_label = False
    for run in para.findall(w('r')):
        rt = run.find(w('t'))
        if rt is None or not rt.text:
            continue
        text = rt.text
        stripped = text.strip()

        if language == 'cn':
            if stripped.startswith('关键词'):
                seen_label = True
                continue
            if seen_label:
                rt.text = text.replace(';', '；').replace(',', '；')
        else:
            if stripped in ('Keywords', 'Key Words', 'Key words'):
                rt.text = text.replace(stripped, 'Key words')
                seen_label = True
                continue
            if stripped in (':', '：'):
                rt.text = ':'
                seen_label = True
                continue
            if seen_label:
                normalized = text.replace('；', ', ').replace(';', ', ')
                normalized = re.sub(r',\s*', ', ', normalized)
                rt.text = normalized


def fix_heading_spacing_after_number(para, level):
    """Replace ASCII space after heading number with Chinese full-width space (U+3000).
    Strategy: collect all run texts, build full string, find the transition from
    number (digits+dots) to content, then modify the appropriate run(s) in place."""
    runs = para.findall(w('r'))
    # Build a list of (run_index, run, text) for runs that have <w:t> with text
    rt_info = []
    for i, run in enumerate(runs):
        rt = run.find(w('t'))
        if rt is not None and rt.text:
            rt_info.append((i, run, rt))

    if not rt_info:
        return

    # Check if the heading starts with a number pattern like "1", "1.1", "3.1", etc.
    full_text = ''.join(info[2].text for info in rt_info)
    m = re.match(r'^(\d+(?:\.\d+)*)', full_text)
    if not m:
        return  # No number prefix, nothing to fix

    num_len = m.end()  # Length of the number prefix (e.g. "1.1" = 3)

    # Now walk through runs, tracking consumed chars from the number
    consumed = 0
    number_end_idx = None  # Index into rt_info where number portion ends

    for idx, (run_idx, run, rt) in enumerate(rt_info):
        run_text = rt.text
        remaining_num = num_len - consumed

        if remaining_num <= 0:
            # We're past the number - check for leading ASCII spaces
            if run_text.startswith(' '):
                # Replace leading ASCII spaces with one full-width space
                space_count = len(run_text) - len(run_text.lstrip(' '))
                rest = run_text[space_count:]
                rt.text = '\u3000' + rest
                number_end_idx = idx
            else:
                # No space, insert full-width space at beginning
                rt.text = '\u3000' + run_text
                number_end_idx = idx
            break

        if len(run_text) <= remaining_num:
            # This entire run is part of the number
            consumed += len(run_text)
        else:
            # This run contains the transition: number + space + content
            # e.g. '.1 ' -> '.1' is number, ' ' is space, '' is content
            # e.g. '2.1实验' -> '2.1' is number, '' is space, '实验' is content
            num_part = run_text[:remaining_num]
            after_num = run_text[remaining_num:]

            if after_num.startswith(' '):
                space_count = len(after_num) - len(after_num.lstrip(' '))
                rest = after_num[space_count:]
                rt.text = num_part + '\u3000' + rest
            elif after_num:
                # No space but content immediately follows (e.g. '2.1实验材料')
                rt.text = num_part + '\u3000' + after_num
            else:
                # Number ends exactly at run boundary, will handle in next run
                pass
            number_end_idx = idx
            break

    # If number ended exactly at a run boundary, check the next run for space
    if number_end_idx is not None:
        next_idx = number_end_idx + 1
        if next_idx < len(rt_info):
            next_rt = rt_info[next_idx][2]
            if next_rt.text.startswith(' '):
                space_count = len(next_rt.text) - len(next_rt.text.lstrip(' '))
                rest = next_rt.text[space_count:]
                next_rt.text = '\u3000' + rest


def set_heading_number_text(para, number):
    """Materialize heading numbers and use one full-width space after them."""
    disable_para_numbering(para)
    runs = para.findall(w('r'))
    text_runs = []
    for run in runs:
        rt = run.find(w('t'))
        if rt is not None:
            text_runs.append(rt)
    if not text_runs:
        run = ET.SubElement(para, w('r'))
        rt = ET.SubElement(run, w('t'))
        text_runs.append(rt)

    full_text = ''.join(rt.text or '' for rt in text_runs).strip()
    title = re.sub(r'^\d+(?:\.\d+)*[\.、]?[ \t\u3000]*', '', full_text).strip()
    title = title.lstrip('\u3000').strip()
    text_runs[0].text = f'{number}\u3000{title}' if title else str(number)
    for rt in text_runs[1:]:
        rt.text = ''


def fix_heading(para, text, level, number=None):
    """Fix heading formatting based on level (1, 2, 3, 4)."""
    if number is not None:
        set_heading_number_text(para, number)

    # Fix all runs in the paragraph
    for run in para.findall(w('r')):
        if level == 1:
            set_run_font(run, '楷体', size=28, bold=True, bold_cs=True)
        elif level == 2:
            set_run_font(run, '黑体', size=24, bold=True, bold_cs=True)
        elif level == 3:
            set_run_font(run, '黑体', size=21, bold=False, bold_cs=False)
        elif level == 4:
            set_run_font(run, '宋体', size=21, bold=False, bold_cs=False)

    # Fix spacing on pre-numbered headings. Computed headings already use one
    # full-width space from set_heading_number_text().
    if number is None:
        fix_heading_spacing_after_number(para, level)

    # Fix paragraph properties
    if level == 1:
        set_para_alignment(para, 'center')
        set_para_spacing(para, before=120, after=120, line='400', line_rule='exact')
        # Remove indent for centered headings
        ppr = para.find(w('pPr'))
        if ppr is not None:
            ind = ppr.find(w('ind'))
            if ind is not None:
                ppr.remove(ind)
    elif level in (2, 3, 4):
        set_para_alignment(para, 'left')
        set_para_spacing(para, before=60, after=60, line='400', line_rule='exact')
        # 左侧缩进2字符（不是首行缩进！）
        set_para_indent(para, left='480', left_chars='200')


def fix_body_paragraph(para):
    """Fix normal body paragraph formatting."""
    for run in para.findall(w('r')):
        # Check if this run has math content - skip math runs
        has_math = False
        for child in run:
            tag = child.tag.split('}')[1] if '}' in child.tag else child.tag
            if 'math' in tag.lower() or 'oMath' in tag:
                has_math = True
                break
        if has_math:
            continue

        # Check if fonts are already correct (eastAsia=宋体, ascii/hAnsi=TNR)
        rpr = run.find(w('rPr'))
        if rpr is not None:
            fonts = rpr.find(w('rFonts'))
            if fonts is not None:
                ea = fonts.get(w('eastAsia'), '')
                asc = fonts.get(w('ascii'), '')
                hansi = fonts.get(w('hAnsi'), '')
                # Only skip if ALL fonts are already correct
                if ea == '宋体' and asc == 'Times New Roman' and hansi == 'Times New Roman':
                    sz_elem = rpr.find(w('sz'))
                    sz_val = sz_elem.get(w('val'), '') if sz_elem is not None else ''
                    if sz_val == '21':
                        continue

        set_run_font(run, '宋体', size=21, bold=False, bold_cs=False)

    # Set paragraph properties
    set_para_alignment(para, 'both')
    set_para_spacing(para, before='0', after='0', line='400', line_rule='exact')
    set_para_indent(para, first_line='480', first_line_chars='200')


def fix_reference_paragraph(para):
    """Fix reference entry paragraph: 宋体五号左对齐，悬挂缩进2字符."""
    for run in para.findall(w('r')):
        set_run_font(run, '宋体', size=21, bold=False, bold_cs=False)

    set_para_alignment(para, 'left')
    set_para_spacing(para, before='0', after='0', line='400', line_rule='exact')
    # 悬挂缩进2字符（参考文献条目）
    set_para_indent(para, hanging='480', hanging_chars='200')


def fix_reference_heading(para):
    """Fix '参考文献' heading: 楷体加粗四号居中，段前段后6磅."""
    for run in para.findall(w('r')):
        set_run_font(run, '楷体', size=28, bold=True, bold_cs=True)
    disable_para_numbering(para)
    set_para_alignment(para, 'center')
    set_para_spacing(para, before=120, after=120, line='400', line_rule='exact')
    # Remove indent
    ppr = para.find(w('pPr'))
    if ppr is not None:
        ind = ppr.find(w('ind'))
        if ind is not None:
            ppr.remove(ind)


def fix_ack_heading(para):
    """Fix acknowledgement heading: 楷体四号加粗居中，写作'致  谢'."""
    disable_para_numbering(para)
    text_runs = []
    for run in para.findall(w('r')):
        rt = run.find(w('t'))
        if rt is not None:
            text_runs.append(rt)
        set_run_font(run, '楷体', size=28, bold=True, bold_cs=True)
    if text_runs:
        text_runs[0].text = '致  谢'
        for rt in text_runs[1:]:
            rt.text = ''
    set_para_alignment(para, 'center')
    set_para_spacing(para, before=120, after=120, line='400', line_rule='exact')
    ppr = para.find(w('pPr'))
    if ppr is not None:
        ind = ppr.find(w('ind'))
        if ind is not None:
            ppr.remove(ind)


def fix_caption_paragraph(para):
    """Fix table/figure caption formatting."""
    for run in para.findall(w('r')):
        set_run_font(run, '宋体', size=18, bold=False, bold_cs=False)
    set_para_alignment(para, 'center')
    set_para_spacing(para, before='0', after='0', line='400', line_rule='exact')


def fix_table_note_paragraph(para):
    """Fix table notes: 宋体小五, English/numbers Times New Roman, after spacing 1 line."""
    for run in para.findall(w('r')):
        set_run_font(run, '宋体', size=18, bold=False, bold_cs=False)
    set_para_alignment(para, 'both')
    set_para_spacing(para, before='0', line='400', line_rule='exact')
    ppr = para.find(w('pPr'))
    if ppr is None:
        ppr = ET.SubElement(para, w('pPr'))
    spacing = ppr.find(w('spacing'))
    if spacing is None:
        spacing = ET.SubElement(ppr, w('spacing'))
    if w('after') in spacing.attrib:
        del spacing.attrib[w('after')]
    spacing.set(w('afterLines'), '100')


def fix_cn_punctuation(para):
    """Fix half-width punctuation in Chinese body/abstract text to full-width.
    Only fixes commas and periods in text that contains Chinese characters.
    Does NOT modify: Key Words content, English-only paragraphs, formulas, references."""
    changed = False
    for run in para.findall(w('r')):
        rt = run.find(w('t'))
        if rt is None or not rt.text:
            continue
        text = rt.text
        # Only process if the text contains Chinese characters
        if not any('\u4e00' <= c <= '\u9fff' for c in text):
            continue
        # Replace half-width punctuation with full-width
        new_text = text
        new_text = new_text.replace(',', '\uff0c')   # , → ，
        new_text = new_text.replace('.', '\u3002')    # . → 。 (only if surrounded by Chinese context)
        # Don't replace period in numbers like "3.14" or "www."
        # Actually, we need to be smarter about periods - only replace if it's a sentence-ending period
        # Revert: just do commas for safety, and let periods be handled manually if needed
        new_text = new_text.replace('\u3002', '.')   # revert period changes for now
        # Actually replace period only if followed by space/end or Chinese char
        result = []
        for i, ch in enumerate(text):
            if ch == ',':
                result.append('\uff0c')
            elif ch == '.':
                # Replace period with full-width if it looks like a sentence-ending period
                # (followed by end-of-string, space, or a Chinese char)
                next_ch = text[i+1] if i+1 < len(text) else ''
                prev_ch = text[i-1] if i > 0 else ''
                # Don't replace if between digits (like 3.14) or part of URL/abbreviation
                if prev_ch.isdigit() and (next_ch.isdigit() or next_ch == ' ' or next_ch == ''):
                    result.append('.')
                elif next_ch == '\u3002' or (prev_ch in '\u4e00\u4e01\u4e02' and next_ch == ''):
                    result.append('\u3002')
                elif prev_ch in '\u4e00\u4e01\u4e02' and next_ch in ' \u300a\u300b\uff08\uff09':
                    result.append('\u3002')
                else:
                    result.append('.')
            else:
                result.append(ch)
        new_text = ''.join(result)
        if new_text != text:
            rt.text = new_text
            changed = True
    return changed


def fix_title_paragraph(para):
    """Fix thesis title (中文): 黑体三号居中."""
    for run in para.findall(w('r')):
        set_run_font(run, '黑体', size=32, bold=False, bold_cs=False)
    set_para_alignment(para, 'center')
    set_para_spacing(para, line='400', line_rule='exact')
    # Remove indent
    ppr = para.find(w('pPr'))
    if ppr is not None:
        ind = ppr.find(w('ind'))
        if ind is not None:
            ppr.remove(ind)


def fix_en_title_paragraph(para):
    """Fix English title: Times New Roman 三号居中, 固定行距20磅."""
    for run in para.findall(w('r')):
        set_run_font(run, 'Times New Roman', ascii_font='Times New Roman',
                     h_ansi='Times New Roman', size=32, bold=False, bold_cs=False)
    set_para_alignment(para, 'center')
    set_para_spacing(para, line='400', line_rule='exact')
    # Remove indent
    ppr = para.find(w('pPr'))
    if ppr is not None:
        ind = ppr.find(w('ind'))
        if ind is not None:
            ppr.remove(ind)


def fix_page_setup(root):
    """Fix page margins in all sectPr elements."""
    fixes = 0
    for sect_pr in root.iter(w('sectPr')):
        pgMar = sect_pr.find(w('pgMar'))
        if pgMar is None:
            pgMar = ET.SubElement(sect_pr, w('pgMar'))
        pgMar.set(w('top'), '1531')
        pgMar.set(w('bottom'), '1531')
        pgMar.set(w('left'), '1531')
        pgMar.set(w('right'), '1531')
        pgMar.set(w('header'), '1021')
        pgMar.set(w('footer'), '1051')
        pgMar.set(w('gutter'), '0')

        # Ensure A4 page size
        pgSz = sect_pr.find(w('pgSz'))
        if pgSz is None:
            pgSz = ET.SubElement(sect_pr, w('pgSz'))
        pgSz.set(w('w'), '11906')
        pgSz.set(w('h'), '16838')
        fixes += 1

    return fixes


def pack_docx(unpacked_dir, output_path):
    """Pack an unpacked OOXML directory back into a DOCX file."""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
        content_types = os.path.join(unpacked_dir, '[Content_Types].xml')
        if os.path.exists(content_types):
            z.write(content_types, '[Content_Types].xml')
        for root_dir, _dirs, files in os.walk(unpacked_dir):
            for filename in files:
                file_path = os.path.join(root_dir, filename)
                arcname = os.path.relpath(file_path, unpacked_dir).replace(os.sep, '/')
                if arcname == '[Content_Types].xml':
                    continue
                z.write(file_path, arcname)


def fix_toc_styles(styles_path):
    """Fix TOC styles in styles.xml: 宋体 + Times New Roman + 五号(21).
    Sets font on TOC1, TOC2, TOC3 (and their 'toc 1'/'toc 2'/'toc 3' aliases).
    Also sets proper spacing and alignment."""
    if not os.path.exists(styles_path):
        return 0

    tree = ET.parse(styles_path)
    root = tree.getroot()

    # Match explicit TOC style ids/names only. Numeric style ids are unsafe
    # because Word may assign them to heading styles in localized documents.
    toc_style_ids = {'TOC1', 'TOC2', 'TOC3', 'toc 1', 'toc 2', 'toc 3'}
    fixed = 0

    for style in root.findall(w('style')):
        sid = style.get(w('styleId'), '')
        name_elem = style.find(w('name'))
        name = name_elem.get(w('val'), '') if name_elem is not None else ''

        if sid not in toc_style_ids and 'toc' not in name.lower():
            continue

        # Determine level: TOC1/toc 1 -> level 1, TOC2/toc 2 -> level 2, etc.
        if sid in ('TOC1', 'toc 1') or 'toc 1' in name.lower():
            level = 1
        elif sid in ('TOC2', 'toc 2') or 'toc 2' in name.lower():
            level = 2
        elif sid in ('TOC3', 'toc 3') or 'toc 3' in name.lower():
            level = 3
        else:
            level = 1

        # Set run properties: 宋体 + TNR + 五号(21)
        rpr = style.find(w('rPr'))
        if rpr is None:
            rpr = ET.SubElement(style, w('rPr'))

        fonts = rpr.find(w('rFonts'))
        if fonts is None:
            fonts = ET.SubElement(rpr, w('rFonts'))
        fonts.set(w('eastAsia'), '宋体')
        fonts.set(w('ascii'), 'Times New Roman')
        fonts.set(w('hAnsi'), 'Times New Roman')

        # Remove bold for TOC entries
        b = rpr.find(w('b'))
        if b is not None:
            rpr.remove(b)

        # Set size to 21 (五号)
        sz = rpr.find(w('sz'))
        if sz is None:
            sz = ET.SubElement(rpr, w('sz'))
        sz.set(w('val'), '21')
        szCs = rpr.find(w('szCs'))
        if szCs is None:
            szCs = ET.SubElement(rpr, w('szCs'))
        szCs.set(w('val'), '21')

        # Set paragraph properties: both alignment, fixed line spacing 20pt
        ppr = style.find(w('pPr'))
        if ppr is None:
            ppr = ET.SubElement(style, w('pPr'))

        spacing = ppr.find(w('spacing'))
        if spacing is None:
            spacing = ET.SubElement(ppr, w('spacing'))
        spacing.set(w('line'), '400')
        spacing.set(w('lineRule'), 'exact')
        spacing.set(w('before'), '0')
        spacing.set(w('after'), '0')

        jc = ppr.find(w('jc'))
        if jc is None:
            jc = ET.SubElement(ppr, w('jc'))
        jc.set(w('val'), 'both')

        fixed += 1
        print(f"  Fixed TOC style: {sid} ({name}) level={level}")

    if fixed > 0:
        tree.write(styles_path, xml_declaration=True, encoding='UTF-8')
    return fixed


def fix_toc_heading(para, text):
    """Fix '目录' heading: 宋体二号居中."""
    if text.strip() in ('目  录', '目 录', '目录'):
        for run in para.findall(w('r')):
            set_run_font(run, '宋体', size=44, bold=False, bold_cs=False)
        set_para_alignment(para, 'center')
        set_para_spacing(para, line='480', line_rule='exact')
        # Remove indent
        ppr = para.find(w('pPr'))
        if ppr is not None:
            ind = ppr.find(w('ind'))
            if ind is not None:
                ppr.remove(ind)
        return True
    return False


def main():
    if len(sys.argv) < 2:
        print("Usage: python fix_zafu_thesis.py input.docx [output.docx]")
        print("   or: python fix_zafu_thesis.py <unpacked_dir>")
        sys.exit(1)

    input_path = sys.argv[1]
    temp_dir = None
    output_docx = None

    if os.path.isfile(input_path) and input_path.lower().endswith('.docx'):
        output_docx = sys.argv[2] if len(sys.argv) >= 3 else os.path.splitext(input_path)[0] + '_formatted.docx'
        temp_dir = tempfile.mkdtemp(prefix='zafu_thesis_')
        unpacked_dir = os.path.join(temp_dir, 'unpacked')
        with zipfile.ZipFile(input_path, 'r') as z:
            z.extractall(unpacked_dir)
    else:
        unpacked_dir = input_path

    doc_xml_path = os.path.join(unpacked_dir, 'word', 'document.xml')
    styles_path = os.path.join(unpacked_dir, 'word', 'styles.xml')

    if not os.path.exists(doc_xml_path):
        print(f"Error: {doc_xml_path} not found")
        sys.exit(1)

    # Parse XML
    tree = ET.parse(doc_xml_path)
    root = tree.getroot()
    body = root.find(w('body'))
    style_names = load_style_names(styles_path)

    stats = {
        'page_setup': 0,
        'title': 0,
        'en_title': 0,
        'h1': 0,
        'h2': 0,
        'h3': 0,
        'h4': 0,
        'abstract': 0,
        'body': 0,
        'reference': 0,
        'ref_heading': 0,
        'ack_heading': 0,
        'caption': 0,
        'table_note': 0,
        'toc': 0,
        'en_abstract': 0,
        'punctuation': 0,
    }

    # 1. Fix page setup
    stats['page_setup'] = fix_page_setup(root)

    # 2. Process each paragraph
    heading_counters = [0, 0, 0, 0]
    for para in body.findall(w('p')):
        text = get_all_text(para)
        if not text:
            continue

        # Get style
        ppr = para.find(w('pPr'))
        style_id = ''
        if ppr is not None:
            ps = ppr.find(w('pStyle'))
            if ps is not None:
                style_id = ps.get(w('val'), '')
        style_name = style_names.get(style_id, '')

        ptype = classify_paragraph(text, style_id, style_name)

        if ptype == 'toc':
            # Fix TOC heading
            if fix_toc_heading(para, text):
                stats['toc'] += 1
            continue

        if ptype == 'en_title':
            fix_en_title_paragraph(para)
            stats['en_title'] += 1
        elif ptype == 'h1':
            heading_counters[0] += 1
            heading_counters[1:] = [0, 0, 0]
            number = str(heading_counters[0])
            fix_heading(para, text, 1, number)
            stats['h1'] += 1
        elif ptype == 'h2':
            if heading_counters[0] == 0:
                heading_counters[0] = 1
            heading_counters[1] += 1
            heading_counters[2:] = [0, 0]
            number = f'{heading_counters[0]}.{heading_counters[1]}'
            fix_heading(para, text, 2, number)
            stats['h2'] += 1
        elif ptype == 'h3':
            if heading_counters[0] == 0:
                heading_counters[0] = 1
            if heading_counters[1] == 0:
                heading_counters[1] = 1
            heading_counters[2] += 1
            heading_counters[3] = 0
            number = f'{heading_counters[0]}.{heading_counters[1]}.{heading_counters[2]}'
            fix_heading(para, text, 3, number)
            stats['h3'] += 1
        elif ptype == 'h4':
            if heading_counters[0] == 0:
                heading_counters[0] = 1
            if heading_counters[1] == 0:
                heading_counters[1] = 1
            if heading_counters[2] == 0:
                heading_counters[2] = 1
            heading_counters[3] += 1
            number = f'{heading_counters[0]}.{heading_counters[1]}.{heading_counters[2]}.{heading_counters[3]}'
            fix_heading(para, text, 4, number)
            stats['h4'] += 1
        elif ptype == 'abstract_cn':
            fix_abstract_paragraph(para, text, '黑体', '楷体', 21)
            # 标签用左侧缩进2字符
            set_para_indent(para, left='480', left_chars='200')
            set_para_spacing(para, line='400', line_rule='exact')
            stats['abstract'] += 1
        elif ptype == 'keywords_cn':
            fix_abstract_paragraph(para, text, '黑体', '楷体', 21)
            normalize_keyword_separators(para, 'cn')
            # 标签用左侧缩进2字符
            set_para_indent(para, left='480', left_chars='200')
            set_para_spacing(para, line='400', line_rule='exact')
            stats['abstract'] += 1
        elif ptype == 'abstract_en':
            fix_abstract_paragraph(para, text, 'Times New Roman', 'Times New Roman', 21)
            # 标签用左侧缩进2字符
            set_para_indent(para, left='480', left_chars='200')
            set_para_spacing(para, line='400', line_rule='exact')
            stats['en_abstract'] += 1
        elif ptype == 'keywords_en':
            fix_abstract_paragraph(para, text, 'Times New Roman', 'Times New Roman', 21)
            normalize_keyword_separators(para, 'en')
            # 标签用左侧缩进2字符
            set_para_indent(para, left='480', left_chars='200')
            set_para_spacing(para, line='400', line_rule='exact')
            stats['en_abstract'] += 1
        elif ptype == 'reference':
            fix_reference_paragraph(para)
            stats['reference'] += 1
        elif ptype == 'ref_heading':
            fix_reference_heading(para)
            stats['ref_heading'] = stats.get('ref_heading', 0) + 1
        elif ptype == 'ack_heading':
            fix_ack_heading(para)
            stats['ack_heading'] += 1
        elif ptype == 'caption':
            fix_caption_paragraph(para)
            stats['caption'] += 1
        elif ptype == 'table_note':
            fix_table_note_paragraph(para)
            stats['table_note'] += 1
        elif ptype == 'formula':
            # Formula paragraphs - just fix alignment, keep content as-is
            set_para_alignment(para, 'both')
            set_para_indent(para, first_line='480', first_line_chars='200')
            set_para_spacing(para, before='0', after='0', line='400', line_rule='exact')
            stats['body'] += 1
        elif ptype == 'body':
            fix_body_paragraph(para)
            # Fix half-width punctuation to full-width in Chinese body text
            if fix_cn_punctuation(para):
                stats['punctuation'] += 1
            stats['body'] += 1

        # Also fix punctuation in abstract_cn content
        if ptype == 'abstract_cn':
            if fix_cn_punctuation(para):
                stats['punctuation'] += 1

    # Special handling: fix the thesis title (中文题目)
    # Look for a paragraph that is:
    # - Longer than 5 Chinese chars (typical thesis title)
    # - Not a heading style (style 1,2,3,4 or TOC)
    # - Not starting with a number (not a section heading)
    # - Contains Chinese characters (not English title)
    # - Appears before the first heading or abstract section
    found_abstract = False
    for para in body.findall(w('p')):
        text = get_all_text(para)
        if not text:
            continue
        if '摘要' in text or 'Abstract' in text:
            found_abstract = True
            break

    if not found_abstract:
        # Scan for title before any heading
        for para in body.findall(w('p')):
            text = get_all_text(para)
            if not text or len(text) < 5:
                continue
            ppr = para.find(w('pPr'))
            style_id = ''
            if ppr is not None:
                ps = ppr.find(w('pStyle'))
                if ps is not None:
                    style_id = ps.get(w('val'), '')
            if style_id in ('1', '2', '3', '4', 'TOC1', 'TOC2', 'Heading1', 'Heading2'):
                continue
            # Check if it looks like a Chinese thesis title
            has_cn = any('\u4e00' <= c <= '\u9fff' for c in text)
            if has_cn and len(text) < 50:
                # Not a heading (doesn't start with number), not a label
                if not re.match(r'^[\d（\[]', text.strip()) and text.strip() not in ('目录', '摘要', '关键词', '致谢', '附录'):
                    fix_title_paragraph(para)
                    stats['title'] += 1
                    break

    # Write back document.xml
    tree.write(doc_xml_path, xml_declaration=True, encoding='UTF-8')

    # Fix TOC styles in styles.xml: 宋体 + TNR + 五号
    toc_fixed = fix_toc_styles(styles_path)
    stats['toc_styles'] = toc_fixed

    # Set updateFields in settings.xml so Word auto-updates TOC on open
    settings_path = os.path.join(unpacked_dir, 'word', 'settings.xml')
    if os.path.exists(settings_path):
        stree = ET.parse(settings_path)
        sroot = stree.getroot()
        uf = sroot.find(w('updateFields'))
        if uf is None:
            uf = ET.SubElement(sroot, w('updateFields'))
        uf.set(w('val'), 'true')
        stree.write(settings_path, xml_declaration=True, encoding='UTF-8')
        print("Settings: updateFields enabled (TOC will auto-update on open)")

    if output_docx:
        pack_docx(unpacked_dir, output_docx)
        print(f"Output written: {output_docx}")

    print("=== Fix Summary ===")
    print(f"Page setup sections fixed: {stats['page_setup']}")
    print(f"Title (CN) paragraphs fixed: {stats['title']}")
    print(f"Title (EN) paragraphs fixed: {stats['en_title']}")
    print(f"Level 1 headings fixed: {stats['h1']}")
    print(f"Level 2 headings fixed: {stats['h2']}")
    print(f"Level 3 headings fixed: {stats['h3']}")
    print(f"Level 4 headings fixed: {stats['h4']}")
    print(f"Abstract/Keywords fixed: {stats['abstract']}")
    print(f"EN Abstract/Keywords fixed: {stats['en_abstract']}")
    print(f"Body paragraphs fixed: {stats['body']}")
    print(f"Reference entries fixed: {stats['reference']}")
    print(f"Reference heading fixed: {stats['ref_heading']}")
    print(f"Acknowledgement heading fixed: {stats['ack_heading']}")
    print(f"Caption paragraphs fixed: {stats['caption']}")
    print(f"Table note paragraphs fixed: {stats['table_note']}")
    print(f"TOC headings fixed: {stats['toc']}")
    print(f"Punctuation fixes applied in: {stats['punctuation']} additional paragraphs")
    print("Done!")

    if temp_dir:
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == '__main__':
    main()
