#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
论文格式化工具 — Word/Markdown → LaTeX → PDF
用法:
  python thesis_formatter.py input.docx --university zafu
  python thesis_formatter.py input.md --university zafu
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
TEMPLATE_DIR = SCRIPT_DIR.parent / "templates"


def detect_input_format(path: str) -> str:
    ext = Path(path).suffix.lower()
    if ext == ".docx":
        return "docx"
    if ext in (".md", ".markdown"):
        return "markdown"
    if ext == ".tex":
        return "tex"
    raise ValueError(f"Unsupported format: {ext}")


def docx_to_markdown(docx_path: str, work_dir: Path) -> Path:
    """Convert .docx to .md via pandoc, extracting images."""
    md_path = work_dir / "content.md"
    media_dir = work_dir / "images"
    cmd = [
        "pandoc", docx_path,
        "-t", "markdown+tex_math_dollars+pipe_tables+grid_tables",
        "--wrap=none",
        f"--extract-media={media_dir}",
        "-o", str(md_path),
    ]
    subprocess.run(cmd, check=True, capture_output=True, encoding="utf-8")
    return md_path


def strip_bold(text: str) -> str:
    """Remove **bold** markers from text."""
    return re.sub(r'\*\*(.*?)\*\*', r'\1', text)


def extract_metadata(md_text: str) -> dict:
    """Extract thesis metadata from markdown content."""
    meta = {
        "title_zh": "",
        "title_en": "",
        "author": "",
        "student_id": "",
        "college": "",
        "major": "",
        "advisor": "",
        "abstract_zh": "",
        "abstract_zh_keywords": "",
        "abstract_en": "",
        "abstract_en_keywords": "",
    }

    # Title: first non-empty line before any heading
    lines = md_text.strip().split("\n")
    for line in lines:
        stripped = line.strip().lstrip("*").strip()
        if stripped and not stripped.startswith("#") and not stripped.startswith(">"):
            meta["title_zh"] = strip_bold(stripped).strip()
            break

    # Chinese abstract: between 摘要 and 关键词
    zh_abs = re.search(
        r'\*?\*?摘\s*要\*?\*?[：:\s]+(.*?)(?=\*?\*?关键词\*?\*?|Key\s*words|\n#|\Z)',
        md_text, re.DOTALL | re.IGNORECASE
    )
    if zh_abs:
        meta["abstract_zh"] = strip_bold(zh_abs.group(1).strip())

    zh_kw = re.search(
        r'\*?\*?关键词\*?\*?[：:\s]+(.*?)(?=\n\n|\n#|\Z)',
        md_text, re.DOTALL
    )
    if zh_kw:
        meta["abstract_zh_keywords"] = strip_bold(zh_kw.group(1).strip())

    # English abstract
    en_abs = re.search(
        r'\*?\*?Abstract\*?\*?[：:\s]+(.*?)(?=\*?\*?Key\s*words\*?\*?|\n#|\Z)',
        md_text, re.DOTALL | re.IGNORECASE
    )
    if en_abs:
        meta["abstract_en"] = strip_bold(en_abs.group(1).strip())

    en_kw = re.search(
        r'\*?\*?Key\s*words\*?\*?[：:\s]+(.*?)(?=\n\n|\n#|\Z)',
        md_text, re.DOTALL | re.IGNORECASE
    )
    if en_kw:
        meta["abstract_en_keywords"] = strip_bold(en_kw.group(1).strip())

    return meta


def get_body_after_toc(md_text: str) -> str:
    """Extract body content: everything from the first # heading onward."""
    # Find first markdown heading (# ...)
    match = re.search(r'^#\s+.+', md_text, re.MULTILINE)
    if match:
        return md_text[match.start():].strip()
    # Fallback: skip first few lines (title, abstract, etc.)
    return md_text.strip()


def remove_toc_section(md_text: str) -> str:
    """Remove the TOC blockquote section that pandoc generates."""
    # Remove blockquote TOC (lines starting with >)
    lines = md_text.split("\n")
    result = []
    in_toc = False
    for line in lines:
        if line.strip().startswith("> [") or line.strip().startswith(">\\"):
            in_toc = True
            continue
        if in_toc and (line.strip() == "" or line.strip() == ">"):
            in_toc = False
            continue
        if not in_toc:
            result.append(line)
    return "\n".join(result)


def markdown_to_latex(md_text: str, work_dir: Path) -> str:
    """Convert markdown body to LaTeX via pandoc."""
    md_file = work_dir / "body.md"
    md_file.write_text(md_text, encoding="utf-8")

    cmd = [
        "pandoc", str(md_file),
        "-f", "markdown+tex_math_dollars+pipe_tables",
        "-t", "latex",
        "--wrap=none",
    ]
    result = subprocess.run(cmd, capture_output=True, encoding="utf-8", check=True)
    return result.stdout


def fix_latex_headings(latex_text: str) -> str:
    """Shift pandoc heading levels for ctexbook: section→chapter, subsection→section, etc."""
    # Must process from deepest to shallowest to avoid double-replacement
    latex_text = re.sub(r'\\subsubsection\{', r'\\subsubsection*{', latex_text)
    latex_text = re.sub(r'\\subsection\{', r'\\subsubsection{', latex_text)
    latex_text = re.sub(r'\\section\{', r'\\chapter{', latex_text)
    # Also fix \label references that pandoc adds
    latex_text = re.sub(r'\\subsubsection\*\{', r'\\subsubsection{', latex_text)
    return latex_text


def fix_image_paths(latex_text: str, work_dir: Path) -> str:
    """Fix image paths in LaTeX to point to extracted images."""
    if not latex_text:
        return ""
    images_dir = work_dir / "images"
    if images_dir.exists():
        # Normalize path separators
        dir_str = str(images_dir).replace("\\", "/")
        latex_text = latex_text.replace(dir_str + "/", "images/")
        latex_text = latex_text.replace(str(images_dir) + os.sep, "images/")
    return latex_text


def build_thesis(
    work_dir: Path,
    template_dir: Path,
    meta: dict,
    body_latex: str,
) -> Path:
    """Assemble all parts and compile to PDF."""
    thesis_dir = work_dir / "thesis"
    thesis_dir.mkdir(exist_ok=True)

    # Copy template files
    for f in template_dir.glob("*.tex"):
        shutil.copy2(f, thesis_dir / f.name)
    for f in template_dir.glob("*.cls"):
        shutil.copy2(f, thesis_dir / f.name)
    for f in template_dir.glob("*.bib"):
        shutil.copy2(f, thesis_dir / f.name)

    # Copy images
    src_images = work_dir / "images"
    if src_images.exists():
        dst_images = thesis_dir / "images"
        if dst_images.exists():
            shutil.rmtree(dst_images)
        shutil.copytree(src_images, dst_images)

    # Write metadata into main.tex
    main_tex = thesis_dir / "main.tex"
    content = main_tex.read_text(encoding="utf-8")
    replacements = {
        "论文中文题目": meta.get("title_zh", "论文题目"),
        "English Title of the Thesis": meta.get("title_en", "Thesis Title"),
        "作者姓名": meta.get("author", "作者"),
        "20200000000": meta.get("student_id", ""),
        "林学与生物技术学院": meta.get("college", "学院"),
        "林学": meta.get("major", "专业"),
    }
    for old, new in replacements.items():
        content = content.replace(old, new)
    # Handle advisor separately (contains LaTeX commands)
    advisor = meta.get("advisor", "")
    if advisor:
        content = content.replace(
            r"导师姓名\hspace{1\ccwd}教授",
            advisor
        )
    main_tex.write_text(content, encoding="utf-8")

    # Write Chinese abstract
    if meta.get("abstract_zh"):
        abs_zh = thesis_dir / "abstract-zh.tex"
        keywords = meta.get("abstract_zh_keywords", "")
        abs_zh.write_text(
            f"\\begin{{abstractzh}}\n{meta['abstract_zh']}\n\n"
            f"\\keywordszh{{{keywords}}}\n\\end{{abstractzh}}\n\\clearpage\n",
            encoding="utf-8"
        )

    # Write English abstract
    if meta.get("abstract_en"):
        abs_en = thesis_dir / "abstract-en.tex"
        keywords = meta.get("abstract_en_keywords", "")
        abs_en.write_text(
            f"\\begin{{abstracten}}\n{meta['abstract_en']}\n\n"
            f"\\keywordsen{{{keywords}}}\n\\end{{abstracten}}\n\\clearpage\n",
            encoding="utf-8"
        )

    # Write body content - split by \chapter{} into separate files
    chapters_dir = thesis_dir / "chapters"
    if chapters_dir.exists():
        shutil.rmtree(chapters_dir)
    chapters_dir.mkdir(exist_ok=True)

    # Remove existing chapter includes from main.tex
    main_content = main_tex.read_text(encoding="utf-8")
    main_content = re.sub(
        r'% ========== 正文 ==========\n.*?% ========== 参考文献 ==========',
        '% ========== 正文 ==========\n\n% ========== 参考文献 ==========',
        main_content, flags=re.DOTALL
    )

    # Split body by \chapter{}
    parts = re.split(r'(\\chapter\{[^}]*\})', body_latex)
    chapter_idx = 0
    chapter_content = ""

    for part in parts:
        if part.startswith("\\chapter{"):
            if chapter_content.strip():
                chapter_idx += 1
                ch_file = chapters_dir / f"chapter{chapter_idx}.tex"
                ch_file.write_text(chapter_content, encoding="utf-8")
            chapter_content = part
        else:
            chapter_content += part

    if chapter_content.strip():
        chapter_idx += 1
        ch_file = chapters_dir / f"chapter{chapter_idx}.tex"
        ch_file.write_text(chapter_content, encoding="utf-8")

    if chapter_idx == 0:
        # No \chapter{} found — write entire body as single chapter
        chapter_idx = 1
        ch_file = chapters_dir / "chapter1.tex"
        ch_file.write_text(body_latex, encoding="utf-8")

    # Insert chapter includes
    chapter_includes = "\n".join(
        f"\\input{{chapters/chapter{i}}}" for i in range(1, chapter_idx + 1)
    )
    main_content = main_content.replace(
        "% ========== 参考文献 ==========",
        f"{chapter_includes}\n\n% ========== 参考文献 =========="
    )
    main_tex.write_text(main_content, encoding="utf-8")

    # Create empty references.bib if not exists
    bib_file = thesis_dir / "references.bib"
    if not bib_file.exists() or bib_file.stat().st_size < 20:
        bib_file.write_text("% 在此添加参考文献\n", encoding="utf-8")

    # Compile with XeLaTeX (2 passes)
    for pass_num in range(2):
        cmd = [
            "xelatex",
            "-interaction=nonstopmode",
            "-halt-on-error",
            "main.tex",
        ]
        result = subprocess.run(
            cmd, capture_output=True, encoding="utf-8", errors="replace",
            cwd=str(thesis_dir), timeout=180
        )
        if result.returncode != 0 and pass_num == 1:
            log_file = thesis_dir / "main.log"
            if log_file.exists():
                log_text = log_file.read_text(encoding="utf-8", errors="replace")
                errors = [l for l in log_text.split("\n") if l.startswith("!")]
                if errors:
                    print("LaTeX errors:", "\n".join(errors[:10]), file=sys.stderr)

    pdf_path = thesis_dir / "main.pdf"
    if pdf_path.exists():
        return pdf_path
    else:
        raise RuntimeError("PDF compilation failed. Check LaTeX log for errors.")


def main():
    parser = argparse.ArgumentParser(description="论文格式化工具")
    parser.add_argument("input", help="输入文件路径 (.docx / .md / .tex)")
    parser.add_argument("--university", "-u", default="zafu", help="学校模板 (default: zafu)")
    parser.add_argument("--output", "-o", help="输出 PDF 路径")
    parser.add_argument("--keep-temp", action="store_true", help="保留临时文件")
    parser.add_argument("--title-zh", help="论文中文标题")
    parser.add_argument("--title-en", help="论文英文标题")
    parser.add_argument("--author", help="作者姓名")
    parser.add_argument("--student-id", help="学号")
    parser.add_argument("--college", help="学院")
    parser.add_argument("--major", help="专业")
    parser.add_argument("--advisor", help="指导教师")
    args = parser.parse_args()

    input_path = Path(args.input).resolve()
    if not input_path.exists():
        print(f"Error: File not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    template_dir = TEMPLATE_DIR / args.university
    if not template_dir.exists():
        print(f"Error: Template not found: {args.university}", file=sys.stderr)
        sys.exit(1)

    fmt = detect_input_format(str(input_path))
    work_dir = Path(tempfile.mkdtemp(prefix="thesis_"))

    try:
        # Step 1: Get markdown content
        if fmt == "docx":
            print("Step 1/4: Converting Word to Markdown...")
            md_path = docx_to_markdown(str(input_path), work_dir)
            md_text = md_path.read_text(encoding="utf-8")
        elif fmt == "markdown":
            md_text = input_path.read_text(encoding="utf-8")
        elif fmt == "tex":
            print("Compiling LaTeX directly...")
            pdf = build_thesis(work_dir, template_dir, {}, input_path.read_text(encoding="utf-8"))
            output = args.output or str(input_path.with_suffix(".pdf"))
            shutil.copy2(str(pdf), output)
            print(f"Done: {output}")
            return

        # Step 2: Extract metadata
        print("Step 2/4: Extracting metadata...")
        meta = extract_metadata(md_text)

        if args.title_zh: meta["title_zh"] = args.title_zh
        if args.title_en: meta["title_en"] = args.title_en
        if args.author: meta["author"] = args.author
        if args.student_id: meta["student_id"] = args.student_id
        if args.college: meta["college"] = args.college
        if args.major: meta["major"] = args.major
        if args.advisor: meta["advisor"] = args.advisor

        print(f"  Title: {meta['title_zh'] or '(not detected - use --title-zh)'}")
        print(f"  Abstract: {len(meta['abstract_zh'])} chars")

        # Step 3: Convert body to LaTeX
        print("Step 3/4: Converting to LaTeX...")
        body_md = get_body_after_toc(md_text)
        body_md = remove_toc_section(body_md)
        body_latex = markdown_to_latex(body_md, work_dir)
        body_latex = fix_latex_headings(body_latex)
        body_latex = fix_image_paths(body_latex, work_dir)

        # Step 4: Compile PDF
        print("Step 4/4: Compiling PDF (2 passes)...")
        pdf = build_thesis(work_dir, template_dir, meta, body_latex)

        output = args.output or str(input_path.with_name(
            f"{input_path.stem}_{args.university}.pdf"
        ))
        shutil.copy2(str(pdf), output)
        print(f"\nDone! Output: {output}")

    finally:
        if not args.keep_temp:
            shutil.rmtree(work_dir, ignore_errors=True)
        else:
            print(f"Temp files: {work_dir}")


if __name__ == "__main__":
    main()
