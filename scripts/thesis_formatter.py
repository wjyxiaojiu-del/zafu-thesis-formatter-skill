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
    subprocess.run(cmd, check=True, capture_output=True, text=True)
    return md_path


def extract_metadata(md_text: str) -> dict:
    """Try to extract thesis metadata from markdown content."""
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

    # Try to find Chinese title (usually first heading or bold text)
    title_match = re.search(r'^#\s+(.+)$', md_text, re.MULTILINE)
    if title_match:
        meta["title_zh"] = title_match.group(1).strip()

    # Try to find abstract sections
    zh_abs = re.search(
        r'(?:摘\s*要|摘　要)[：:\s]*(.*?)(?=关键词|Keywords|Abstract|\Z)',
        md_text, re.DOTALL | re.IGNORECASE
    )
    if zh_abs:
        meta["abstract_zh"] = zh_abs.group(1).strip()

    zh_kw = re.search(
        r'关键词[：:\s]*(.*?)(?=\n\n|\n#|\Z)',
        md_text, re.DOTALL
    )
    if zh_kw:
        meta["abstract_zh_keywords"] = zh_kw.group(1).strip()

    en_abs = re.search(
        r'Abstract[：:\s]*(.*?)(?=Key\s*words|Keywords|\Z)',
        md_text, re.DOTALL | re.IGNORECASE
    )
    if en_abs:
        meta["abstract_en"] = en_abs.group(1).strip()

    en_kw = re.search(
        r'Key\s*words[：:\s]*(.*?)(?=\n\n|\n#|\Z)',
        md_text, re.DOTALL | re.IGNORECASE
    )
    if en_kw:
        meta["abstract_en_keywords"] = en_kw.group(1).strip()

    return meta


def remove_metadata_sections(md_text: str) -> str:
    """Remove abstract/keyword sections from body (they go into separate files)."""
    # Remove everything before the first real chapter heading
    # Keep content from first ## or # 绪论 / # 第一章 etc.
    patterns = [
        r'^#\s*(?:摘\s*要|摘　要).*?(?=^#\s)',
        r'^#\s*Abstract.*?(?=^#\s)',
    ]
    for p in patterns:
        md_text = re.sub(p, '', md_text, flags=re.MULTILINE | re.DOTALL | re.IGNORECASE)

    # Remove standalone keyword lines
    md_text = re.sub(r'^关键词[：:].*$', '', md_text, flags=re.MULTILINE)
    md_text = re.sub(r'^Key\s*words[：:].*$', '', md_text, flags=re.MULTILINE | re.IGNORECASE)

    return md_text.strip()


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
    result = subprocess.run(cmd, capture_output=True, text=True, check=True)
    return result.stdout


def fix_latex_headings(latex_text: str) -> str:
    """Convert pandoc headings to ctex chapter/section commands."""
    # pandoc outputs: \chapter{...}, \section{...}, etc. with numbered chapters
    # Ensure consistent numbering style
    return latex_text


def fix_image_paths(latex_text: str, work_dir: Path) -> str:
    """Fix image paths in LaTeX to point to extracted images."""
    images_dir = work_dir / "images"
    if images_dir.exists():
        # Make paths relative to the tex file
        latex_text = latex_text.replace(
            str(images_dir) + os.sep,
            "images/"
        )
        # Also handle Windows backslashes
        latex_text = latex_text.replace(
            str(images_dir).replace("\\", "/") + "/",
            "images/"
        )
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
        "导师姓名\\hspace{1\\ccwd}教授": meta.get("advisor", "导师"),
    }
    for old, new in replacements.items():
        content = content.replace(old, new)
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

    # Write body content
    chapters_dir = thesis_dir / "chapters"
    chapters_dir.mkdir(exist_ok=True)

    # Split by \chapter{} into separate files
    parts = re.split(r'(\\chapter\{[^}]*\})', body_latex)
    chapter_idx = 0
    chapter_content = ""

    for part in parts:
        if part.startswith("\\chapter{"):
            # Save previous chapter
            if chapter_content.strip():
                chapter_idx += 1
                ch_file = chapters_dir / f"chapter{chapter_idx}.tex"
                ch_file.write_text(chapter_content, encoding="utf-8")
            chapter_content = part
        else:
            chapter_content += part

    # Save last chapter
    if chapter_content.strip():
        chapter_idx += 1
        ch_file = chapters_dir / f"chapter{chapter_idx}.tex"
        ch_file.write_text(chapter_content, encoding="utf-8")

    # Update main.tex to include correct number of chapters
    main_content = main_tex.read_text(encoding="utf-8")
    # Remove default chapter includes
    main_content = re.sub(
        r'\\input\{chapters/chapter\d+\}\n?',
        '', main_content
    )
    # Add correct includes before bibliography
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
    if not bib_file.exists():
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
            cmd, capture_output=True, text=True,
            cwd=str(thesis_dir), timeout=120
        )
        if result.returncode != 0 and pass_num == 1:
            # Show error on second pass failure
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
        # Step 1: Convert to markdown (if needed)
        if fmt == "docx":
            print("Step 1/4: Converting Word to Markdown...")
            md_path = docx_to_markdown(str(input_path), work_dir)
            md_text = md_path.read_text(encoding="utf-8")
        elif fmt == "markdown":
            md_text = input_path.read_text(encoding="utf-8")
        elif fmt == "tex":
            # Direct LaTeX - just compile
            print("Compiling LaTeX directly...")
            pdf = build_thesis(work_dir, template_dir, {}, input_path.read_text(encoding="utf-8"))
            output = args.output or str(input_path.with_suffix(".pdf"))
            shutil.copy2(str(pdf), output)
            print(f"Done: {output}")
            return

        # Step 2: Extract metadata
        print("Step 2/4: Extracting metadata...")
        meta = extract_metadata(md_text)

        # CLI args override extracted metadata
        if args.title_zh: meta["title_zh"] = args.title_zh
        if args.title_en: meta["title_en"] = args.title_en
        if args.author: meta["author"] = args.author
        if args.student_id: meta["student_id"] = args.student_id
        if args.college: meta["college"] = args.college
        if args.major: meta["major"] = args.major
        if args.advisor: meta["advisor"] = args.advisor

        print(f"  Title: {meta['title_zh'] or '(not detected - use --title-zh)'}")

        # Step 3: Convert body to LaTeX
        print("Step 3/4: Converting to LaTeX...")
        body_md = remove_metadata_sections(md_text)
        body_latex = markdown_to_latex(body_md, work_dir)
        body_latex = fix_latex_headings(body_latex)
        body_latex = fix_image_paths(body_latex, work_dir)

        # Step 4: Compile PDF
        print("Step 4/4: Compiling PDF...")
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
