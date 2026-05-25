"""Split thesis markdown into chapter LaTeX files."""
import re, subprocess, sys
from pathlib import Path

md_file = Path(r"C:\Users\wangjunyi\Desktop\毕业论文\thesis.md")
build_dir = Path(r"C:\Users\wangjunyi\Desktop\毕业论文\build")
chapters_dir = build_dir / "chapters"
chapters_dir.mkdir(exist_ok=True)

# Read full markdown
md_text = md_file.read_text(encoding="utf-8")

# Find body start (first # heading)
body_match = re.search(r'^# .+', md_text, re.MULTILINE)
if not body_match:
    print("No headings found!")
    sys.exit(1)

body_md = md_text[body_match.start():]

# Convert to LaTeX via pandoc
result = subprocess.run(
    ["pandoc", "-f", "markdown+pipe_tables+grid_tables", "-t", "latex", "--wrap=none"],
    input=body_md, capture_output=True, encoding="utf-8", check=True
)
body_latex = result.stdout

# Shift heading levels: \section → \chapter, \subsection → \section, \subsubsection → \subsection
body_latex = re.sub(r'\\section\{', r'\\chapter{', body_latex)
body_latex = re.sub(r'\\subsection\{', r'\\section{', body_latex)
body_latex = re.sub(r'\\subsubsection\{', r'\\subsection{', body_latex)

# Fix image paths: images/media/ → images/
body_latex = body_latex.replace('images/media/', 'images/')

# Split by \chapter{} into separate files
parts = re.split(r'(\\chapter\{[^}]*\})', body_latex)

chapter_names = {}
idx = 0
current = ""
for part in parts:
    if part.startswith("\\chapter{"):
        if current.strip():
            idx += 1
            fname = f"chapter{idx}.tex"
            (chapters_dir / fname).write_text(current.strip(), encoding="utf-8")
            title = re.search(r'\\chapter\{(.+?)\}', current)
            if title:
                chapter_names[idx] = title.group(1)
        current = part
    else:
        current += part

if current.strip():
    idx += 1
    # Check if this is references or acknowledgments
    title_match = re.search(r'\\chapter\{(.+?)\}', current)
    title = title_match.group(1) if title_match else ""

    if "参考文献" in title or "致谢" in title:
        fname = "references.tex" if "参考文献" in title else "acknowledgments.tex"
    else:
        fname = f"chapter{idx}.tex"
    (chapters_dir / fname).write_text(current.strip(), encoding="utf-8")
    chapter_names[idx] = title

print(f"Generated {idx} chapter files:")
for i, name in chapter_names.items():
    print(f"  chapter{i}: {name}")
