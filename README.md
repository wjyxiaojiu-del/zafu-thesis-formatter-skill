# ZAFU Thesis Formatter-Skill

浙江农林大学本科生毕业论文（设计）格式自动排版工具。

基于学校 2022 版撰写规范，自动修正 Word 文档（.docx）中的字体、字号、行距、缩进、对齐、标点等格式问题。

## 功能

- **页面设置**：A4 纸、2.7cm 四边距、1.8cm 页眉、1.85cm 页脚
- **字体统一**：中文按规范使用对应中文字体（宋体/黑体/楷体），英文/数字统一 Times New Roman
- **标题格式**：一级（楷体四号加粗居中）、二级（黑体小四加粗）、三级（黑体五号）、四级（宋体五号）
- **正文格式**：宋体五号、两端对齐、首行缩进2字符、固定行距20磅
- **摘要/关键词**：中文摘要（黑体标签+楷体内容）、英文摘要（Times New Roman）
- **参考文献**：宋体五号、悬挂缩进2字符、左对齐
- **标点修正**：正文和中文摘要中标点自动转全角，Key Words 用英文半角逗号
- **目录样式**：宋体五号+Times New Roman、固定行距20磅
- **标题编号间距**：编号后自动替换为中文全角空格（U+3000）
- **域自动更新**：打开文档时自动更新目录等域

## 快速开始

### 环境要求

- Python 3.8+

### 安装依赖

```bash
pip install python-docx
```

### 使用方法

```bash
python scripts/fix_zafu_thesis.py <解包后的目录路径>
```

### 完整工作流

1. **解包** .docx 文件（本质是 ZIP）：
   ```bash
   python -c "
   import zipfile, os, sys
   src = sys.argv[1]
   dst = sys.argv[2]
   with zipfile.ZipFile(src, 'r') as z:
       z.extractall(dst)
   " "input.docx" "unpacked/"
   ```

2. **运行格式修复**：
   ```bash
   python scripts/fix_zafu_thesis.py unpacked/
   ```

3. **重新打包** 为 .docx：
   ```bash
   python -c "
   import zipfile, os, sys
   src_dir = sys.argv[1]
   out = sys.argv[2]
   with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
       for root, dirs, files in os.walk(src_dir):
           for f in files:
               fpath = os.path.join(root, f)
               arcname = os.path.relpath(fpath, src_dir)
               z.write(fpath, arcname)
   " "unpacked/" "output.docx"
   ```

## 格式规范

详见 [FORMAT_SPEC.md](FORMAT_SPEC.md)。

## 局限性（已知缺点）

1. **三线表无法处理**：Word 中三线表的格式（顶线、底线粗，栏目线细）需要手动在 Word 中调整，脚本无法自动识别和修改表格边框样式。
2. **图片插入/排版无法处理**：图片的位置、大小、环绕方式（嵌入型）需要手动操作，脚本只能设置题注文字的格式。
3. **题注无法生成**：表序/表题、图序/图名的题注文字需要手动输入，脚本只能修正已有题注的字体字号。
4. **目录内容依赖 Word 域更新**：脚本修改标题格式后，目录的文本和页码需要 Word 打开时自动更新（已设置 `updateFields`），或在 Word 中右键目录手动"更新域"。首次使用建议在 Word 中验证目录是否正确刷新。
5. **页眉页脚分节需要手动确认**：脚本不修改分节符和页眉页脚的链接关系，这部分需要确保文档已有正确的分节结构（封面→摘要→目录→正文各节）。
6. **公式保留原格式**：含有 Office Math 公式的段落会被跳过字体设置，以避免公式渲染异常。
7. **段落分类依赖文本模式匹配**：标题识别基于正则表达式（如 `^1.1`、`^2.3.1`），如果论文编号格式不规范（如使用全角数字、特殊符号等），可能导致误分类。
8. **标点替换非 100% 准确**：句号（`.`）的半角转全角采用启发式规则（判断前后字符是否为中文），在极少数边界情况（如"图1.2.3"）下可能误替换，需人工复核。
9. **仅适用于浙江农林大学**：格式规范硬编码为浙农林大 2022 版要求，其他学校或其他年份的规范可能不适用。

## 项目结构

```
zafu-thesis-formatter/
├── scripts/
│   └── fix_zafu_thesis.py    # 核心格式修复脚本
├── FORMAT_SPEC.md             # 完整格式规范文档
├── requirements.txt           # Python 依赖
├── LICENSE                    # MIT 许可证
└── README.md                  # 项目说明
```

## 许可证

MIT License
