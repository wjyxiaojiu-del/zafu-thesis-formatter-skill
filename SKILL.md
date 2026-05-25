---
name: zafu-thesis-formatter
description: >
  Format a Zhejiang A&F University undergraduate thesis .docx file.
  Fixes fonts, headings, spacing, indentation, figure/table numbering,
  and checks cross-references. Outputs an editable .docx file.
---

# ZAFU Thesis Formatter

Use this skill when a user asks to format, fix, or check a Zhejiang A&F University undergraduate thesis (.docx).

## Quick Start

```bash
# 从参考文档学习格式规则
python scripts/thesis_format.py learn 参考论文.docx rules.json

# 用规则修格式（输出 _formatted.docx）
python scripts/thesis_format.py apply 待修论文.docx rules.json

# 检查格式差异（不修改）
python scripts/thesis_format.py check 待修论文.docx rules.json
```

## Workflow

1. **learn**：从一篇格式正确的参考论文中提取格式规则（字体、字号、行距、缩进、对齐等），保存为 `rules.json`
2. **apply**：将规则应用到其他论文，自动修复：
   - 页面设置（A4、2.7cm 边距）
   - 目录标题和条目字体
   - 论文标题（保持原样）
   - 一/二/三/四级标题字体（楷体/黑体/宋体 + 字号 + 加粗）
   - 正文格式（宋体五号 + TNR + 首行缩进 + 两端对齐）
   - 摘要/关键词（标签黑体加粗，内容楷体）
   - 图/表按章节自动编号（图 X-Y、表 X-Y）
   - p<0.05 等统计标记自动斜体
   - TOC 样式修复（宋体 + TNR 五号）
3. **check**：对比文档与规则的差异，列出所有不符之处

## Format Rules

详见 [FORMAT_SPEC.md](FORMAT_SPEC.md)。

## Key Behaviors

- **跳过封面/声明**：只处理目录之后的内容
- **论文标题保留**：居中大字号段落不会被修改
- **标题识别**：支持 Word Heading 样式 + 编号模式（`1 标题`、`1.1 标题`、`1.1.1 标题`），有无空格都能识别
- **官方规则优先**：标题字体以 FORMAT_SPEC.md 为准，不完全依赖参考文档
- **摘要拆分字体**："摘要："用黑体加粗，内容用楷体
- **不强制刷新域**：避免损坏的域代码报错

## Requirements

- Python 3.8+
- `pip install python-docx`

## Limitations

- 封面格式不处理（保持原样）
- 公式内容不修改
- 三线表边框需人工确认
- 交叉引用域（REF 域）不自动创建，只检查引用编号是否匹配
