# ZAFU Thesis Formatter 3.0

浙江农林大学本科毕业论文自动排版工具 — **LaTeX 驱动，格式 100% 确定性。**

## 核心思路

Word 改格式 → 痛苦、不确定、每次都要调。
LaTeX 排版 → 写一次模板，所有人永久受益。

## 快速开始

```bash
# Word 转 PDF（最常用）
python scripts/thesis_formatter.py 你的论文.docx --university zafu

# Markdown 转 PDF
python scripts/thesis_formatter.py 你的论文.md --university zafu

# 直接编译 LaTeX
python scripts/thesis_formatter.py 你的论文.tex --university zafu
```

## 工作流

```
Word/PDF/Markdown
      │
      ▼
  pandoc 转换 ──→ LaTeX 源文件
      │
      ▼
  注入 ZAFU 模板（封面/摘要/目录/正文/参考文献）
      │
      ▼
  XeLaTeX 编译（2 遍）
      │
      ▼
  完美格式 PDF ✅
```

## 环境要求

- Python 3.8+
- MiKTeX 或 TeX Live（含 XeLaTeX）
- pandoc 3.0+

### 安装缺失的 LaTeX 包（MiKTeX）

```bash
mpm --install zhnumber amscls natbib multirow subfig caption gbt7714
```

## 项目结构

```
zafu-thesis-formatter-skill/
├── SKILL.md                 # Claude Code skill 定义
├── FORMAT_SPEC.md           # ZAFU 格式规范原文
├── scripts/
│   ├── thesis_formatter.py  # 主转换脚本
│   └── fix_zafu_thesis.py   # Word 修复脚本（旧版）
├── templates/
│   └── zafu/
│       ├── main.tex         # 主文件
│       ├── zafu.cls         # 排版样式定义
│       ├── cover.tex        # 封面
│       ├── abstract-zh.tex  # 中文摘要
│       ├── abstract-en.tex  # 英文摘要
│       ├── acknowledgments.tex
│       ├── references.bib   # 参考文献
│       └── chapters/        # 正文章节
├── requirements.txt
├── LICENSE
└── README.md
```

## 格式规范

详见 [FORMAT_SPEC.md](FORMAT_SPEC.md)。

## 已知限制

1. Word → LaTeX 转换依赖 pandoc，复杂格式可能需要人工微调
2. 公式环境会尽量保留，但复杂 Office Math 可能需要手动修复
3. 图片自动提取，但浮动体位置可能需要调整
4. 参考文献格式按 GB/T 7714，但内容正确性需人工校对

## 扩展到其他学校

1. 在 `templates/` 下新建学校目录（如 `templates/zju/`）
2. 创建对应的 `.cls` 文件定义排版规则
3. 修改 `thesis_formatter.py` 的模板映射

## 许可证

MIT License
