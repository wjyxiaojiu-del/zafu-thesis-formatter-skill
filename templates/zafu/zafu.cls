%% zafu.cls — 浙江农林大学本科毕业论文 LaTeX 模板
%% 基于《浙江农林大学本科生毕业论文（设计）系列材料（2022版）》
\NeedsTeXFormat{LaTeX2e}
\ProvidesClass{zafu}[2026/05/25 ZAFU Undergraduate Thesis Class]

% ---------- 基础选项 ----------
\LoadClass[a4paper, zihao=-4, UTF8, openany]{ctexbook}

% ---------- 页面设置 ----------
\RequirePackage[
  top=2.7cm, bottom=2.7cm, left=2.7cm, right=2.7cm,
  headheight=14pt, headsep=0.6cm,
  footskip=0.85cm
]{geometry}

% ---------- 核心宏包 ----------
\RequirePackage{etoolbox}        % 补丁命令
\RequirePackage{fancyhdr}        % 页眉页脚
\RequirePackage{titletoc}        % 目录格式
\RequirePackage{graphicx}        % 图片
\RequirePackage{booktabs}        % 三线表
\RequirePackage{multirow}        % 合并行
\RequirePackage{tabularx}        % 自动宽度表格
\RequirePackage{longtable}       % 跨页长表
\RequirePackage[caption=false]{subfig}  % 子图
\RequirePackage{amsmath,amssymb,amsthm}  % 数学
\RequirePackage{hyperref}        % 超链接
\RequirePackage[sort&compress]{gbt7714}  % GB/T 7714 参考文献
\RequirePackage{enumitem}        % 列表环境
\RequirePackage{setspace}        % 行距
\RequirePackage{caption}         % 图表标题
\RequirePackage{float}           % 浮动体控制

% ---------- 中文字体设置 ----------
% ctexbook 已自动配置宋体/黑体/仿宋，这里只补楷体和英文字体
\newCJKfontfamily\kaiti{KaiTi}

% 英文字体
\setmainfont{Times New Roman}

% ---------- 字号定义 ----------
\newcommand{\zihaoerhao}{\fontsize{22pt}{26.4pt}\selectfont}       % 二号
\newcommand{\zihaosanhao}{\fontsize{16pt}{19.2pt}\selectfont}      % 三号
\newcommand{\zihaosansanhao}{\fontsize{15pt}{18pt}\selectfont}     % 小三
\newcommand{\zihaosihao}{\fontsize{14pt}{16.8pt}\selectfont}       % 四号
\newcommand{\zihaoxiaosihao}{\fontsize{12pt}{14.4pt}\selectfont}   % 小四
\newcommand{\zihaowuhao}{\fontsize{10.5pt}{12.6pt}\selectfont}     % 五号
\newcommand{\zihaoxiaowuhao}{\fontsize{9pt}{10.8pt}\selectfont}    % 小五

% ---------- 行距设置 ----------
\linespread{1}
\setlength{\parskip}{0pt}

% 正文固定行距 20pt
\newcommand{\bodytextstyle}{%
  \songti\zihaowuhao
  \setlength{\baselineskip}{20pt}%
  \setlength{\parindent}{2\ccwd}%
}

% ---------- 一级标题：楷体四号加粗居中 ----------
\ctexset{
  chapter = {
    format = \centering\heiti\zihaosihao\bfseries,
    nameformat = {},
    titleformat = {},
    number = \arabic{chapter},
    name = {,},
    aftername = \hspace{1\ccwd},
    beforeskip = 6pt,
    afterskip = 6pt,
    pagestyle = fancy,
  },
  section = {
    format = \raggedright\heiti\zihaoxiaosihao\bfseries,
    aftername = \hspace{1\ccwd},
    beforeskip = 3pt,
    afterskip = 3pt,
    indent = 2\ccwd,
  },
  subsection = {
    format = \raggedright\heiti\zihaowuhao,
    aftername = \hspace{1\ccwd},
    beforeskip = 3pt,
    afterskip = 3pt,
    indent = 2\ccwd,
  },
  subsubsection = {
    format = \raggedright\songti\zihaowuhao,
    aftername = \hspace{1\ccwd},
    beforeskip = 3pt,
    afterskip = 3pt,
    indent = 2\ccwd,
  },
}

% ---------- 页眉页脚 ----------
\pagestyle{fancy}
\fancyhf{}
\fancyhead[C]{\songti\zihaoxiaowuhao 浙江农林大学本科生毕业论文}
\fancyfoot[C]{\thepage}
\renewcommand{\headrulewidth}{0.4pt}
\renewcommand{\footrulewidth}{0pt}

% 首页不显示页眉
\fancypagestyle{plain}{
  \fancyhf{}
  \fancyfoot[C]{\thepage}
  \renewcommand{\headrulewidth}{0pt}
}

% ---------- 目录格式 ----------
\renewcommand\contentsname{目\hspace{2\ccwd}录}

\titlecontents{chapter}[0pt]
  {\songti\zihaowuhao\vspace{6pt}}
  {\thecontentslabel\hspace{1\ccwd}}
  {}
  {\hspace{.5em}\titlerule*[5pt]{.}\contentspage}

\titlecontents{section}[2\ccwd]
  {\songti\zihaowuhao}
  {\thecontentslabel\hspace{1\ccwd}}
  {}
  {\hspace{.5em}\titlerule*[5pt]{.}\contentspage}

\titlecontents{subsection}[4\ccwd]
  {\songti\zihaowuhao}
  {\thecontentslabel\hspace{1\ccwd}}
  {}
  {\hspace{.5em}\titlerule*[5pt]{.}\contentspage}

% ---------- 图表标题格式 ----------
\captionsetup{
  font = {small},           % 小五号
  labelfont = {bf},
  labelsep = space,
  justification = centering,
}
\captionsetup[table]{position = top}
\captionsetup[figure]{position = bottom}

% 图表编号格式
\renewcommand{\thefigure}{\thechapter-\arabic{figure}}
\renewcommand{\thetable}{\thechapter-\arabic{table}}
\renewcommand{\theequation}{\thechapter-\arabic{equation}}

% ---------- 参考文献 ----------
\renewcommand{\bibname}{参考文献}
\setlength{\bibsep}{0pt}

% 参考文献标题样式：楷体四号加粗居中
\patchcmd{\thebibliography}{\section*}{\chapter*}{}{}

% ---------- 超链接设置 ----------
\hypersetup{
  colorlinks = true,
  linkcolor = black,
  citecolor = black,
  urlcolor = blue,
  bookmarksnumbered = true,
  pdfstartview = FitH,
}

% ---------- 用户可设置的元数据命令 ----------
\newcommand{\thesistitlezh}[1]{\def\zafutitlezh{#1}}
\newcommand{\thesistitleen}[1]{\def\zafutitleen{#1}}
\newcommand{\studentid}[1]{\def\zafustudentid{#1}}
\newcommand{\college}[1]{\def\zafucollege{#1}}
\newcommand{\major}[1]{\def\zafumajor{#1}}
\newcommand{\advisor}[1]{\def\zafuadvisor{#1}}

% Aliases for standard LaTeX \author and \date
\let\zafuauthor\@author
\let\zafudate\@date

\newcommand{\zafutitlezh}{论文题目}
\newcommand{\zafutitleen}{Thesis Title}
\newcommand{\zafustudentid}{}
\newcommand{\zafucollege}{}
\newcommand{\zafumajor}{}
\newcommand{\zafuadvisor}{}

% ---------- 中文摘要环境 ----------
\newenvironment{abstractzh}{%
  \chapter*{摘\hspace{2\ccwd}要}
  \addcontentsline{toc}{chapter}{摘要}
  \kaiti\zihaowuhao
  \setlength{\baselineskip}{20pt}
  \setlength{\parindent}{2\ccwd}
}{%
  \par\vspace{12pt}
}

\newcommand{\keywordszh}[1]{%
  \par\vspace{6pt}\noindent
  \heiti\zihaowuhao\bfseries 关键词：\kaiti #1
}

% ---------- 英文摘要环境 ----------
\newenvironment{abstracten}{%
  \chapter*{Abstract}
  \addcontentsline{toc}{chapter}{Abstract}
  \zihaowuhao
  \setlength{\baselineskip}{20pt}
  \setlength{\parindent}{2\ccwd}
}{%
  \par\vspace{12pt}
}

\newcommand{\keywordsen}[1]{%
  \par\vspace{6pt}\noindent
  \textbf{Key words:} #1
}

% ---------- 致谢环境 ----------
\newenvironment{acknowledgments}{%
  \chapter*{致\hspace{2\ccwd}谢}
  \addcontentsline{toc}{chapter}{致谢}
  \bodytextstyle
}{%
  \par
}

% ---------- 附录环境 ----------
\newcommand{\appendixsection}[1]{%
  \section*{#1}
  \addcontentsline{toc}{section}{#1}
}

\endinput
