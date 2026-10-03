# 文档类型与入口

YibinThesis 的一个 `main.tex` 只生成一种文档。文档类型由 `\documentclass` 选项决定，
学科类型由 `humanities` 或 `science` 决定，两者可以组合使用。

| 文档类型 | 类选项 | `template-year` | 主要入口命令 | 版式来源 |
|---|---|---:|---|---|
| 毕业论文（设计） | `thesis`，可省略 | `2024` | `\makeyibincover`、`\makeyibindeclarations` | 2024 论文模板与校级撰写规范 |
| 开题报告 | `proposal` | `2022` | `\makeyibinproposal` | 2022 工作表物理第 5 至 7 页的开题报告模块 |
| 文献综述 | `literature-review` | `2024` | `\makeyibinliteraturereviewcover` | 简洁信息页加 2024 论文正文版式基线 |

`template-year` 必须与上表逐项对应。它登记当前文档实际采用的版式来源，不会自动
选择或混用其他年份；类型与年份不匹配时，LaTeX 和 Word 构建都会报错并停止。
当前不支持 2025 论文模板。

源码按职责拆分：`yibinthesis.cls` 提供三类文档共用的元数据、字体、章节、引用与
图表基础，开题报告和文献综述分别由 `yibinthesis-proposal.sty`、
`yibinthesis-literature-review.sty` 实现。使用者仍只需加载 `yibinthesis` 类，类会按
文档类型自动加载对应模块，不需要在 `main.tex` 中手动 `\usepackage`。

## 毕业论文

`thesis` 是默认类型，旧项目只写 `\documentclass[humanities]{yibinthesis}` 仍然有效。
其元数据使用 `template-year = 2024`。
显式入口可以写成：

```tex
\documentclass[thesis,humanities]{yibinthesis}

\input{metadata}
\addbibresource{references.bib}

\begin{document}
\makeyibincover
\makeyibindeclarations

\frontmatter
\input{chapters/00-abstract-cn}
\input{chapters/01-abstract-en}
\clearpage
\tableofcontents

\mainmatter
\input{chapters/10-introduction}
\input{chapters/20-body}
\input{chapters/30-conclusion}

\backmatter
\input{chapters/99-acknowledgements}
\printyibinnotes
\printyibinbibliography
\input{chapters/90-appendix}
\end{document}
```

删除 `\makeyibindeclarations` 可以不生成原创性声明和版权授权书。后置部分由入口文件
装配，不应为了调整章节顺序修改模板类。

## 开题报告

### 2024 模板核对结果

本项目已对 `01_规范模板/最新要求与模板/罗殷开题报告.docx` 做 OOXML 级只读拆解。
该文件与前一版“简单改造正文”的实现不是同一结构：

- 文件有 2 个 Word 分节。第 1 节是封面信息区，A4 页面使用上/下 1.0 英寸、左/右
  1.25 英寸的工作表边距；第 2 节是正式开题表，使用上/下 1.0 英寸、左/右 1.25
  英寸的表格页边界，并保留页脚距离约 1.75 cm。两节均保留宋体小四为主体字体，标题
  和标签使用黑体/楷体直接格式。
- 正式表单是一个连续的 7 行结构：选题意义、国内外研究现状概述、主要研究内容、拟采用
  的研究思路、研究工作安排及进度、参考文献目录、指导教师意见。行高采用可扩展的
  `atLeast` 语义，长内容需要自然跨页，不能用固定高度裁切。
- Word 样例中存在“可续页”提示和直接格式化的表格边框/字号；本项目的 PDF/Word 生成器
  已按“自然续页、不写可续页字样”的可交付规则处理，提示语不应进入用户正文。
- 页码从 4 开始。LaTeX `yibinthesis-proposal.sty` 与 Word 后处理都必须保持这一点，
  并确保正文表单不是被错误拆成七个独立页。

当前 `yibinthesis-proposal.sty` 和 `lib/build_word.py` 已实现独立 proposal profile、
七字段顺序校验、单表结构和从第 4 页开始的页码。仍需在装有 Word 或 LibreOffice 的
环境中做最终分页复核；本机当前没有 `soffice`，无法完成 PNG 级视觉验收。

开题报告只复刻 2022 工作表物理第 5 至 7 页的开题栏目，页码从 4 开始；不生成
工作表封面、任务书、中期检查表和指导记录表。模板不插入“可续页”字样，也不按
官方示例的三张物理页强制分页，实际页数由字段内容自然决定。入口必须先调用
`\makeyibinproposal`，随后按以下顺序且各一次填写七个字段：

开题项目的元数据必须使用 `template-year = 2022`。这里的 2022 是开题工作表版本，
不继承、也不冒充 2024 毕业论文模板。

| 字段键 | 对应栏目 |
|---|---|
| `significance` | 选题意义 |
| `research-status` | 国内外研究现状概述 |
| `research-content` | 主要研究内容 |
| `research-approach` | 拟采用的研究思路，包括方法、技术路线和可行性论证 |
| `schedule` | 研究工作安排及进度 |
| `references` | 参考文献目录 |
| `advisor-opinion` | 指导教师意见 |

最小入口：

```tex
\documentclass[proposal,humanities]{yibinthesis}

\input{metadata}
\addbibresource{references.bib}

\begin{document}
\makeyibinproposal
\yibinproposalfield{significance}{选题意义。}
\yibinproposalfield{research-status}{国内外研究现状。}
\yibinproposalfield{research-content}{主要研究内容。}
\yibinproposalfield{research-approach}{研究方法、技术路线与可行性论证。}
\yibinproposalfield{schedule}{工作安排及进度。}
\yibinproposalfield{references}{\printyibinproposalbibliography}
\yibinproposalfield{advisor-opinion}{}
\end{document}
```

字段内容较长时会自然分页，不使用固定高度裁切内容。Word 的七个栏目位于一个连续
表单表中，长内容行允许跨页；短内容仍保持官方表单的最小高度。导师意见默认留空；只有配置了
`advisor-signature` 才会在相应签名槽使用图片。

字段内不要使用普通浮动体。图片和表格使用不会漂出表单框的专用接口：

```tex
\yibinproposalfigure[0.6\linewidth]
  {assets/research-route.png}
  {研究技术路线}
  {fig:research-route}

\begin{yibinproposaltable}{研究任务}{tab:tasks}
  \begin{tabular}{cl}
    \toprule
    序号 & 任务 \\
    \midrule
    1 & 收集并核验资料 \\
    \bottomrule
  \end{tabular}
\end{yibinproposaltable}
```

这些题注在 Word 中仍转换为 Caption 样式、`SEQ` 域、书签和可用的 `REF` 交叉引用。
完整案例见 [`examples/proposal`](../examples/proposal/README.md)。

## 文献综述

### 2024 模板核对结果

本项目已对工科类和文科类 2024 文献综述模板做 OOXML 级只读拆解。两份模板共同的硬参数为：

- 4 个 A4 竖版分节，页面尺寸约 11906×16838 twips（21.0×29.7 cm）；左 1701 twips
  （约 3.0 cm），右/上/下分别 1418 twips（约 2.5 cm），页眉/页脚距离 851 twips
  （约 1.5 cm）。第 3 节设置了不同首页页眉/页脚标记，不能把四节压成一个普通正文节。
- 封面后有“毕业论文题目—姓名—单位”信息段，随后是中文摘要、关键词、英文摘要、英文关键词，
  再进入不少于 3000 字的正文。综述模板不生成论文原创性声明、版权授权书和目录。
- 工科类与文科类只在标题编号和个别说明文字上不同：工科采用 `1/1.1/1.1.1`，文科采用
  `一、/（一）/1.`；两者正文均为中文宋体小四、英数 Times New Roman、1.5 倍行距，
  首行缩进两字。一级标题黑体三号，二级标题楷体小三，三级标题宋体四号；图题在图下，
  表题在表上，表格采用三线表。
- 模板文件自身大量使用直接格式和 Normal 段落，而不是完全依赖 Heading 样式；这意味着
  不能仅凭 Word 样式名判断合规，需同时检查段落/运行级字体、字号、行距、缩进和表格 XML。
  `lib/docx_audit.py` 提供只读审计入口，用于输出这些参数。

当前 `yibinthesis-literature-review.sty`、`lib/build_word.py` 已实现独立信息页、正文
分节、阿拉伯页码从 1 开始、标题样式、图表题注和 GB/T 7714 参考文献链。由于本机没有
LibreOffice/Word，尚未完成三份源模板的逐页 PNG 对照；这属于交付前必须补做的视觉风险，
不能仅以 OOXML 静态检查替代。

文献综述生成一页简洁信息页，不生成论文原创性声明、版权授权书、摘要和目录；正文
采用 2024 论文正文版式基线，因此元数据使用 `template-year = 2024`，页码从阿拉伯
数字 1 开始。入口示例：

```tex
\documentclass[literature-review,humanities]{yibinthesis}

\input{metadata}
\addbibresource{references.bib}

\begin{document}
\makeyibinliteraturereviewcover
\input{chapters/body}
\printyibinbibliography
\end{document}
```

正文可使用 `\chapter`、`\section`、`\subsection` 和 `\subsubsection`。Word 输出
分别套用独立的“宜宾论文-一级标题”至“宜宾论文-四级标题”及其多级编号，由这些
自定义样式保证标题格式符合规范。Word 自带的 `Heading 1` 至 `Heading 4` 保持原有
名称和定义，不会被重命名或改写；自定义标题样式无需进入 Word“交叉引用”的“标题”
类型列表。正文、题注、表格与参考文献同样使用可复用的论文命名样式。完整案例见
[`examples/literature-review`](../examples/literature-review/README.md)。

## 构建

从模板仓库根目录构建公开示例：

```powershell
.\build.ps1 all -Main examples/proposal/main.tex
.\build.ps1 all -Main examples/literature-review/main.tex
```

外部项目不需要复制 `build.ps1`。每个文档目录维护自己的
`yibinthesis.project.json`，具体见[外部项目构建配置](project-config.md)。
