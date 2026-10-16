# YibinThesis LaTeX 写作规范

适用于宜宾学院 2024 年正式论文、文献综述模板，以及 YibinThesis 的 2022 年开题报告表单。

## 1. 作者输入与生成工具

| 作者负责 | YibinThesis 负责 |
|---|---|
| 真实题目、作者、单位、导师与日期；文种和学科类别 | 封面标签、信息槽及其位置 |
| 章节层级、正文论证、摘要与关键词、结论 | 字体、字号、行距、页边距、页眉页脚、目录与分页 |
| 图表的数据、图像文件、标题、表头、单位、必要注释和来源 | 图表编号、题注位置、三线表线型、表格续页 |
| 公式的学术含义、符号、单位、来源及交叉引用 | 公式编号、Word 原生公式及版式 |
| 准确的 Bib 文献信息、引用键和引用位置 | GB/T 7714 著录呈现、引用序号和链接 |

在 LaTeX 源文件中写正文，使用 `\chapter`、`\section`、`\subsection`、`\caption`、`\label` 等结构命令。标题只写名称，编号由模板产生；图表提供内容、标题、单位和来源即可。

正文基于实际研究和真实文献。区分已经完成的工作与拟开展的工作，不编造数据、结果或引用。

三类稿件的正文首次出现非通用英文缩写时写英文全称，可用“中文名（English Full Name，缩写）”。DNA、RNA、HPLC 等常见缩写可直接使用。K-means 是方法名称，GB/T 是标准编号前缀，按各自含义书写。

正文引用国外文献并提及作者时，只写作者的姓。文后姓名著录按第 7.1 节处理。

## 2. 选择文种和学科

一个 `main.tex` 对应一种文档，文种选项与学科选项分开。

| 文种 | 类选项 | 学科选项（二选一） | `template-year` |
|---|---|---|---|
| 正式毕业论文（设计） | `thesis`，也是默认值 | `humanities` 人文社科；`science` 理工农医 | `2024` |
| 文献综述 | `literature-review` | `humanities` 文科；`science` 工科 | `2024` |
| 开题报告 | `proposal` | `humanities` 或 `science` | `2022` |

模板年份标识生成工具采用的版式版本，提交年份来自实际提交计划或日期。二者用途不同；近五年文献统计使用提交年份。

## 3. 填写元数据

元数据集中放在 `metadata.tex` 的 `\yibinsetup{...}` 中，填写实际资料。

| 键 | 含义与适用范围 |
|---|---|
| `title`、`english-title` | 实际中文、英文题目；英文题目是否展示由所选文种控制 |
| `author`、`student-id` | 作者姓名、学号 |
| `college` | 学院或学部；综述封面“学部（院）”与摘要前单位行共用此值 |
| `major`、`grade`、`class-name` | 专业、年级与确有需要的班级 |
| `affiliation-city`、`affiliation-postcode` | 综述摘要前单位行中的城市、邮编 |
| `advisor`、`advisor-title` | 综述的指导教师 1 及职称；正式论文的校内导师及职称 |
| `second-advisor`、`second-advisor-title` | 综述的指导教师 2 及职称 |
| `external-advisor`、`external-advisor-title` | 正式论文的校外导师及职称 |
| `advisor-signature` | 导师签名图片在稿件内的相对路径，例如 `assets/signature.png`；图片由生成工具放入签名位置 |
| `date`、`template-year` | 实际完成日期、所选模板版本；日期呈现由文种决定 |

导师或职称暂缺时用 `{}` 留空，生成器保留对应标签与空槽。综述没有第二导师时，两项第二导师字段均留空。

邮编填入 `affiliation-postcode`，例如 `644000`；学号填入 `student-id`。没有独立的“编号”字段。姓名、单位、地区和日期填写作者的实际信息。

以下仅展示综述元数据结构，空值需由实际资料补全；没有第二导师时保留空值：

```tex
\yibinsetup{
  template-year = 2024,
  title = {},
  author = {},
  student-id = {},
  college = {},
  major = {},
  grade = {},
  affiliation-city = {},
  affiliation-postcode = {},
  advisor = {},
  advisor-title = {},
  second-advisor = {},
  second-advisor-title = {},
  date = {}
}
```

## 4. 文献综述

### 4.1 入口和章节

入口按以下顺序编写：

```tex
\documentclass[literature-review,humanities]{yibinthesis}
\input{metadata}
\addbibresource{references.bib}

\begin{document}
\makeyibinliteraturereviewcover
\input{chapters/00-abstract-cn}
\input{chapters/01-abstract-en}
\input{chapters/body}
\printyibinbibliography
\end{document}
```

工科把 `humanities` 换成 `science`，其余作者接口一致。两个摘要文件之间直接衔接；同页摘要、正文另页及参考文献另页由综述生成模块安排。

`chapters/body.tex` 的每一章都用 `\chapter{实际主题名称}` 写作，交代综述对象、范围和组织线索的内容并入首章，正文不写无标题的悬置段。首章之后按实际主题继续分章，按需要使用更低层级；末章固定为 `\chapter{结论与展望}`，与前面主题章同级同样式，归纳主要认识、研究局限与后续方向。主题数量取决于内容。

综述采用“封面—双语摘要—正文—参考文献”的顺序，不设置目录、原创性声明或版权授权页。

### 4.2 摘要页内容

生成器在摘要页放置题目、作者及“学部、城市、邮编”单位行，再依次呈现中文摘要、中文关键词、英文摘要和英文关键词。作者只提供元数据与下列内容。

| 内容 | 作者要求 |
|---|---|
| 中文摘要 | 约 250 字，单段、独立成文；概括综述范围、主要归纳和综合判断 |
| 英文摘要 | 100–200 个英文单词，内容与中文对应，保持语法、时态和单复数一致 |
| 中文关键词 | 3–5 个实际主题词，用中文分号 `；` 分隔，末尾无标点 |
| 英文关键词 | 与中文数量、顺序、含义一一对应，用英文分号 `;` 分隔，末尾无标点 |
| 内容边界 | 摘要不标注引用文献编号；不宜使用公式、图表；不把拟开展工作写成已验证成果 |

中文摘要在 `cnabstract` 环境内写一个正文段，并调用 `\cnkeywords{词条列表}`；英文摘要在 `enabstract` 环境内写对应文本，并调用 `\enkeywords{词条列表}`。关键词命令置于对应环境结束前。输入文本可以换行以便编辑，正文中不以空行或 `\par` 拆成数段。

中英文摘要排在同一页，两个摘要文件之间不添加 `\clearpage`。文科英文摘要尽量避免 “This paper introduces…”、“the author discusses…” 和 “in this paper”等开头。

### 4.3 正文写法

净正文至少 3000 字，计数方法见第 7 节。围绕文献的研究对象、方法、主要认识及适用范围展开比较和归纳。

文科综述（`humanities`）正文提及外国作者时，将姓的全部字母大写。

## 5. 正式毕业论文

正式论文依次包含封面、原创性声明与版权授权书、中英文摘要、目录、绪论、主体、结论及后置部分。中英文摘要各自成页，直接从摘要标题开始，无需另写作者单位行。

入口使用 `\makeyibincover`、`\makeyibindeclarations`；`\frontmatter` 中装配两份摘要与 `\tableofcontents`，`\mainmatter` 中装配绪论、主体及结论。文后参考文献用 `\printyibinbibliography`，附录、致谢按稿件需要装配。

中文题目一般不宜超过 20 字，英文一般不宜超过 10 个实词，必要时加副标题。正文标题一般在 15 字以内，不使用标点符号，尽量少用英文缩写。

摘要写成单段，中文约 250 字，概括本论文的内容、方法、成果和结论；英文完整对应中文内容，按内容确定长度。使用 `cnabstract`、`enabstract`、`\cnkeywords`、`\enkeywords`。关键词为 3–5 个，中英文对应、使用各自分号且末尾不加标点。摘要不标引用编号，不宜使用公式和图表。

### 5.1 绪论

绪论不少于 3000 字，包含以下五节；摘要保持单段：

```tex
\yibinopeningchapter
\section{研究的背景}
\section{国内外研究现状}
\section{选题的目的}
\section{课题研究的意义}
\section{研究的内容及方法}
```

背景说明问题形成的条件；研究现状比较已有研究；目的界定本研究要回答的问题；意义说明回答该问题的价值；内容及方法说明实际研究任务、材料与方法。目的与意义分别展开。

### 5.2 主体与结论

主体按实际研究组织。理工农医论文写清推导、方案、实验与数据分析；人文社科论文写清问题分析、比较、案例或实证论证。使用真实研究材料支撑结论。

主体与结论合计一般不少于 8000 字，与绪论分别计算。

结论用 `\yibinclosingchapter`，独立成章、不加章号，概括主要发现、局限和建议。两类正式论文均应写结论。

## 6. 开题报告入口

开题报告采用 `\documentclass[proposal,humanities]{yibinthesis}` 或 `science`，`template-year = 2022`，按表单栏目填写。

`\makeyibinproposal` 后使用 `\yibinproposalfield{键}{内容}`，以下七栏按顺序各填写一次：

| 键 | 内容 |
|---|---|
| `significance` | 选题意义 |
| `research-status` | 国内、国外研究现状概述 |
| `research-content` | 主要研究内容 |
| `research-approach` | 拟采用的方法、技术路线及可行性 |
| `schedule` | 真实工作安排与进度 |
| `references` | 使用 `\printyibinproposalbibliography` |
| `advisor-opinion` | 有真实导师意见时填入；否则留空 |

按“拟研究、拟采用、计划”表达尚未完成的工作。字段内容自然续页，表单尺寸和分页由工具处理。

每个 `\yibinproposalfield` 的正文内，用 `\section`、`\subsection`、`\subsubsection`、`\paragraph` 表示该单元格的一级至四级标题。标题参数只写名称，PDF 和 Word 自动使用以下编号：

| 学科选项 | 一级 | 二级 | 三级 | 四级 |
|---|---|---|---|---|
| `humanities` | 一、 | （一） | 1. | （1） |
| `science` | 1 | 1.1 | 1.1.1 | 1.1.1.1 |

每个字段的标题从头编号，字段之间不承接；图表编号不随字段标题重置。下面是两个字段的标题片段：

```tex
\yibinproposalfield{research-content}{
  \section{研究对象}
  \subsection{分析范围}
}
\yibinproposalfield{research-approach}{
  \section{研究方法}
  \subsection{资料整理}
  \subsubsection{分类步骤}
  \paragraph{类别定义}
}
```

## 7. 引文和字数

### 7.1 引文内容和著录

在 `references.bib` 中填写真实、准确的文献信息，用 `\cite{key}` 在相应论述处引用。引文编号由工具按首次出现顺序生成，同一文献复用一个键和编号。引用具体段落或结论时，按需要补充页码。

参考文献采用 GB/T 7714-2015。Bib 内保留完整的作者及机构、题名、年份、载体、版本和 DOI/URL 等信息，由工具处理大小写、作者省略和著录格式。正文“只写姓”的规则不用于删改 Bib 作者名字。

### 7.2 数量与近五年

| 文种 | 文献数量 | 近五年比例 |
|---|---|---|
| 两类文献综述 | 至少 20 篇 | 至少 50% |
| 正式论文 | 非艺体类至少 20 篇；艺体类一般 15 篇以上 | 至少 50% |

按正文实际引用的不同文献计数，重复引用只算一篇。普通网页材料不计入上述数量；通过 DOI、URL 或在线平台获取的正式学术出版物仍按其文献类型计数。

本手册按提交年 `Y` 计算近五年，范围为含当年的五个公历年 `Y−4` 至 `Y`；例如 2026 年提交时为 **2022–2026**。

年份以实际出版信息及 Bib 的 `year`／`date` 为准，不从 citekey 的数字推断。政策版名、印发日期和公开日期分别按实际引用版本填写。近五年比例为“上述范围内的文献数／可计入的引用文献总数”；缺失年份先补全。

### 7.3 净正文字数

按正文段落统计汉字，排除 LaTeX 命令、注释、引文标号及标点；封面、摘要、章节标题、图表及题注、参考文献不计入净正文。连续外文或数字词可单独统计，非空字符数不等于正文字数。
