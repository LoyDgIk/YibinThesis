# 文档类型与入口

YibinThesis 的一个 `main.tex` 只生成一种文档。文档类型由 `\documentclass` 选项决定，
学科类型由 `humanities` 或 `science` 决定，两者可以组合使用。

| 文档类型 | 类选项 | `template-year` | 主要入口命令 | 版式来源 |
|---|---|---:|---|---|
| 毕业论文（设计） | `thesis`，可省略 | `2024` | `\makeyibincover`、`\makeyibindeclarations` | 2024 论文模板与校级撰写规范 |
| 开题报告 | `proposal` | `2022` | `\makeyibinproposal` | 2022 工作表物理第 5 至 7 页的开题报告模块 |
| 文献综述 | `literature-review` | `2024` | `\makeyibinliteraturereviewcover` | 2024 工科类、文科类文献综述模板 |

`template-year` 必须与上表逐项对应。它登记当前文档实际采用的版式来源，不会自动
选择或混用其他年份；类型与年份不匹配时，LaTeX 和 Word 构建都会报错并停止。
当前不支持 2025 论文模板。

源码按职责拆分：`yibinthesis.cls` 提供三类文档共用的元数据、字体、章节、引用与
图表基础，开题报告和文献综述分别由 `yibinthesis-proposal.sty`、
`yibinthesis-literature-review.sty` 实现。使用者仍只需加载 `yibinthesis` 类，类会按
文档类型自动加载对应模块，不需要在 `main.tex` 中手动 `\usepackage`。

`new` 创建的三类项目默认带有 `no-doi`，隐藏参考文献列表中的 DOI，URL 仍保留。
需要显示 DOI 时，删除 `main.tex` 中的 `no-doi` 即可；未带隐藏选项的既有项目仍显示来源中已有的字段。
若所在学部有不同著录要求，可在
`\documentclass` 选项中分别加入 `no-url` 或 `no-doi`；两项互相独立，也可同时使用。
该设置同时控制 PDF 与 Word 中参考文献列表的显示，不删除 `references.bib` 中的来源数据。
例如：

```tex
\documentclass[proposal,humanities,no-doi]{yibinthesis}
```

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

正式论文的 `chapters/10-introduction.tex` 使用 `\yibinopeningchapter` 开始绪论，
并按两类 2024 正式论文模板保留以下五个小节：

```tex
\yibinopeningchapter
\section{研究的背景}
\section{国内外研究现状}
\section{选题的目的}
\section{课题研究的意义}
\section{研究的内容及方法}
```

模板要求绪论不少于 3000 字。这一结构属于正式论文；文献综述使用下文的无标题引言。

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
- 页码从 1 开始。封面不参与页码，正文表单从第 1 页起编。LaTeX
  `yibinthesis-proposal.sty` 与 Word 后处理都必须保持这一点，并确保正文表单不是
  被错误拆成七个独立页。

当前 `yibinthesis-proposal.sty` 和 `lib/build_word.py` 已实现独立 proposal profile、
七字段顺序校验、单表结构和从第 1 页开始的页码。仍需在装有 Word 或 LibreOffice 的
环境中做最终分页复核；本机当前没有 `soffice`，无法完成 PNG 级视觉验收。

开题报告只复刻 2022 工作表物理第 5 至 7 页的开题栏目，页码从 1 开始；不生成
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

每个 `\yibinproposalfield` 的内容可按需要使用以下四级标题。编号在每个字段开始时独立
重置，字段之间不承接；标题参数只填写名称，序号由 PDF 和 Word 自动生成。

| 标题命令 | 字段内级别 | `humanities` 编号 | `science` 编号 |
|---|---|---|---|
| `\section{标题}` | 一级 | 一、 | 1 |
| `\subsection{标题}` | 二级 | （一） | 1.1 |
| `\subsubsection{标题}` | 三级 | 1. | 1.1.1 |
| `\paragraph{标题}` | 四级 | （1） | 1.1.1.1 |

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

本项目已对《质检学部-文献综述模板（工科类）2024（模板）》和文科类对应模板做
OOXML 级只读核查，包括正文、表格和批注。两类综述均使用 `template-year = 2024`，
结构为封面、同页中英文摘要、正文、参考文献。

- 摘要页先放论文题目、作者，以及由学部、城市、邮编组成的单位行，再依次放中文摘要、
  中文关键词、英文摘要、英文关键词。两份模板的批注均明确要求中英文摘要在一页。
- 中文摘要为单段，模板标为 250 字；英文摘要为 100—200 个英文单词。关键词为 3—5 个，
  中英文一一对应，分别使用中文、英文分号分隔，末尾不加标点。
- 正文另起一页，以无标题的引言段落开始；引言说明综述目的、概念与范围、
  资料来源和研究现状，不另设“引言”章标题。随后是连续排列的主题章，末章为
  “结论与展望”，各章一律使用 `\chapter{章标题}` 写作，与主题章同级同样式。
  各主题章不强制另起一页。参考文献单独起页。
- 按模板结构默认不生成目录，也不生成正式论文的原创性声明和版权授权书。模板没有
  “禁止目录”的明文规定；正式论文的目录和五节绪论不移入综述。
- 正文不少于 3000 字，参考文献不少于 20 篇，其中近五年的文献不少于 50%。两份综述模板
  均无艺体类 15 篇的例外。“近五年”没有规定截止日或日历年口径，统计时应记录采用的
  截止日和口径，不能由 `template-year = 2024` 固定出一组文献年份。摘要、参考文献、
  模板说明和占位文字不用于凑正文长度。

源文件使用 A4 竖版，左边距约 3.0 cm，右/上/下约 2.5 cm，页眉/页脚距离约 1.5 cm。
正文为中文宋体小四、英数 Times New Roman、1.5 倍行距、首行缩进两字。工科标题编号
采用 `1/1.1/1.1.1`，文科采用 `一、/（一）/1.`；一级标题黑体三号，二级标题楷体小三，
三级标题宋体四号。图题在图下，表题在表上，表格采用三线表。工科模板有一处三级标题的
字号说明与其余同级示例冲突，不能据该孤例改写整个标题层级。

模板大量使用直接格式和 Normal 段落，不能只凭 Word 样式名判断版式。
`lib/docx_audit.py` 可输出段落、运行、表格等 OOXML 参数；分页和摘要同页效果还需检查
实际构建输出。

入口按以下顺序装配，两个摘要使用 `cnabstract`、`enabstract` 环境，摘要文件之间不插入
`\clearpage` 或 `\tableofcontents`：

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

`chapters/body.tex` 开头直接写引言段落，不使用 `\chapter{引言}` 或
`\yibinopeningchapter`；主题章使用 `\chapter{实际主题名称}`，结尾使用
`\chapter{结论与展望}`。主题章数量由综述内容决定，不把模板的示例章数当作固定要求。

封面“学部（院）”与摘要前单位行共用 `college`；单位行的城市、邮编分别使用
`affiliation-city`、`affiliation-postcode`。`advisor`、`advisor-title` 对应指导教师 1
及职称，`second-advisor`、`second-advisor-title` 对应指导教师 2 及职称；旧的
`external-advisor`、`external-advisor-title` 保持兼容。综述的第二导师不据字段旧名
推定为校外导师。所有字段均按实际资料填写，第二导师无资料时留空；模板中的单位、
年级和完成年份示例不作为新项目默认事实。CLI 对应选项见[外部项目构建配置](project-config.md)。

正文还可使用 `\section`、`\subsection` 和 `\subsubsection`。Word 将四级标题
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
