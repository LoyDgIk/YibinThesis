# LaTeX 作者接口

本页提供编写用户稿件所需的语义接口和入口示例，可直接据此组织源文件。接续稿件时保留其已有配置与版本字段；若用户安装的插件版本给出了不同接口，采用该版本公开支持的写法。涉及正文改写时，按技能入口选择本次文种的写作指南；只处理语法时无需读取其他文种要求。

## 1. 文种与文件组织

`main.tex` 装配文档，`metadata.tex` 填元数据，`chapters/` 保存内容，`references.bib` 保存来源。这些是用户稿件内的相对位置，不要求预先存在；接续稿件时沿用其实际目录，新建稿件时再按需要创建。

文种选项为 `thesis`、`literature-review`、`proposal`；学科选项为 `humanities` 或 `science`。选择与任务一致的文种，不把正式论文的结构直接套到综述或开题报告。

### 正式论文入口

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
\tableofcontents
\mainmatter
\input{chapters/body}
\printyibinbibliography
\end{document}
```

`chapters/body.tex` 中使用 `\yibinopeningchapter` 写绪论，`\chapter{主题}` 写主体，`\yibinclosingchapter` 写结论。已有稿件的附录、致谢及后置部分顺序保持不变；需要新增时使用插件支持的对应入口。

### 文献综述入口

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

正文先写无标题引言，再用 `\chapter{主题}` 组织综合分析，末章写 `\chapter{结论与展望}`。主题数量依证据与论证需要确定。

### 开题报告入口

```tex
\documentclass[proposal,humanities]{yibinthesis}
\input{metadata}
\addbibresource{references.bib}
\begin{document}
\makeyibinproposal
\yibinproposalfield{significance}{}
\yibinproposalfield{research-status}{}
\yibinproposalfield{research-content}{}
\yibinproposalfield{research-approach}{}
\yibinproposalfield{schedule}{}
\yibinproposalfield{references}{\printyibinproposalbibliography}
\yibinproposalfield{advisor-opinion}{}
\end{document}
```

空值是待填内容，不代表完成稿。字段依次对应选题意义、研究现状、研究内容、方法与技术路线及可行性、工作进度、参考文献、导师意见。导师意见由真实意见提供，否则留空。字段内可用 `\section`、`\subsection` 等表达层级。

## 2. 元数据

使用 `\yibinsetup{...}`。下面只展示匿名输入结构，不预置任何人的资料：

```tex
\yibinsetup{
  title = {},
  english-title = {},
  author = {},
  student-id = {},
  college = {},
  major = {},
  grade = {},
  class-name = {},
  advisor = {},
  advisor-title = {},
  date = {}
}
```

综述所需的单位城市、邮编使用 `affiliation-city`、`affiliation-postcode`；第二导师使用 `second-advisor`、`second-advisor-title`。正式论文的校外导师使用 `external-advisor`、`external-advisor-title`。这些角色按真实资料区分，不互相代填；不存在或未提供的信息留空。

## 3. 摘要与标题

```tex
\begin{cnabstract}
% 在此填写一个完整的中文摘要段落。
\cnkeywords{}
\end{cnabstract}
```

英文对应使用 `enabstract` 和 `\enkeywords{}`。关键词参数填写实际主题词，中文用 `；`、英文用 `;` 分隔，末尾不加标点。示例注释和空参数在形成交付稿时应替换为真实内容。

标题使用 `\chapter`、`\section`、`\subsection` 等语义命令，参数只写标题文字。图表使用 `\caption`、`\label`，正文用 `\ref` 引用；公式见[公式编写](equations.md)。编号由模板生成，作者不把显示编号写进标题或题注。

## 4. 引文与来源数据

使用 `\yibincite{source-key}`（或模板已支持的 `\cite{source-key}`）在具体论述处引用，
`\addbibresource{references.bib}` 载入文献库。同一来源复用一个键。引用不直接提及作者时，
通常将引用命令放在相关句子的句号、逗号、分号等标点之前；若正文直接写出作者姓名，通常
将引用命令紧跟作者姓名之后。这里是依据国标示例的默认写法，不是生成器的强制纠正；如果学校、
学部或导师对标点位置另有要求，按适用要求调整源文件。多个编号按插件压缩为连续范围，例如
`[1-3,5]`，逗号使用英文逗号且不加空格。其他引文命令仅在用户安装的插件明确支持时使用。

Bib 数据保留完整作者、原始题名、实际出版年、文献类型、版本、页码和可核验的 DOI 或 URL。英文书籍数字版次填 `edition = {4}`，不把显示用的版次文字写入数字字段。条目语言用 `langid = {chinese}` 或 `langid = {english}` 按实际来源填写，便于插件选择本地化术语。

正文中的作者简称不用于缩减 Bib 作者名单；著录中的作者省略、大小写与标点由插件处理。文献条目使用适当类型，例如 `@article`、`@book`、`@report`、`@standard`，以来源实际属性为准。

## 5. 源文件检查

- UTF-8 保存；文本中的 `%`、`&`、`_` 等 LaTeX 特殊字符按语义转义，命令参数内的引用键和路径沿用对应接口。
- 检查环境和花括号配对，`\input` 指向存在的文件，标签唯一，引用键在 Bib 库中存在。
- 使用本页列出的作者命令；若插件报告不支持的命令，向用户说明具体位置，并按该版本的公开接口调整稿件，不修改插件实现。
