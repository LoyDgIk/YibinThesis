# 文献综述示例

该示例使用 `literature-review` 类选项与 `template-year = 2024`，按 2024 综述模板
依次装配封面、中英文摘要、正文和参考文献，不生成论文声明或目录。

入口先调用 `\makeyibinliteraturereviewcover`，再依次载入 `chapters/00-abstract-cn`、
`chapters/01-abstract-en`、`chapters/body`，最后调用 `\printyibinbibliography`。
中英文摘要自动同页，摘要前显示题目、作者及学部、所在地、邮编；正文从下一页开始，
引言不设标题，主题章连续排版，最后一章为“结论与展望”，参考文献另起一页。正文可使用
`\chapter`、`\section`、`\subsection` 和 `\subsubsection`；Word 输出使用独立的
“宜宾论文-一级标题”至“宜宾论文-四级标题”及可复用的正文、题注和参考文献样式，
由自定义标题样式保证格式符合规范。Word 自带的 `Heading 1` 至 `Heading 4` 保持原有
名称和定义，不会被重命名或改写；自定义标题样式无需出现在 Word“交叉引用”的“标题”
类型列表。当前不支持 2025 论文版式。

`metadata.tex` 中的 `college` 表示学部，`affiliation-city` 和 `affiliation-postcode`
分别填写所在地和邮编；`advisor`、`advisor-title` 对应指导教师1，`second-advisor`、
`second-advisor-title` 对应指导教师2，没有第二导师时留空。示例中的题目、作者和学校
信息均为待填写占位；`template-year` 是版式版本，不能代替真实年级或完成日期。

中文摘要为 250 字、单段，英文摘要为 100—200 词、单段；中英文关键词一一对应，
各有 3—5 个。综述正文不少于 3000 字，参考文献不少于 20 篇，近五年文献占比不少于
50%。近五年的统计截止日与口径应在研究记录中说明，不能由模板示例年份直接推定。
本例保留可编译占位及空参考文献库，未达到成稿字数或文献要求；成稿须替换占位，
添加已核验文献并在正文实际引用。

构建：

```powershell
.\build.ps1 all -Main examples\literature-review\main.tex
```

可显式选择 Word 引用和刷新模式：

```powershell
.\build.ps1 word -Main examples\literature-review\main.tex `
  -CitationMode linked -WordRefresh always
```

完整接口说明见 [`docs/document-types.md`](../../docs/document-types.md)。
