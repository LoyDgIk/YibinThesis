# 文献综述示例

该示例使用 `literature-review` 类选项与 `template-year = 2024`，包含简洁信息页，
正文采用 2024 论文正文版式基线，不生成论文声明、摘要和目录。

入口先调用 `\makeyibinliteraturereviewcover`，再装配正文和参考文献。正文可使用
`\chapter`、`\section`、`\subsection` 和 `\subsubsection`；Word 输出使用独立的
“宜宾论文-一级标题”至“宜宾论文-四级标题”及可复用的正文、题注和参考文献样式，
由自定义标题样式保证格式符合规范。Word 自带的 `Heading 1` 至 `Heading 4` 保持原有
名称和定义，不会被重命名或改写；自定义标题样式无需出现在 Word“交叉引用”的“标题”
类型列表。当前不支持 2025 论文版式。

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
