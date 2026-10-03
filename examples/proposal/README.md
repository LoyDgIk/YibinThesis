# 开题报告示例

该示例使用 `proposal` 类选项与 `template-year = 2022`，严格复刻 2022 年工作表
物理第 5 至 7 页的开题报告模块，不套用 2024 毕业论文封面或前置页。

入口按固定顺序填写七个字段：`significance`、`research-status`、
`research-content`、`research-approach`、`schedule`、`references` 和
`advisor-opinion`。字段内图片使用 `\yibinproposalfigure`，表格使用
`yibinproposaltable`，以便 PDF 和 Word 都保留题注与交叉引用。

构建：

```powershell
.\build.ps1 all -Main examples\proposal\main.tex
```

最终 Word 交付建议让 Microsoft Word 刷新域和分页：

```powershell
.\build.ps1 word -Main examples\proposal\main.tex `
  -CitationMode linked -WordRefresh always
```

完整接口说明见 [`docs/document-types.md`](../../docs/document-types.md)。
