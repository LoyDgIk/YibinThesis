# 完整功能回归案例

该案例是与任何具体论文相互独立的排版夹具，围绕“学术文档双格式转换与质量校验”这一通用主题，专门测试：

- 默认 `thesis` 文档类型与 `template-year = 2024`；
- `science` 一至四级标题和图表编号；
- 长中文/英文题目；
- 双语摘要、正文目录、图目录、表目录、science 编号绪论与结论；
- 外部 PNG、三线表、公式、交叉引用和文末注释；
- GB/T 7714-2015 引文和参考文献；
- 从同一组 LaTeX 章节生成 PDF 和 DOCX。

作者、学号、学院和导师均为示例值。案例不包含真实研究数据或评价结果，也不应被当作论文研究结论引用。

请从仓库根目录通过统一构建脚本指定本入口：

```powershell
.\build.ps1 pdf -Main examples/full-featured/main.tex
.\build.ps1 word -Main examples/full-featured/main.tex
```

需要 Word 完成域更新、重新分页和续表处理时：

```powershell
.\build.ps1 word -Main examples/full-featured/main.tex `
  -CitationMode linked -WordRefresh always
```

也可一次生成两种格式：

```powershell
.\build.ps1 all -Main examples/full-featured/main.tex
```
