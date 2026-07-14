# 低空经济标准需求回归案例

该案例从当前工作区已经核验的课题材料中提取题目、四层方法框架、评价指标与技术路线图，专门测试：

- `science` 一至四级标题和图表编号；
- 长中文/英文题目；
- 双语摘要、独立目录、science 编号绪论与结论；
- 外部 PNG、三线表、公式、交叉引用和文末注释；
- GB/T 7714-2015 引文和参考文献；
- 从同一组 LaTeX 章节生成 PDF 和 DOCX。

作者、学号、学院和导师均为示例值。案例不包含实证样本量、评价得分或标准需求排序，也不应被当作论文研究结论引用。

请从仓库根目录通过统一构建脚本指定本入口：

```powershell
.\build.ps1 pdf -Main examples/low-altitude-standard-demand/main.tex
.\build.ps1 word -Main examples/low-altitude-standard-demand/main.tex
```

也可一次生成两种格式：

```powershell
.\build.ps1 all -Main examples/low-altitude-standard-demand/main.tex
```
