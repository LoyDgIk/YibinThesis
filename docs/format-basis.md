# 版式依据与版本差异

YibinThesis 是独立实现的非官方模板，不重新分发学校原始文件。默认版式主要参考
当前项目资料中的《附件：毕业论文（设计）模板（人文社科类）2024.docx》，并用
校级《本科毕业论文（设计）撰写规范》补充理工农医编号、页边距、图表公式、注释
和参考文献规则。

## 默认取舍

- `humanities`：正文一级标题采用“一、”编号、三号黑体、左对齐并首行缩进两字，
  不强制每个一级标题另起一页；绪论和结论使用居中、无编号标题。
- `science`：章节采用 `1`、`1.1`、`1.1.1`、`1.1.1.1` 分级编号；绪论和结论
  参与章编号，图表和公式按章编号。
- 摘要、目录、绪论、结论、参考文献、附录和致谢的标题默认按 2024 Word 模板使用
  三号黑体居中。
- 目录使用独立分节且不显示页码；中英文摘要显示大写罗马页码；正文重新从阿拉伯
  数字 1 开始。
- 示例默认保留 2024 Word 模板中的原创性声明与版权使用授权书。所在学院不要求时，
  删除主文件中的 `\makeyibindeclarations` 即可。

## 后置部分顺序冲突

现有两份材料的顺序并不一致：

- 2024 人文社科 Word 模板：参考文献 → 附录 → 致谢；
- 较早校级撰写规范：致谢 → 参考文献 → 附录，并要求文末注释放在参考文献前。

本项目的根示例优先采用较新的 2024 Word 模板，并把可选注释放在参考文献之前：

```tex
\backmatter
\printyibinnotes
\printyibinbibliography
\input{chapters/90-appendix}
\input{chapters/99-acknowledgements}
```

如果学院仍执行较早顺序，只需改为：

```tex
\backmatter
\input{chapters/99-acknowledgements}
\printyibinnotes
\printyibinbibliography
\input{chapters/90-appendix}
```

顺序由 `main.tex` 控制，不需要修改 `yibinthesis.cls`。

## 字体模式

默认模式优先使用宋体、黑体、楷体和 Times New Roman；缺失时回退到 Fandol 与
TeX Gyre Termes，便于跨平台预览和持续集成。正式提交前可启用：

```tex
\documentclass[humanities,strictfonts]{yibinthesis}
```

`strictfonts` 会在任一学校指定字体缺失时报告类错误，避免把兼容字体预览误当作
最终提交版。

## 字体与段落合同

本项目把下列项目作为自动回归合同，而不是只靠成品目测：

| 内容 | 中文字体与字号 | 西文字体 | 段落规则 |
|---|---|---|---|
| 中文正文、附录、致谢 | 宋体，小四（12 pt） | Times New Roman | 1.5 倍行距，首行缩进两字（包括各级标题后的首段） |
| 一级标题 | 黑体，三号（16 pt），加粗 | Times New Roman | humanities 左对齐并缩进两字；science 居中、另起页 |
| 二级标题 | 楷体，小三（15 pt），加粗 | Times New Roman | 左对齐，缩进两字 |
| 三级、四级标题 | 宋体，四号（14 pt），加粗 | Times New Roman | 左对齐，缩进两字 |
| 中文摘要正文 | 宋体，小四（12 pt） | Times New Roman | 1.5 倍行距，首行缩进两字 |
| 英文摘要正文 | Times New Roman，小四（12 pt） | Times New Roman | 1.5 倍行距，顶格，不设首行缩进 |
| 中英文关键词 | 小四（12 pt） | Times New Roman | 标签加粗，整段顶格 |
| 图表标题 | 宋体，五号（10.5 pt） | Times New Roman | 居中 |
| 注释 | 宋体，小五（9 pt） | Times New Roman | 顶格，集中列于文末 |
| 参考文献 | 宋体，五号（10.5 pt） | Times New Roman | 1.5 倍行距，悬挂缩进 |
| 正文页眉与页码 | 宋体，小五（9 pt） | Times New Roman | 居中 |

2024 Word 文件个别示范文字的直接格式与其括号内说明并不完全一致，例如三级标题
示范运行字号小于其标注的“四号”。此类冲突优先采用校级撰写规范和模板中的文字
说明，并由 `tests/audit_format.py` 固化，避免继续继承示例文件中的偶然直接格式。

## 使用边界

学校、学院、专业或指导教师的新通知优先于本项目。模板只能降低机械排版成本，
不能替代提交前的人工核对，也不代表宜宾学院官方认证。
