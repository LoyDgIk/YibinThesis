# YibinThesis

YibinThesis 是一个面向宜宾学院本科毕业论文（设计）的非官方、独立实现 LaTeX
模板。同一组 LaTeX 源文件可生成排版主输出 PDF，以及便于继续修改的 DOCX。

本项目不是宜宾学院发布或认证的官方模板。学校、学院或专业的新通知优先于本项目，
提交前仍需由指导教师或所在学院复核。

## 主要能力

- `humanities` 人文社科编号：`一、`、`（一）`、`1.`、`（1）`。
- `science` 理工农医编号：`1`、`1.1`、`1.1.1`、`1.1.1.1`。
- 封面、原创性声明、版权授权书、中英文摘要、独立目录、正文和后置部分。
- A4 单面版式，指定页边距、摘要罗马页码、正文阿拉伯页码及题名页眉。
- 三线表、插图、按章公式、顺序编码引用和文末集中注释。
- `biblatex-gb7714-2015` PDF 文献链与 GB/T 7714-2015 Word CSL。
- Pandoc DOCX 转换，保留真实 Heading 样式、四个 Word 分节、动态目录和页码域。
- 根示例与真实课题框架回归案例。

## 快速开始（Windows）

在仓库根目录执行：

```powershell
# 检查 PDF 与 Word 工具链
.\build.ps1 doctor

# 检查字体、字号、行距、首行缩进并运行两种学科烟测
.\build.ps1 check

# 同时生成 PDF 和 DOCX
.\build.ps1 all
```

默认输出：

- `build/pdf/main.pdf`
- `build/word/main.docx`

只构建一种格式：

```powershell
.\build.ps1 pdf
.\build.ps1 word
```

指定其他入口：

```powershell
.\build.ps1 all -Main examples/low-altitude-standard-demand/main.tex
```

对应输出为：

- `build/pdf/examples/low-altitude-standard-demand/main.pdf`
- `build/word/low-altitude-standard-demand.docx`

清理指定入口的构建产物：

```powershell
.\build.ps1 clean
.\build.ps1 clean -Main examples/low-altitude-standard-demand/main.tex
```

## 项目内工具链

`build.ps1` 按以下优先级寻找工具：环境变量覆盖、项目 `.tools/`、PATH、已知本机
缓存。当前验证组合为：

- Tectonic `0.16.9`
- Biber `2.17`
- Pandoc `3.9.0.2`
- Python `3.11+` 与 `python-docx 1.2.x`

Tectonic 0.16.9 使用 BCF 3.8，必须搭配 Biber 2.17；Biber 2.21 与该链不兼容，
构建脚本会直接阻止错误组合。

大体积二进制放在被 Git 忽略的 `.tools/`，不会进入模板源码提交。可从已有安装复制：

```powershell
.\tools\setup_toolchain.ps1 `
  -TectonicPath C:\path\to\tectonic.exe `
  -BiberPath C:\path\to\biber.exe `
  -PandocPath C:\path\to\pandoc.exe
```

该脚本还会在 `.tools/venv/` 创建 Word 转换环境。若只复制二进制，可加
`-SkipPythonEnvironment`。版本和当前验证文件哈希记录在
`tools/toolchain.lock.json`。

也可使用环境变量：

- `YIBINTHESIS_TECTONIC`
- `YIBINTHESIS_BIBER`
- `YIBINTHESIS_PANDOC`
- `YIBINTHESIS_PYTHON`
- `YIBINTHESIS_LATEXMK`
- `YIBINTHESIS_XELATEX`

## TeX Live 与跨平台构建

安装了 XeLaTeX、latexmk、Biber 和 `biblatex-gb7714-2015` 时，可直接运行：

```bash
latexmk -xelatex -outdir=build/pdf main.tex
```

Word 转换可直接调用：

```bash
python -m pip install -r requirements-word.txt
python tools/build_word.py --main main.tex --output build/word/main.docx
```

Pandoc 可通过 `--pandoc`、`PANDOC` 环境变量、项目 `.tools/` 或 PATH 提供。

## 使用步骤

1. 编辑 `metadata.tex`。
2. 替换 `chapters/` 中的摘要、正文、附录和致谢。
3. 在 `references.bib` 中维护已经核验的文献。
4. 执行 `.\build.ps1 all`。
5. 人工检查 PDF 和在 Microsoft Word 中更新后的 DOCX。

集中式元数据示例：

```tex
\yibinsetup{
  title          = {论文中文题目},
  english-title  = {English Title},
  author         = {姓名},
  student-id     = {学号},
  college        = {学院},
  major          = {专业},
  grade          = {年级},
  advisor        = {校内导师},
  advisor-title  = {职称},
  date           = {20XX 年 X 月 X 日},
  secrecy        = public
}
```

英文题名写入 PDF/DOCX 元数据，但 2024 人文社科封面不显示英文题名。

## 学科类型与章节

根示例默认：

```tex
\documentclass[humanities]{yibinthesis}
```

理工农医类改为：

```tex
\documentclass[science]{yibinthesis}
```

示例章节使用 `\yibinopeningchapter` 和 `\yibinclosingchapter`：在人文社科模式中，
绪论、结论为居中无编号标题；在理工农医模式中自动作为编号章，避免出现 `0.1`。

最终提交前可强制检查学校字体：

```tex
\documentclass[humanities,strictfonts]{yibinthesis}
```

未启用 `strictfonts` 时，缺失的宋体、黑体、楷体和 Times New Roman 会回退到
Fandol/TeX Gyre 字体，适合预览和持续集成，不代表最终字体合规。

## 引用与注释

```tex
相关规则见国家标准\yibincite{gbt7714-2015}。
需要说明的文字\yibinnote{注释内容。}
```

主文件中的 `\printyibinnotes` 输出集中注释，`\printyibinbibliography` 输出参考
文献。后置顺序由 `main.tex` 决定；现有资料存在版本差异，详见
[版式依据与版本差异](docs/format-basis.md)。

## PDF 与 Word 的边界

PDF 是版式主输出。DOCX 用于继续编辑，不承诺与 PDF 逐页像素一致。

Word 转换支持本项目受控子集：标准章节、普通段落、常见表格/图片/公式、引用、
集中注释和示例自定义命令。复杂 TikZ、PSTricks、任意低层宏、复杂浮动体或手工
分页可能需要在 Word 中调整。

转换时会解析图、表、公式和 `\ref`/`\eqref` 的全局编号；这些可见编号目前是
转换时生成的文本，不是 Word 的 `SEQ/REF` 动态域。若在 Word 中增删或移动图表公式，
应回到 LaTeX 源重新转换。目录与页码使用动态 Word 域，打开 Word 后可更新。

安装 Microsoft Word 的 Windows 电脑可自动更新域并可选导出检查 PDF：

```powershell
.\tools\refresh_word.ps1 `
  -InputPath build/word/main.docx `
  -PdfOutput build/qa/word/main.pdf
```

## 回归案例

`examples/low-altitude-standard-demand/` 使用当前毕业论文资料中的低空经济课题方法
框架测试长题名、双语摘要、science 一至四级标题与编号、技术路线图、三线表、公式、交叉引用、
注释和参考文献。个人信息均为示例值，不包含虚构实验结果或需求排序。

`tests/` 还包含 humanities/science 编译烟测，覆盖跨章编号、目录标点和交叉引用。
`tests/audit_format.py` 会直接审计 LaTeX 类与 Word 样式中的宋体/黑体/楷体/
Times New Roman、字号、1.5 倍行距、中文正文及标题后首段的两字缩进、英文摘要顶格、页边距
和页眉页脚距离；`build.ps1 check` 会把格式审计与两项烟测一次执行完。

## 目录结构

```text
YibinThesis/
├── yibinthesis.cls
├── main.tex
├── metadata.tex
├── references.bib
├── chapters/
├── examples/
├── tests/
├── assets/
├── word/
├── tools/
├── docs/
├── build.ps1
├── Makefile
└── latexmkrc
```

## 许可证与来源

模板源代码采用 MIT License。校名、校徽、校名字样及其他机构标识不包含在 MIT
授权范围内。参考项目和独立实现边界见 `NOTICE`。
