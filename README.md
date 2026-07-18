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
- 三线表、插图、点号式按章图表/公式编号、顺序编码引用和文末集中注释。
- `biblatex-gb7714-2015` PDF 文献链与 GB/T 7714-2015 Word CSL。
- Pandoc 负责正文语义转换；Word 版以命名样式驱动字体、字号、段落节奏和封面填写线，
  OOXML 仅补充分节、域、书签、固定表格几何等 Word 对象模型不能稳定表达的部分。
- 根占位示例与不绑定具体课题的完整功能回归案例。

## 快速开始（Windows）

在仓库根目录执行：

```powershell
# 检查 PDF 与 Word 工具链
.\build.ps1 doctor

# 检查字体、字号、行距、首行缩进并运行两种学科烟测
.\build.ps1 check

# 同时生成 PDF 和 DOCX
.\build.ps1 all

# 最终 Word 交付：要求 Microsoft Word 完成域、续表和分页刷新
.\build.ps1 word -CitationMode linked -WordRefresh always
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
.\build.ps1 all -Main examples/full-featured/main.tex
```

对应输出为：

- `build/pdf/examples/full-featured/main.pdf`
- `build/word/full-featured.docx`

构建仓库外的论文源时，推荐在论文项目根目录放置
`yibinthesis.project.json`，不再为每个论文复制一份 `build.ps1`：

```json
{
  "$schema": "../YibinThesis/yibinthesis.project.schema.json",
  "schemaVersion": 1,
  "main": "latex/main.tex",
  "outputRoot": "build",
  "checkOutputRoot": "build/checks",
  "citationMode": "linked",
  "wordRefresh": "auto",
  "deliverables": {
    "pdf": "build/deliverables/thesis.pdf",
    "word": "build/deliverables/thesis.docx"
  }
}
```

从论文项目目录执行模板仓库中的统一构建器时，会自动发现该文件；也可显式传入：

```powershell
& 'D:\tools\YibinThesis\build.ps1' doctor -Config .\yibinthesis.project.json
& 'D:\tools\YibinThesis\build.ps1' check  -Config .\yibinthesis.project.json
& 'D:\tools\YibinThesis\build.ps1' all    -Config .\yibinthesis.project.json
```

外部入口的 `metadata.tex`、`references.bib`、章节和图件只从入口目录解析；缺失时
构建会直接失败，不会静默使用模板仓库里的同名占位文件。模板类会复制到输出目录的
隔离运行区，Word 参考样式和 CSL 仍由本仓库提供。`-OutputRoot` 可以使用中文路径；
Biber 2.17 阶段会通过临时 ASCII 目录联接访问同一输出目录。

清理指定入口的构建产物：

```powershell
.\build.ps1 clean
.\build.ps1 clean -Main examples/full-featured/main.tex
```

## 项目内工具链

`build.ps1` 按以下优先级寻找工具：环境变量覆盖、项目 `.tools/`、PATH、已知本机
缓存。当前验证组合为：

- Tectonic `0.16.9`
- Biber `2.17`
- Pandoc `3.9.0.2`
- Python `3.11+`、`python-docx 1.2.x` 与 Pillow

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
直接调用 Python 只生成结构化 DOCX，不执行 Microsoft Word COM 刷新，因此不会拆分
跨页续表。最终交付建议使用 `build.ps1`。

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
  class-name     = {班级},
  advisor        = {校内导师},
  advisor-title  = {职称},
  secrecy        = public,
  logo           = {builtin},
  author-signature = {assets/本人签名.jpg},
  advisor-signature = {},
  signature-background = whiten
}
```

`grade` 建议只填写入学年级数字，模板会兼容旧项目并自动补“级”；可选的
`class-name` 填写完整班级名称。例如 `grade={2023}, class-name={7班}` 在 PDF 与
DOCX 封面均显示为“2023级7班”。若旧源文件已将完整文字写入 `grade`，只要其中已含
“级”就不会再次追加。

英文题名写入 PDF/DOCX 元数据，但 2024 人文社科封面不显示英文题名。该版封面也
不显示提交日期，因此默认元数据不再提供日期项；旧项目中的 `date` 键仍可保留，
但不会出现在默认封面上。

`logo={builtin}` 使用模板仓库内置且经哈希审计的官方校徽，不需要在外部论文中填写
跨目录长路径。`author-signature` 和 `advisor-signature` 均相对于 `metadata.tex` 所在目录
解析；留空或填写 `none` 会保留干净的空签名槽。作者签名会同时进入原创性声明和版权
授权书，PDF 与 DOCX 使用同一图片源。`signature-background=preserve`（默认）保留图片
原貌；设为 `whiten` 时，构建器在 `build/` 中生成白底派生 PNG，原始签名文件不会被改写。

## 表格与附录

表格列的 Word 宽度和对齐以 LaTeX 列规格为准。例如：

```tex
\begin{tabularx}{\textwidth}{C{0.18\textwidth}L{0.27\textwidth}>{\raggedright\arraybackslash}X}
  \toprule
  编号 & 名称 & 说明 \\
  \midrule
  1 & 示例 & 多行文字 \\
  \bottomrule
\end{tabularx}
```

- `L/C/R{宽度}` 分别为左、中、右对齐，使用 `m` 列的垂直居中语义。
- `X` 取得固定列之后的剩余宽度，并采用垂直居中语义；可用 `>{...}X` 指定水平对齐。
- 原生 `l/c/r` 保留水平对齐；`p/m/b{宽度}` 分别映射到顶端、居中、底端对齐。
- 普通表格单元格不继承正文首行缩进。Word 会按源码列类型套用“宜宾论文-表格正文/
  居中/右对齐”和对应表头样式，用户也可在样式库中手动修改或复用。

`longtable` 中的 `\endfirsthead`、`\endhead` 仍按标准 LaTeX 写法维护。PDF 由
LaTeX 自行分页；Word 需要 `-WordRefresh always` 才会根据最终分页插入“续表”题、
重复表头并压缩可以安全合并的续页。

多个附录使用：

```tex
\begin{yibinappendices}
\yibinappendix{材料清单}
% 附录 A 内容

\yibinappendix{参数说明}
% 附录 B 内容
\end{yibinappendices}
```

每个 `\yibinappendix` 都生成新的顶层附录并另起一页，PDF 与 Word 使用同一标题语义。

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
绪论、结论均为居中无编号标题；在理工农医模式中，绪论自动作为编号章以避免出现
`0.1`，结论仍按校级规范单独成章但不加章号，并加入目录。

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
文献。校级规范要求后置部分采用“致谢 → 参考文献 → 附录”；存在集中注释时，注释
置于参考文献之前。具体装配命令见[版式依据与版本差异](docs/format-basis.md)。

Word 引用有两种构建模式：

- `-CitationMode linked`（默认）：保留 citeproc 生成的 GB/T 7714 顺序编号，正文
  上标通过 `REF` 域跳转到参考文献书签。
- `-CitationMode native`：在上述稳定参考文献列表之外，写入 Word bibliography
  `customXml` 来源并生成 `CITATION` 域。需要在 Word 中选择项目附带的数字样式时，
  可显式运行 `./tools/install_word_bibliography_style.ps1`，构建过程不会静默修改全局配置。

Word 图题和表题使用项目注册的中文题注标签 `图`、`表`，公式使用 `公式`。首次在某个
Windows 用户下使用模板时，先关闭所有 Word 窗口并显式运行：

```powershell
.\tools\install_word_caption_labels.ps1
```

脚本会先备份当前 `Normal.dotm`，再注册 `图/表/公式`，使 Word 的“插入题注”和
“交叉引用”下拉框与生成文档中的 `SEQ 图`、`SEQ 表`、`SEQ 公式` 保持一致。构建过程
不会静默修改 Word 全局模板；安装后重新打开 Word 即可。

`-WordRefresh auto|always|never` 默认为 `auto`：检测到 Microsoft Word 时刷新全部
StoryRanges、重新分页并处理续表；`always` 在 Word 不可用时失败；`never` 只保留
`w:updateFields=true`，适合调试中间结构，不适合作为含跨页长表的最终交付件。

## PDF 与 Word 的边界

PDF 是版式主输出。DOCX 用于继续编辑，不承诺与 PDF 逐页像素一致。

Word 构建并不是把 Markdown 当作最终稿。转换器只在临时目录中生成 Pandoc 可读的
正文语义中间表示，再生成 DOCX；封面、声明和正文格式优先复用 `reference.docx`
中的命名 Word 样式，目录、分节、图片关系、题注域和交叉引用再由 python-docx/OOXML
补齐。封面九个填写槽统一使用 `CoverValueLine` 段落样式的单一底边框，不使用字符
下划线、NBSP 填充或 tab leader。默认不会保留 `.pandoc.md`，只有显式传入
`--keep-intermediate` 时才会输出它供调试。

Word 转换支持本项目受控子集：标准章节、普通段落、常见表格/图片/公式、引用、
集中注释和示例自定义命令。复杂 TikZ、PSTricks、任意低层宏、复杂浮动体或手工
分页可能需要在 Word 中调整。

转换时会解析图、表、公式和 `\ref`/`\eqref` 的全局编号。图题和表题是继承 Word
内置 `Caption` 的论文样式，编号使用真实 `SEQ` 域并由唯一书签包围；正文图表引用使用
可点击的 `REF` 域。目录、页码和文献引用同样保留 Word 域，最终构建会在 Word 可用时
刷新。LaTeX 仍是内容与编号关系的源文件，结构性增删建议回到 LaTeX 后重新转换。

安装 Microsoft Word 的 Windows 电脑可自动更新域并可选导出检查 PDF：

```powershell
.\tools\refresh_word.ps1 `
  -InputPath build/word/main.docx `
  -PdfOutput build/qa/word/main.pdf
```

## 回归案例

`examples/full-featured/` 是与任何具体毕业论文相互独立的排版夹具，用于测试长题名、
双语摘要、science 一至四级标题与编号、外部 PNG、三线表、公式、交叉引用、注释和
参考文献。示例只描述文档转换与格式校验流程，不包含真实作者信息、研究数据或实证结论。

`tests/` 还包含 humanities/science 编译烟测，覆盖跨章点号编号、无编号结论、
目录写入和交叉引用。
`tests/audit_format.py` 会直接审计 LaTeX 类与 Word 样式中的宋体/黑体/楷体/
Times New Roman、字号、1.5 倍行距、中文正文及标题后首段的两字缩进、英文摘要顶格、页边距、
页眉页脚距离、封面填写槽、Caption/SEQ/REF/书签和重复表头。表格单元格的宽度与
对齐由 LaTeX 列规格驱动，不以独立内容词表或字符数阈值限制作者写法；`build.ps1 check`
会把结构审计与两项烟测一次执行完。

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
