# YibinThesis CLI

YibinThesis 是一个面向宜宾学院本科毕业材料的非官方 LaTeX 排版工具。工具仓库只
维护文档类、转换器、检查器和命令行程序，不保存任何论文稿件。论文源文件应由
`new` 命令创建到工具仓库之外的独立目录。

版式依据为 `01_规范模板/最新要求与模板` 中的 2024 人文社科/理工农医模板，以及
宜宾学院本科毕业论文（设计）撰写规范。学校、学院或专业的新通知优先于本项目，
提交前仍需由指导教师或所在学院复核。

## 安装与调用

需要 Python 3.11 或更高版本。未安装为包时，可直接从源码运行：

```powershell
$env:PYTHONPATH = (Resolve-Path .\lib).Path
python -m yibinthesis_cli --help
```

安装为命令行程序：

```powershell
python -m pip install .
yibinthesis --help
```

Windows 发行版可以用 `build_exe.ps1` 生成 PyInstaller 可执行文件：

```powershell
.\build_exe.ps1
dist\yibinthesis\yibinthesis.exe --help
```

脚本默认生成 onedir 发行目录，也支持 `-OneFile` 生成单文件 exe。exe 仍需要外部
LaTeX/Pandoc/Word 工具链才能完成相应的 PDF 或 DOCX 构建。

仓库内的 `.github/workflows/release.yml` 会在推送 `v*` 标签后自动运行 CLI 测试，构建
单文件 Windows exe，生成 `SHA256SUMS.txt`，并将二者发布到对应的 GitHub Release。也
可以从 Actions 页面手动运行，并填写已有标签；发布所需的 `contents: write` 权限已在
工作流中声明。

## CLI 接口

命令的第一个位置参数是动作，选项名称保持一致，适合手工使用和脚本调用：

```text
yibinthesis new PROJECT [options]
yibinthesis build [pdf|word|all] [--project PROJECT | --config CONFIG]
yibinthesis doctor [--project PROJECT | --config CONFIG]
yibinthesis check [--project PROJECT | --config CONFIG]
yibinthesis clean [--project PROJECT | --config CONFIG]
yibinthesis version
```

`init` 是 `new` 的兼容别名。`build` 默认使用 `all`，项目默认从当前目录发现
`yibinthesis.project.json`。`--project` 指向项目目录，`--config` 指向明确的配置文件；
两者不能同时使用。`--main`、`--output-root`、`--citation-mode` 和 `--word-refresh`
只用于覆盖本次构建的配置。

`new` 会在指定目录创建独立的 LaTeX 项目，不会复制 `build.ps1`、模板类或工具源码：

```powershell
yibinthesis new ..\论文项目\我的论文 `
  --type thesis `
  --discipline humanities `
  --title "论文中文题目" `
  --author "姓名" `
  --student-id "学号" `
  --college "学院" `
  --major "专业" `
  --grade 2023 `
  --advisor "导师" `
  --advisor-title "职称"
```

生成目录包含 `latex/main.tex`、`latex/metadata.tex`、`latex/references.bib`、
`latex/chapters/`、`latex/assets/`、`build/` 和 `yibinthesis.project.json`。`proposal` 和
`literature-review` 会生成对应的专用入口和字段；`template-year` 由文档类型固定为
2022 或 2024，不能由调用者随意混用。`--dry-run` 只显示将要创建的文件，`--force`
仅允许覆盖脚手架已生成的文件，不会删除目标目录中的其他内容。

CLI 实现按职责拆分在 `lib/yibinthesis_cli/`：参数调度、脚手架生成、构建器适配、运行时
路径与 UTF-8 文件读写分别位于独立模块。

工具仓库自身拒绝作为 `new` 的目标目录，也没有根 `main.tex`、`metadata.tex` 或
`chapters/` 写作入口。`examples/` 仅用于回归测试和格式夹具，不是论文工作区。

## 构建外部项目

在生成的项目目录中：

```powershell
yibinthesis doctor
yibinthesis check
yibinthesis build pdf
yibinthesis build word --word-refresh always
yibinthesis build all
yibinthesis clean
```

也可以从任意目录指定配置：

```powershell
yibinthesis build all --config X:\论文项目\我的论文\yibinthesis.project.json
```

PDF 是版式主输出，DOCX 用于继续编辑。构建器会从入口目录解析章节、图件、元数据和
文献，不会回退到工具仓库中的同名文件。具体配置字段见
[docs/project-config.md](docs/project-config.md)。

交付文件名支持引用 `metadata.tex`：

```json
{
  "deliverables": {
    "pdf": "build/deliverables/{{title}}-{{author}}.pdf",
    "word": "build/deliverables/{{title}}-{{author}}.docx"
  }
}
```

占位字段会在构建时读取并安全化；字段缺失时构建失败，避免发布出错名文件。

工具链探测顺序为环境变量、项目 `.tools/`、PATH 和已知本机缓存。常用环境变量有：
`YIBINTHESIS_TECTONIC`、`YIBINTHESIS_BIBER`、`YIBINTHESIS_PANDOC`、
`YIBINTHESIS_PYTHON`、`YIBINTHESIS_LATEXMK` 和 `YIBINTHESIS_XELATEX`。

## 文档类型与版式

| 文档类型 | 类选项 | 年份 | 用途 |
| --- | --- | ---: | --- |
| 毕业论文（设计） | `thesis` | 2024 | 封面、声明、双语摘要、目录、正文、后置部分 |
| 开题报告 | `proposal` | 2022 | 工作表中的七个开题栏目 |
| 文献综述 | `literature-review` | 2024 | 独立信息页与 2024 正文版式 |

`humanities` 使用“一、/（一）/1./（1）”编号，`science` 使用“1/1.1/1.1.1/1.1.1.1”。
模板提供 A4 单面、左 3 cm 及其余 2.5 cm 页边距、宋体/黑体/楷体与 Times New Roman
字体层级、1.5 倍行距、首行缩进两字、三线表、按章图表公式编号、GB/T 7714-2015
顺序编码引用，以及“致谢→参考文献→附录”的后置顺序。缺少学校字体时默认使用兼容
字体预览，最终提交可在 `\\documentclass` 中加入 `strictfonts` 强制检查。

## 开发检查

```powershell
python -m compileall -q .
python -m unittest discover -s lib/tests -p "test_*.py" -v
python lib/audit_format.py
python lib/docx_audit.py <template.docx> --json build/docx-audit.json
```

完整 PDF/Word 检查需要安装 XeLaTeX 或 Tectonic、Biber、Pandoc、python-docx、Pillow，
以及在需要刷新域和分页时安装 Microsoft Word。

针对学校模板的只读结构审计：

```powershell
python lib/docx_audit.py `
  "..\01_规范模板\最新要求与模板\质检学部-文献综述模板（工科类）2024（模板）.docx" `
  "..\01_规范模板\最新要求与模板\质检学部-文献综述模板（文科类）2024（模板）.docx" `
  "..\01_规范模板\最新要求与模板\罗殷开题报告.docx" `
  --json build/docx-audit
```

该审计只读取 DOCX，不修改学校原件；它报告分节、页面几何、字体字号、段落参数、表格列网格、
单元格边框和包内页眉页脚/批注部件。视觉验收仍需在安装 Word 或 LibreOffice 的环境中用
`render_docx.py` 渲染并逐页检查。

## 目录

```text
YibinThesis/
├── lib/                    # 所有生产 Python 模块
│   ├── yibinthesis_cli/
│   ├── build_word.py
│   ├── build_reference_docx.py
│   ├── prepare_signature_assets.py
│   ├── docx_audit.py
│   └── audit_format.py
├── yibinthesis.cls
├── yibinthesis-proposal.sty
├── yibinthesis-literature-review.sty
├── yibinthesis.project.schema.json
├── build.ps1
├── build_exe.ps1
├── examples/                 # 只读回归夹具
├── tests/
├── assets/
├── word/
├── tools/
└── docs/
```

模板源代码采用 MIT License。校名、校徽、校名字样及其他机构标识不包含在 MIT 授权
范围内，参考项目和独立实现边界见 `NOTICE`。
