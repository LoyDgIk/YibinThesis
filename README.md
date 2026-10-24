# YibinThesis

YibinThesis 是面向宜宾学院本科毕业材料的非官方写作与排版工具，支持正式论文、
文献综述和开题报告的独立项目创建、PDF/DOCX 构建与检查。仓库还提供供 AI 按需读取的
写作技能，将内容写作要求与模板排版职责分开。

仓库维护模板、转换器、检查器、CLI、写作技能与测试示例，不保存用户论文稿件。
论文源文件由 `new` 命令创建到仓库之外的独立目录。

版式与内容规则依据宜宾学院本科毕业论文（设计）撰写规范及相应文种模板整理，
支持范围见下文。使用时不需要另行取得规范原件；原始学校文件和已填写的个人样例不随
发行版提供。学校、学部或专业后续发布的适用要求优先，提交前仍需由指导教师或所在学部复核。

## 主要功能

- 三类材料独立建项，正文、元数据、文献库和图件与工具仓库分离。
- 从 LaTeX 源文件生成 PDF 和可继续编辑的 DOCX，提供配置化交付文件名与环境检查。
- 参考文献列表的 URL 和 DOI 可通过 `no-url`、`no-doi` 类选项分别隐藏，设置同时作用于 PDF 与 Word。
- `new` 创建的三类项目默认加入 `no-doi`，URL 仍保留；需要显示 DOI 时删除该选项即可。
- Word 默认使用可跳转的数字引文，也可生成原生引用来源数据；中英文作者省略词与数字版次按条目语言处理。
- AI 写作指南按通用要求、正式论文、文献综述、开题报告、LaTeX 输入及公式分别组织，按任务读取。

## 安装与调用

### Windows 发行版

从 [GitHub Releases](https://github.com/LoyDgIk/YibinThesis/releases/latest) 下载
`yibinthesis.exe` 与 `SHA256SUMS.txt`。在 PowerShell 中运行：

```powershell
.\yibinthesis.exe --help
.\yibinthesis.exe new ..\my-thesis --type thesis --discipline humanities
.\yibinthesis.exe doctor --project ..\my-thesis
```

exe 可直接用于 CLI 调用，但不是完整的离线排版环境。构建仍按输出类型需要外部
PowerShell、Python 及相应工具链，具体见下文；请先运行 `doctor` 检查。

### 从源码使用

需要 Python 3.11 或更高版本。在仓库根目录中，可直接运行：

```powershell
$env:PYTHONPATH = (Resolve-Path .\lib).Path
python -m yibinthesis_cli --help
```

推荐以可编辑模式安装命令行入口，使其使用当前源码目录中的模板资源：

```powershell
python -m pip install -e .
yibinthesis --help
```

### 构建 Windows 可执行文件

在源码目录安装打包依赖，再使用 `build_exe.ps1`：

```powershell
python -m pip install -e ".[exe]"
.\build_exe.ps1
dist\yibinthesis\yibinthesis.exe --help
# 单文件输出：
.\build_exe.ps1 -OneFile
```

脚本默认生成 onedir 发行目录；`-OneFile` 输出 `dist/yibinthesis.exe`。

仓库内的 `.github/workflows/release.yml` 会在推送 `v*` 标签后自动运行 CLI 测试，构建
单文件 Windows exe，生成 `SHA256SUMS.txt`，并连同 `LICENSE`、`NOTICE` 发布到对应的 GitHub Release。也
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
2024，不能由调用者随意混用。`--dry-run` 只显示将要创建的文件，`--force`
仅允许覆盖脚手架已生成的文件，不会删除目标目录中的其他内容。

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

工具链可通过环境变量显式指定，也会探测 `.tools/`、PATH 和受支持的本机缓存位置。
实际可用工具以 `doctor` 输出为准。常用环境变量有：
`YIBINTHESIS_TECTONIC`、`YIBINTHESIS_BIBER`、`YIBINTHESIS_PANDOC`、
`YIBINTHESIS_PYTHON`、`YIBINTHESIS_LATEXMK` 和 `YIBINTHESIS_XELATEX`。

## 文档类型与版式

| 文档类型 | 类选项 | 年份 | 用途 |
| --- | --- | ---: | --- |
| 毕业论文（设计） | `thesis` | 2024 | 封面、声明、双语摘要、目录、正文、后置部分 |
| 开题报告 | `proposal` | 2024 | 新版封面与七个开题栏目 |
| 文献综述 | `literature-review` | 2024 | 封面、同页双语摘要、连续正文与参考文献 |

两类学科采用各自的章节编号体系，具体层级随文种处理，不需要手写显示编号。综述正文
以无标题引言开始，按主题综合分析，末章为“结论与展望”；不套用正式论文的五节绪论。
开题报告按七个栏目填写，尚未完成的工作按计划表述。

模板处理页面、字体、标题、图表、公式编号及 GB/T 7714-2015 顺序编码引用。
正式论文后置部分由入口文件装配，脚手架采用“致谢→参考文献→附录”的顺序。
不同文种的具体版式与接口见[文档类型与入口](docs/document-types.md)，不以一套
页面或编号参数覆盖三类材料。缺少学校字体时默认使用兼容字体预览；需要严格检查时，
在 `\documentclass` 选项中加入 `strictfonts`。

### Word 引用与刷新

- `linked` 为默认模式：上标数字引文通过书签和 `REF` 域跳转到文后条目。
- `native` 额外写入 Word `CITATION` 域和来源数据；不会自动安装或修改 Word 全局引用样式。
- 4 人及以上作者保留前 3 人，中文显示 `, 等.`，英文显示 `, et al.`；英文数字版次使用 `4th ed.` 等写法。
- 文献库保留完整作者及版次信息，建议显式填写 `langid`；未标注时，Word 构建按题名是否含汉字区分中英文。
- `--word-refresh auto` 检测到 Microsoft Word 时刷新域和分页；`always` 强制要求 Word，`never` 跳过刷新。

DOCX 可编辑不代表其他编辑器的分页与 Word 完全一致。最终交付，尤其是含跨页长表时，
建议在安装 Microsoft Word 的环境使用 `--word-refresh always`。详细行为见
[构建配置](docs/project-config.md#word-引用模式)。

## AI 写作技能

技能入口为 [yibinthesis-writing](skills/yibinthesis-writing/SKILL.md)。在支持技能的 AI
环境中，按该环境的加载方式使用此目录。技能在仓库中独立提供，不应将下载 CLI exe
视为已经为 AI 安装了技能。

| 当前任务 | 按需读取 |
| --- | --- |
| 写作的共用规则与统计口径 | [通用写作要求](skills/yibinthesis-writing/references/writing-requirements.md) |
| 正式论文摘要、绪论、主体、分析讨论与结论 | [正式论文写作](skills/yibinthesis-writing/references/thesis-writing.md) |
| 综述选材、综合比较、导言、主题论述与总结 | [文献综述写作](skills/yibinthesis-writing/references/literature-review-writing.md) |
| 开题计划、研究方法、可行性和进度 | [开题报告写作](skills/yibinthesis-writing/references/proposal-writing.md) |
| 入口、元数据、章节和引文源文件 | [LaTeX 作者接口](skills/yibinthesis-writing/references/latex-writing.md) |
| 公式、变量、单位与公式引用 | [公式编写](skills/yibinthesis-writing/references/equations.md) |

内容写作读取通用要求及当前文种，不默认载入全部指南；只改 LaTeX 语法时仅读作者接口。
技能已内置必要的写作要求，不依赖未发布的学校文件或维护者目录，也不预置个人资料。
AI 负责依据用户材料组织内容，排版和构建由工具处理。导航见[写作指南](docs/writing-guide.md)。

## 构建依赖

| 操作 | 所需环境 |
| --- | --- |
| 源码 CLI | Python 3.11+；构建适配需要 PowerShell |
| PDF 构建 | XeLaTeX/latexmk 或 Tectonic、Biber，以及模板所需 LaTeX 宏包与字体；构建中的 Python 辅助步骤按实际素材使用相应依赖 |
| DOCX 构建 | Python、Pandoc、`requirements-word.txt` 中的 python-docx 与 Pillow |
| DOCX 域刷新与最终分页 | Microsoft Word；跳过刷新可生成待更新域的文档 |

在源码目录安装 Word Python 依赖：

```powershell
python -m pip install -r requirements-word.txt
```

工具链和字体不随单文件 exe 一并提供；先运行 `doctor`，再按所需输出补齐依赖。

## 开发检查

```powershell
python -m compileall -q .
python -m unittest discover -s lib/tests -p "test_*.py" -v
python lib/audit_format.py
python lib/docx_audit.py <template.docx> --json build/docx-audit.json
```

针对用户自行提供的 DOCX 模板做只读结构审计：

```powershell
python lib/docx_audit.py `
  "X:\templates\thesis.docx" `
  "X:\templates\literature-review.docx" `
  --json build/docx-audit
```

该审计只读取 DOCX，不修改学校原件；它报告分节、页面几何、字体字号、段落参数、表格列网格、
单元格边框和包内页眉页脚/批注部件。视觉验收仍需使用 Word 或其他可用的文档渲染工具
逐页检查；仓库不附带 `render_docx.py` 或原始学校模板。

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
├── skills/                   # 按任务读取的 AI 写作与 LaTeX 技能
├── tests/
├── assets/
├── word/
├── tools/
└── docs/
```

## 开源许可

项目自身代码采用 [GNU GPLv3](LICENSE)（`GPL-3.0-only`）。修改和再分发时须遵守
该许可证，发布对应二进制时也须按许可证提供对应源代码。

校名、校徽、校名字样及其他机构标识不包含在项目代码的许可范围内。随附 CSL 文献样式
继续遵循其原有的 [CC BY-SA 3.0 声明](word/CSL-NOTICE.txt)，不改为 GPLv3。
第三方许可、机构标识与独立实现边界见 [NOTICE](NOTICE)。
