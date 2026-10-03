# 外部项目构建配置

论文正文、开题报告和文献综述都可以作为 YibinThesis 仓库之外的独立消费项目。
每个项目只维护一个 `yibinthesis.project.json` 和一个 `main` 入口，不需要复制模板
仓库的 `build.ps1`。

CLI 会把配置文件所在目录当作项目根目录。建议把源文件集中放在 `latex/`，把构建
产物集中放在 `build/`；这样项目可以独立提交到自己的仓库，模板仓库只作为工具依赖。

## 推荐目录

```text
my-document/
├── yibinthesis.project.json
├── latex/
│   ├── main.tex
│   ├── metadata.tex
│   ├── references.bib
│   ├── chapters/
│   └── assets/
└── build/
```

配置示例：

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
    "pdf": "build/deliverables/document.pdf",
    "word": "build/deliverables/document.docx"
  }
}
```

`$schema` 只为编辑器提供补全和校验，应按模板仓库与消费项目的实际相对位置调整。
配置中所有相对路径都以 `yibinthesis.project.json` 所在目录为基准。

## 字段

| 字段 | 必需 | 说明 |
|---|---|---|
| `schemaVersion` | 是 | 当前固定为 `1`。 |
| `main` | 是 | 唯一 LaTeX 入口；文档类型由此文件的 `\documentclass` 选项确定。 |
| `outputRoot` | 否 | `pdf`、`word` 和临时产物的构建根目录，默认使用模板仓库的 `build`。 |
| `checkOutputRoot` | 否 | `check` 命令的独立构建根目录；省略时使用 `outputRoot`。 |
| `citationMode` | 否 | `linked` 或 `native`，默认 `linked`。 |
| `wordRefresh` | 否 | `auto`、`always` 或 `never`，默认 `auto`。 |
| `deliverables.pdf` | 否 | 成功构建后发布的稳定 PDF 文件名，扩展名必须为 `.pdf`；文件名可使用 `{{title}}`、`{{author}}` 等元数据占位符。 |
| `deliverables.word` | 否 | 成功构建后发布的稳定 DOCX 文件名，扩展名必须为 `.docx`；文件名可使用 `{{title}}`、`{{author}}` 等元数据占位符。 |

交付文件名中的 `{{field}}` 会从入口文件同目录的 `metadata.tex` 读取。可用字段与
元数据键一致，例如 `title`、`english-title`、`author`、`student-id`、`college`、
`major`、`grade`、`class-name`、`advisor`、`template-year`、`date` 和 `version`。
占位值会经过文件名安全化处理；目录部分仍按配置文件解析。字段缺失或入口旁没有
`metadata.tex` 时，构建会直接失败，不会生成带有未替换占位符的交付件。

文档类型和 `template-year` 不在 JSON 中重复配置。前者写在 `main.tex` 的类选项中，
后者写在 `metadata.tex` 的 `\yibinsetup` 中。年份必须按文档类型分别填写：

- `thesis`：`template-year = 2024`；
- `proposal`：`template-year = 2022`，对应 2022 工作表物理第 5 至 7 页；
- `literature-review`：`template-year = 2024`，对应 2024 论文正文版式基线。

类型与年份不匹配时构建会停止。当前不支持 2025 论文模板。

## 调用方式

推荐通过仓库提供的 CLI 调用。CLI 不把模板仓库当作论文项目，默认在当前目录发现
`yibinthesis.project.json`：

```powershell
yibinthesis new ..\论文项目\我的论文 --type thesis --discipline humanities
yibinthesis doctor
yibinthesis build all
yibinthesis clean
```

`new`/`init` 只在工具仓库之外创建消费项目，并以 UTF-8 写入 LaTeX 源文件。`build`
的格式位置参数为 `pdf`、`word` 或 `all`，省略时使用 `all`。`--project` 与 `--config`
互斥；`--config` 指向配置文件时，项目根自动取其所在目录。`--dry-run` 可用于预览
脚手架文件计划，`--force` 只允许覆盖 CLI 自己标记过的脚手架文件。

需要在没有安装 Python 的 Windows 机器上运行时，可用 `build_exe.ps1` 生成 PyInstaller
发行版。默认生成 onedir 目录，传入 `-OneFile` 才生成单文件 exe；两种发行版都只
封装 CLI 和模板资源，XeLaTeX/Tectonic、Biber、Pandoc 与 Word 仍属于运行时依赖。

在消费项目根目录执行时，构建器会自动发现当前目录下的
`yibinthesis.project.json`：

```powershell
& 'X:\path\to\YibinThesis\build.ps1' doctor
& 'X:\path\to\YibinThesis\build.ps1' check
& 'X:\path\to\YibinThesis\build.ps1' all
```

也可以从任意目录显式指定配置：

```powershell
& 'X:\path\to\YibinThesis\build.ps1' all `
  -Config 'X:\path\to\my-document\yibinthesis.project.json'
```

显式传入的 `-Main`、`-OutputRoot`、`-CitationMode` 和 `-WordRefresh` 优先于配置文件。
`doctor` 检查工具链，`check` 运行结构和格式审计，`pdf`、`word`、`all` 分别构建
对应交付格式，`clean` 只清理所选入口的构建产物。

CLI 的退出码约定为：`0` 表示命令完成，`2` 表示用户输入、项目配置或运行时依赖不满足，
`130` 表示用户中断。构建失败会保留失败现场，便于查看 `build/` 下的日志；再次执行
`clean` 才会删除本项目可再生的构建输出。

## Word 引用模式

| 值 | 行为 | 适用场景 |
|---|---|---|
| `linked` | 上标编号使用参考文献书签和 `REF` 域，可点击跳转 | 默认交付，优先保持 GB/T 7714 数字版式稳定 |
| `native` | 在稳定列表之外写入 Word `CITATION` 域和 bibliography `customXml` | 需要在 Word 中继续使用原生引文来源 |

`native` 不会在构建时静默安装或修改 Word 的全局参考文献样式。需要项目附带的数字
样式时，用户必须显式运行 `tools/install_word_bibliography_style.ps1`。

## Word 刷新模式

| 值 | 行为 |
|---|---|
| `auto` | 检测到 Microsoft Word 时更新所有 StoryRanges 中的域、重新分页并处理续表；未检测到时保留 `w:updateFields=true` 并警告。 |
| `always` | 强制要求 Microsoft Word 完成刷新；Word 不可用时构建失败。 |
| `never` | 跳过 COM 刷新，只保留待更新域，适合检查中间结构。 |

最终 Word 交付，尤其是含跨页长表的文档，建议使用 `wordRefresh: "always"`。WPS 和
LibreOffice 可作兼容性预览，但当前 Word 分页与域刷新以 Microsoft Word 为验收实现。

## 多份材料

一个配置文件只指向一个入口。论文、开题报告和文献综述应各自放在独立项目目录中，
分别维护配置和交付文件名。例如：

```text
manuscripts/
├── thesis/yibinthesis.project.json
├── proposal/yibinthesis.project.json
└── literature-review/yibinthesis.project.json
```

这样可以共享同一套模板与工具链，同时避免每个项目维护一份会逐渐漂移的构建脚本。
