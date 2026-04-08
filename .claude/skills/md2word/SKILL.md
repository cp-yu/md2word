---
name: md2word
description: 将 Markdown 在 Windows + Microsoft Word 流水线中稳定转换为 `.mht`、`.docx` 和 `.wordmath.docx`。用于交底书、学术论文、技术报告等 Markdown→Word 导出场景，支持 `.mht` 模板、`.docx` 模板、模板规格文件、普通或加粗头字段、Markdown 图片、表格、Mermaid 渲染、样式预设，以及需要把 LaTeX 公式交给 Word 转 Office Math 的任务。
---

# md2word

目标：

`Markdown -> MHT -> DOCX -> WordMath DOCX`

这个 skill 默认面向 Windows 直接运行。内置模板只是 fallback，不要把它理解成“只能导出专利交底书”。

## 入口判断

- 用户已经给了 `.mht` 模板
- 用户已经给了 `.docx` 模板，希望复用现有 Word 样式
- 用户没有模板文件，只有一段“模板长什么样”的文字描述
- 用户要求论文格式、标题字号、行距、封面结构等版式约束
- Markdown 里有 `$...$`、`$$...$$`、Mermaid 代码块，且最终必须落到 Word

## 环境要求

- 在 Windows 中运行
- PowerShell 可用
- Windows 桌面版 Microsoft Word 已安装，且 COM 可用
- Python 3 可通过 `py -3`、`python` 或 `python3` 调起
- 如果 Markdown 含 Mermaid：
  - 优先复用本机已有 `mmdc`
  - 否则允许 `npx -p @mermaid-js/mermaid-cli mmdc`
  - 首次 `npx` 兜底可能需要联网下载

## 输入约定

- 行内公式保持 `$...$`
- 行间公式保持 `$$...$$`
- 本地图片使用标准 Markdown 语法：`![说明](relative/path.png)`
- 标准 Markdown 表格会被渲染成 Word 表格
- Markdown 头字段支持两种写法：
  - `**标题：** 示例文档`
  - `标题: 示例文档`
- 如果字段区后紧跟空行或 `# 标题`，正文也会被正确识别，不要求一定写 `---`
- 标题优先从 `标题` / `题目` / `主题` 取值；没有时再回退到 `发明名称` / `项目名称`
- 正文 Markdown 标题会映射为 Word 正文标题层级，可在导航窗格中看到
- 默认只有正文章节标题进入导航；封面或文档主标题不进入导航
- 已显式写出的标题编号会保留；只有无编号标题才会做最小自动补号

## 工作流

1. 按模板来源三选一分流：
   - `.mht`：直接使用
   - `.docx`：先归一化导出为 `.mht`
   - 模板规格/文字描述：先生成可复用 `.mht`
2. 用 `render_mht.py` 把 Markdown 填进模板。
3. 调 Windows Word 打开 MHT，另存成 `.docx`。
4. 在 Word 内把 LaTeX 文本转换成 Office Math，输出 `.wordmath.docx`。

## 模板与样式

- 如果用户给模板文件，优先复用模板，不要重造
- 如果用户只给一句模板描述，先写成 `template-spec.md` 再传给 `--template-spec`
- 模板规格参考：`resources/template-spec.md`
- 当前内置样式预设：
  - `default`
  - `academic-paper`
- 静态模板、模板规格和样式预设统一放在 `resources/`
- 用户明确给出论文标题字号、标题层级、正文行距时，优先：
  - 在模板规格里写 `style-preset: academic-paper`
  - 调用时再显式传 `--style-preset academic-paper`

## 正文插入策略

- 如果模板里有 `{{CONTENT}}` 或 `<!--MD_CONTENT-->`，优先按占位符替换
- 如果模板里有 `{{METADATA_TABLE}}`，会把 Markdown 头字段渲染成信息表
- 如果模板里已有表格标签，例如 `项目名称`、`电话`、`版本`，会按行标签尝试填值
- 如果没有显式占位符，会推断 `WordSection`，优先替换最后一个正文 section
- 推断过程可通过 `--template-report` 输出报告

## 命令

使用内置默认模板：

```powershell
<skill-dir>\scripts\md2word.cmd `
  --input disclosure.md
```

使用模板规格生成模板：

```powershell
<skill-dir>\scripts\md2word.cmd `
  --input example.md `
  --template-spec template-spec.md `
  --template-out generated-template.mht `
  --template-report template-report.md
```

生成学术论文格式：

```powershell
<skill-dir>\scripts\md2word.cmd `
  --input example.md `
  --template-spec template-spec.md `
  --template-out academic-paper.mht `
  --template-report academic-paper.report.md `
  --style-preset academic-paper
```

使用现成 Word 宏而不是内置公式转换：

```powershell
<skill-dir>\scripts\md2word.cmd `
  --input disclosure.md `
  --macro-name "LatexToWordMath_Better"
```

直接调用 Python 主脚本：

```powershell
py -3 <skill-dir>\scripts\md2word.py `
  --input disclosure.md `
  --style-preset academic-paper
```

兼容旧的 PowerShell 包装入口：

```powershell
powershell -ExecutionPolicy Bypass -File <skill-dir>\scripts\md2word.ps1 `
  -Input disclosure.md `
  -StylePreset academic-paper
```

## 公式与 Mermaid

- `render_mht.py` 会把 Markdown 图片作为 MHT 关联资源打包进去
- 标准 Markdown 表格会转成正文 Word 表格，并保留基本对齐方式
- `render_mht.py` 会把 Mermaid 代码块渲染成图片插入 MHT
- 当前 Mermaid 渲染默认写入保守的 Puppeteer 配置，规避 `headless shell` 在受限环境下失败
- `md2word.py` 会在生成的 DOCX 中把 Markdown 标题映射回 Word Heading 样式和大纲级别
- `word_mht_pipeline.ps1` 会把常见块公式环境归一化后再交给 Word `BuildUp`
- 当前已覆盖这些块环境：
  - `matrix`
  - `pmatrix`
  - `bmatrix`
  - `Bmatrix`
  - `vmatrix`
  - `Vmatrix`
  - `cases`
- Word 不支持把 `\begin{matrix}...\end{matrix}` 原样直接 BuildUp；要先转成 `\matrix{...}` 或外层加分隔符

## 附带文件

- `scripts/md2word.py`：Python 主入口，统一编排模板分流、渲染和 Word 流水线
- `scripts/md2word.cmd`：Windows 命令行包装器，转发到 Python 主入口
- `scripts/md2word.ps1`：PowerShell 兼容包装器，转发到 Python 主入口
- `scripts/pipeline_common.py`：主入口共享的参数校验、路径规划和文件校验逻辑
- `scripts/render_mht.py`：渲染 MHT、处理正文和 Mermaid
- `scripts/generate_template_mht.py`：从模板规格生成可复用 MHT 模板
- `scripts/style_presets.py`：统一加载样式预设配置
- `scripts/word_common.ps1`：Word worker 共享的 COM 清理与路径 helper
- `scripts/word_template_to_mht.ps1`：DOCX 模板归一化为 MHT
- `scripts/word_mht_pipeline.ps1`：Word 导出与 Office Math 转换
- `resources/template-spec.md`：模板规格参考
- `resources/style-presets.json`：样式预设配置
- `resources/专利交底书模板.mht`：默认 fallback 模板

## 失败处理

- 如果 PowerShell、Python 或 Word COM 不可用，不要假装 Windows 之外能补全最终 Word 效果，直接报环境问题
- 如果 Mermaid 渲染失败，先看是否有本机 `mmdc`；没有的话确认是否允许 `npx` 下载
- 如果模板效果不对，先看 `--template-report`
- 如果公式异常，先区分：
  - 原始 `.docx` 已经错：问题在 MHT 或公式文本归一化
  - 原始 `.docx` 正常、`.wordmath.docx` 才错：问题在 Word `BuildUp`
- 如果矩阵类公式再次出问题，先检查传给 Word 的线性格式是否已经被归一化成 `\matrix{...}` 族，而不是残留 `\begin{...}` / `\end{...}`
