# md2word

这是一个用于分享 `.claude` 自定义 skill 的仓库。它提供了一个名为 `md2word` 的 skill，用于在 Windows 环境下把 Markdown 稳定转换为 Word 文档，并支持模板复用、Markdown 图片/表格、标题导航与公式后处理。

核心链路很直接：

`Markdown -> MHT -> DOCX -> WordMath DOCX`

## 当前能力

- 模板分流：支持直接使用 `.mht` 模板、先把 `.docx` 模板归一化成 `.mht`，或根据模板规格 / 文字描述生成可复用模板。
- 正文渲染：支持普通段落、标题、有序和无序列表、代码块，以及将 Markdown 内容插入模板正文区域。
- 元数据填充：支持 Markdown 头字段，既可以渲染成信息表，也可以填充模板中的命名占位符和封面表格。
- 图片与图表：支持本地 Markdown 图片打包进 MHT；支持 Mermaid 代码块渲染为图片后插入正文。
- 表格：支持标准 Markdown 表格，并在 Word 中保留基础表格结构和左右居中对齐信息。
- 标题导航：支持把 Markdown 正文标题映射为 Word Heading 样式和导航窗格层级。
- 公式处理：支持行内 / 行间 LaTeX 公式，导出后继续交给 Word 转成 Office Math；已覆盖常见矩阵和 `cases` 等块公式归一化。
- 样式预设：内置 `default` 和 `academic-paper` 两套样式，可同时影响模板生成和正文渲染。


## 适合什么场景

- 你已经有 `.mht` 模板，想把 Markdown 内容灌进去
- 你只有 `.docx` 模板，想复用已有 Word 样式
- 你没有模板文件，只有一段模板描述，希望先生成模板再导出
- 你的 Markdown 里有本地图片、标准表格，且希望在 Word 中保留结构
- 你的 Markdown 里带有 `$...$` 或 `$$...$$` LaTeX 公式，想交给 Word 转成 Office Math

## 仓库结构

实际 skill 内容放在隐藏目录里：

```text
.claude/
└── skills/
    └── md2word/
        ├── SKILL.md
        ├── agents/
        ├── resources/
        └── scripts/
```

其中：

- `.claude/skills/md2word/SKILL.md`：skill 入口说明
- `.claude/skills/md2word/scripts/`：PowerShell 与 Python 脚本
- `.claude/skills/md2word/resources/`：默认模板、样式预设与模板规格参考

## 环境要求

这个 skill 默认面向 Windows，依赖也不隐藏：

- Windows
- PowerShell
- 桌面版 Microsoft Word，且 COM 可用
- Python 3
- 如果 Markdown 中包含 Mermaid，需本机可执行 `mmdc`，或允许通过 `npx` 调起 Mermaid CLI

如果这些条件不满足，就不要假设最终 `.docx` 和公式转换还能“自动补齐”。

## 安装方式

把仓库里的 skill 目录复制到你的项目中：

```text
<your-project>/.claude/skills/md2word
```

最简单的做法就是直接复制本仓库的：

```text
.claude/skills/md2word
```

复制完成后，Claude / Codex 在匹配到相关任务时就可以加载这个 skill。

## 使用方式

默认入口是：

```powershell
.claude\skills\md2word\scripts\md2word.cmd --input disclosure.md
```

使用 `.mht` 模板：

```powershell
.claude\skills\md2word\scripts\md2word.cmd `
  --input disclosure.md `
  --template-mht custom-template.mht `
  --template-report template_inference_report.md
```

使用 `.docx` 模板：

```powershell
.claude\skills\md2word\scripts\md2word.cmd `
  --input disclosure.md `
  --template-docx contract-template.docx `
  --template-out contract-template.normalized.mht
```

根据模板描述先生成模板再转换：

```powershell
.claude\skills\md2word\scripts\md2word.cmd `
  --input disclosure.md `
  --template-spec template-spec.md `
  --template-out generated-template.mht
```

## 产物说明

常见输出包括：

- `<input>.mht`
- `<input>.docx`
- `<input>.wordmath.docx`
- `--template-out` 指定的归一化模板或生成模板
- `--template-report` 指定的模板推断报告

## 这个 skill 做了什么

- 根据模板来源在 `.mht`、`.docx`、模板描述三种输入之间分流
- 将 Markdown 渲染进模板，支持正文替换、占位符替换、元数据表填充、Markdown 图片和表格
- 调用 Windows Word 将 MHT 另存为 DOCX
- 在 Word 中把 Markdown 标题映射为 Word 导航层级，并把 LaTeX 文本进一步转换为 Office Math

## 输入约定

- 头字段支持 `**标题：** 示例文档` 和 `标题: 示例文档` 两种写法。
- 标题优先从 `标题` / `题目` / `主题` 取值，没有时再回退到 `发明名称` / `项目名称`。
- 行内公式保持 `$...$`，行间公式保持 `$$...$$`。
- 本地图片使用 `![说明](relative/path.png)`。
- Markdown 字段区后面即使紧跟空行或 `# 标题`，正文也能被正确识别。

## 限制

- 不是跨平台方案，核心链路依赖 Windows Word COM。
- Mermaid 渲染依赖本机 `mmdc`，或允许通过 `npx` 调起 Mermaid CLI。
- 模板没有显式正文锚点时，会走 section 推断逻辑；复杂模板建议同时输出 `--template-report` 检查。

## 致谢

- https://linux.do/t/topic/1217729感谢这篇文章提供的思路
- 没有Linux Do，确实没有这个skills AWA

## License

MIT，见 [LICENSE](./LICENSE)。
