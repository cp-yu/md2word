# md2word

这是一个用于分享 `.claude` 自定义 skill 的仓库。它提供了一个名为 `md2word` 的 skill，用于在 Windows 环境下把 Markdown 稳定转换为 Word 文档，并支持模板复用、Markdown 图片/表格、标题导航与公式后处理。

核心链路很直接：

`Markdown -> MHT -> DOCX -> WordMath DOCX`

这个仓库的重点不是“演示模板”，而是把一套可复用的 `.claude/skills/md2word` 能力独立出来，方便放进你自己的 Claude / Codex 工作流里。

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

## 限制

- 不是跨平台方案，核心链路依赖 Windows Word COM
- 默认模板只是示例，不代表只能处理专利用文档
- 如果模板没有显式锚点，正文插入会走推断逻辑，因此复杂模板建议配合 `--template-report` 检查
- Mermaid 渲染依赖额外工具，不会静默降级

## 适合拿去怎么分享

这是一个典型的 skill 分发仓库：

- 仓库根目录保留 `README.md` 供人看
- skill 本体放在 `.claude/skills/md2word`
- 具体行为、命令和边界写在 `SKILL.md`

这样仓库既能直接分享，也能直接复制进项目里使用。

## 致谢

- https://linux.do/t/topic/1217729感谢这篇文章提供的思路
- 没有Linux Do，确实没有这个skills AWA

## License

MIT，见 [LICENSE](./LICENSE)。
