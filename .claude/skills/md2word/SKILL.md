---
name: md2word
description: 将 Markdown 在 Windows 环境下稳定转换为 `.mht`、`.docx` 和公式处理后的 `.wordmath.docx`。适用于直接运行 PowerShell、已安装桌面版 Microsoft Word，并支持 `.mht` 模板、`.docx` 模板，以及从模板规格或文字描述生成模板。
---

# md2word

这个 skill 的目标不是“只服务交底书模板”，而是把任意 Markdown 文档稳定塞进 Word 可消费的模板链里，并且默认面向 Windows 直接运行：

`Markdown -> MHT -> DOCX -> WordMath DOCX`

默认仍然带一个专利用的示例模板，但那只是 fallback，不是能力边界。

## 适用场景

- 用户已经提供了 `.mht` 模板
- 用户只有 `.docx` 模板，希望复用现有 Word 样式
- 用户没有模板文件，只有一段“模板长什么样”的文字描述
- Markdown 中带有 `$...$` 和 `$$...$$` 的 LaTeX，必须原样保留，后续再交给 Word 转公式

## 要求

- 在 Windows 中运行
- PowerShell 可用
- Windows 桌面版 Microsoft Word 已安装，且 COM 可用
- Python 3 可通过 `py -3`、`python` 或 `python3` 调起
- 行内公式保持 `$...$`
- 行间公式保持 `$$...$$`
- 如果 Markdown 里包含 Mermaid 代码块，在 Codex / Claude Code 环境里可直接复用已有 `Node.js + npx`；脱离该环境时，需要本机可执行 `mmdc`，或允许 `npx -p @mermaid-js/mermaid-cli mmdc`
- 如果走 `npx` 兜底，首次 Mermaid 渲染可能联网下载 `@mermaid-js/mermaid-cli`

## 工作方式

1. 先在入口脚本里检查 Markdown 中公式分隔符是否完整，优先拦截未闭合的 `$...$` 和 `$$`。
2. 根据模板来源做三选一分流：
   - `.mht` 直接使用
   - `.docx` 先归一化导出为 `.mht`
   - 模板规格/文字描述先生成可复用 `.mht`
3. 用 `render_mht.py` 把 Markdown 填进模板。
4. 调 Windows Word 打开该 MHT，另存成 `.docx`。
5. 在 Word 内把 LaTeX 文本转换成 Office Math，输出 `.wordmath.docx`。

## 正文插入策略

这个 skill 默认按“自动猜测”工作，不要求模板事先埋锚点。

- 如果模板里有 `{{CONTENT}}` 或 `<!--MD_CONTENT-->`，优先按占位符替换
- 如果模板里有 `{{METADATA_TABLE}}`，会把 Markdown 头部字段自动渲染成信息表
- 如果模板里已有表格标签，例如 `项目名称`、`电话`、`版本`，会尝试按行标签自动填值
- 如果没有显式占位符，会推断 `WordSection`，优先替换最后一个正文 section
- 推断过程可通过 `-TemplateReport` 输出报告

## 命令

使用 `.cmd` 包装器和内置默认模板：

```powershell
<skill-dir>\scripts\md2word.cmd `
  -Input results\repelMouse\交底书.md
```

使用用户提供的 `.mht` 模板：

```powershell
<skill-dir>\scripts\md2word.cmd `
  -Input disclosure.md `
  -TemplateMht custom-template.mht `
  -TemplateReport template_inference_report.md
```

使用用户提供的 `.docx` 模板：

```powershell
<skill-dir>\scripts\md2word.cmd `
  -Input disclosure.md `
  -TemplateDocx contract-template.docx `
  -TemplateOut contract-template.normalized.mht `
  -TemplateReport template_inference_report.md
```

用户只有文字描述时，先在当前项目里写一个 `template-spec.md`，再生成模板并转换：

```powershell
<skill-dir>\scripts\md2word.cmd `
  -Input disclosure.md `
  -TemplateSpec template-spec.md `
  -TemplateOut generated-template.mht `
  -TemplateReport template_inference_report.md
```

如果用户明确要求调用现成 Word 宏，而不是内置公式转换逻辑：

```powershell
<skill-dir>\scripts\md2word.cmd `
  -Input disclosure.md `
  -MacroName "LatexToWordMath_Better"
```

如果需要直接调用 PowerShell 主脚本：

```powershell
powershell -ExecutionPolicy Bypass -File <skill-dir>\scripts\md2word.ps1 -Input disclosure.md
```

## 模板规格

`-TemplateSpec` 接收的是一个本地文本文件。它可以是：

- 简单键值规格
- 或者纯文字描述

参考格式见：

- `references/template-spec.md`

如果用户只给一句话，例如“帮我做成带封面、正文、文档信息表的技术报告模板”，先把这句话写进一个 `template-spec.md`，再把该文件传给 `-TemplateSpec`。首版目标是结构可复用，不追求像素级还原。

## 输出

- `<input>.mht`: 已填充内容的单文件网页
- `<input>.docx`: Word 原始另存结果
- `<input>.wordmath.docx`: Word 公式处理后的结果
- `-TemplateOut`: 归一化或生成后的可复用模板
- `-TemplateReport`: 模板推断报告

## 附带文件

- `scripts/md2word.cmd`: Windows 命令行包装器
- `scripts/md2word.ps1`: Windows 主入口，统一调度模板分流和 Word 流水线
- `scripts/render_mht.py`: 通用模板渲染器，负责占位符、表格标签填充和正文推断替换
- `scripts/generate_template_mht.py`: 从模板规格或文字描述生成可复用 MHT 模板
- `scripts/word_template_to_mht.ps1`: 用 Windows Word COM 把 DOCX 模板归一化为 MHT
- `scripts/word_mht_pipeline.ps1`: 用 Windows Word COM 生成 `.docx` 并把 LaTeX 转成 Office Math
- `assets/专利交底书模板.mht`: 默认示例模板（已清理可识别元数据）
- `references/template-spec.md`: 模板规格参考

## 失败处理

- 如果 PowerShell、Python 或 Word COM 不可用，不要假装能在别的链路里补全最终 Word 效果，直接报环境问题。
- 如果模板效果不对，先看 `-TemplateReport`，确认正文被替换到了哪一段。
- 如果公式在最终 `.wordmath.docx` 中异常，先检查原始 `.docx`，区分问题是在 MHT 阶段还是 Word 公式阶段。
- 如果 Markdown 里有 Mermaid 代码块但当前环境既没有 `mmdc`，也没有可用的 `Node.js + npx`，直接提示补依赖，不要静默降级。
- 如果走 `npx` 兜底失败，先检查是否允许联网拉取 `@mermaid-js/mermaid-cli`。
