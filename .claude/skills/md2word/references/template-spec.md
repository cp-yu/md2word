# Template Spec 参考

`-TemplateSpec` 接收一个本地文本文件。这个文件既可以写成简单规格，也可以只是一段文字描述。

## 方式一：键值规格

```md
title: 技术评审报告
subtitle: 结构优先模板
metadata-heading: 文档信息
body-heading: 详细内容
cover-note: 首页保留标题、说明和元数据表，第二页开始放正文。
style-preset: default
```

支持字段：

- `title`: 封面标题；默认 `{{TITLE}}`
- `subtitle`: 封面副标题
- `metadata-heading`: 元数据区标题；留空可不显示
- `body-heading`: 正文区标题；留空可不显示
- `cover-note`: 封面说明文字
- `style-preset`: 模板样式预设；当前支持 `default`、`academic-paper`

## 学术论文示例

当用户给出“标题黑体小二、一级标题黑体小三、正文宋体小四、1.5 倍行距”这类要求时，直接写：

```md
title: {{TITLE}}
subtitle:
metadata-heading: 文档信息
body-heading: 正文
style-preset: academic-paper
cover-note: 学术论文导出版式。标题使用黑体小二，一级标题使用黑体小三，正文使用宋体小四，统一按 1.5 倍行距处理。
```

如果要确保渲染和模板同时走论文样式，调用脚本时再额外传：

```powershell
-StylePreset academic-paper
```

## 方式二：纯文字描述

```md
做一个简洁的技术方案模板。
第一页放标题、摘要说明和文档信息表。
第二页开始放正文。
整体风格正式、适合导出到 Word。
```

这种写法不会做复杂语义设计，它只会生成一个“结构优先”的可复用模板：

- 封面
- 可自动填充的元数据表
- 正文区

## 可用占位符

生成出的模板和通用 MHT 模板渲染器支持这些占位符：

- `{{TITLE}}`: 从 Markdown 头字段自动取标题
- `{{METADATA_TABLE}}`: 把 Markdown 头字段渲染为表格
- `{{CONTENT}}`: 正文插入点
- `{{字段名}}`: 直接按字段名替换，例如 `{{项目名称}}`

## Markdown 头字段格式

模板填充值来自 Markdown 开头的字段区，当前兼容这两种写法：

```md
**项目名称：** 示例项目
**申请人名称：** 示例公司
**版本：** V1.0

---
```

```md
标题: 示例文档
作者: 张三
版本: v1.0

# 示例文档
```

字段区之后的正文会进入 `{{CONTENT}}` 或自动推断出的正文 section。
