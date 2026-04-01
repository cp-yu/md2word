#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html
import re
from email.message import EmailMessage
from email.generator import BytesGenerator
from io import BytesIO
from pathlib import Path


SPEC_LINE_RE = re.compile(r"^(?:[-*]\s*)?([A-Za-z0-9_\-\u4e00-\u9fff ]{1,40})[:：]\s*(.+?)\s*$")
DEFAULT_STYLE_PRESET = "default"
ACADEMIC_PAPER_STYLE_PRESET = "academic-paper"
STYLE_PRESETS = {
    DEFAULT_STYLE_PRESET: {
        "title_font": "SimSun",
        "title_size_pt": 22.0,
        "title_margin_top_pt": 48.0,
        "title_line_height_pt": 30.0,
        "subtitle_font": "SimSun",
        "subtitle_size_pt": 12.0,
        "subtitle_line_height_pt": 22.0,
        "metadata_heading_font": "SimSun",
        "metadata_heading_size_pt": 12.0,
        "metadata_heading_line_height_pt": 22.0,
        "body_heading_font": "SimSun",
        "body_heading_size_pt": 14.0,
        "body_heading_line_height_pt": 22.0,
        "note_font": "SimSun",
        "note_size_pt": 11.0,
        "note_line_height_pt": 22.0,
    },
    ACADEMIC_PAPER_STYLE_PRESET: {
        "title_font": "SimHei",
        "title_size_pt": 18.0,
        "title_margin_top_pt": 36.0,
        "title_line_height_pt": 27.0,
        "subtitle_font": "SimSun",
        "subtitle_size_pt": 12.0,
        "subtitle_line_height_pt": 18.0,
        "metadata_heading_font": "SimHei",
        "metadata_heading_size_pt": 15.0,
        "metadata_heading_line_height_pt": 22.5,
        "body_heading_font": "SimHei",
        "body_heading_size_pt": 15.0,
        "body_heading_line_height_pt": 22.5,
        "note_font": "SimSun",
        "note_size_pt": 12.0,
        "note_line_height_pt": 18.0,
    },
}


def ascii_html(text: str) -> str:
    escaped = html.escape(text, quote=False)
    out: list[str] = []
    for ch in escaped:
        if ord(ch) < 128:
            out.append(ch)
        else:
            out.append(f"&#{ord(ch)};")
    return "".join(out)


def normalize_key(text: str) -> str:
    key = text.strip().lower()
    key = key.replace(" ", "")
    key = key.replace("_", "-")
    return key


def parse_spec_text(text: str) -> dict[str, str]:
    config = {
        "title": "{{TITLE}}",
        "subtitle": "",
        "metadata-heading": "文档信息",
        "body-heading": "正文",
        "cover-note": "",
        "description": "",
        "style-preset": DEFAULT_STYLE_PRESET,
    }
    aliases = {
        "title": {"title", "标题", "文档标题", "document-title", "documenttitle"},
        "subtitle": {"subtitle", "副标题", "子标题"},
        "metadata-heading": {"metadata-heading", "metadata", "元数据标题", "信息标题"},
        "body-heading": {"body-heading", "body", "正文标题", "内容标题"},
        "cover-note": {"cover-note", "cover", "封面说明", "模板说明", "note", "备注"},
        "style-preset": {"style-preset", "style", "preset", "样式预设", "样式", "格式预设"},
    }

    description_lines: list[str] = []
    for raw_line in text.splitlines():
        stripped = raw_line.strip()
        if not stripped:
            description_lines.append("")
            continue
        match = SPEC_LINE_RE.match(stripped)
        if not match:
            description_lines.append(stripped)
            continue

        key = normalize_key(match.group(1))
        value = match.group(2).strip()
        mapped_key = None
        for canonical, keys in aliases.items():
            if key in keys:
                mapped_key = canonical
                break
        if mapped_key is None:
            description_lines.append(stripped)
            continue
        config[mapped_key] = value

    description = "\n".join(description_lines).strip()
    if description:
        config["description"] = description
        if not config["cover-note"]:
            config["cover-note"] = description
    return config


def resolve_style_preset(name: str) -> dict[str, float | str]:
    return STYLE_PRESETS.get(name, STYLE_PRESETS[DEFAULT_STYLE_PRESET])


def render_note_paragraphs(text: str, style: dict[str, float | str]) -> str:
    if not text.strip():
        return ""

    parts: list[str] = []
    for raw_paragraph in re.split(r"\n\s*\n", text.strip()):
        lines = [line.strip() for line in raw_paragraph.splitlines() if line.strip()]
        if not lines:
            continue
        html_text = "<br>".join(ascii_html(line) for line in lines)
        parts.append(
            "<p class=MsoNormal style='text-align:left;line-height:"
            f"{float(style['note_line_height_pt']):.1f}pt;mso-line-height-rule:exactly'>"
            "<span style='font-size:"
            f"{float(style['note_size_pt']):.1f}pt;font-family:{style['note_font']}'>"
            f"{html_text}"
            "</span></p>"
        )
    return "\n".join(parts)


def build_html(config: dict[str, str]) -> str:
    style = resolve_style_preset(config.get("style-preset", DEFAULT_STYLE_PRESET))
    title = config["title"] or "{{TITLE}}"
    subtitle = config["subtitle"].strip()
    metadata_heading = config["metadata-heading"].strip()
    body_heading = config["body-heading"].strip()
    cover_note = config["cover-note"].strip()

    title_block = (
        "<p class=MsoTitle align=center style='text-align:center;margin-top:"
        f"{float(style['title_margin_top_pt']):.1f}pt;line-height:{float(style['title_line_height_pt']):.1f}pt;"
        "mso-line-height-rule:exactly'>"
        "<b><span style='font-size:"
        f"{float(style['title_size_pt']):.1f}pt;font-family:{style['title_font']}'>"
        f"{ascii_html(title)}"
        "</span></b></p>"
    )
    subtitle_block = ""
    if subtitle:
        subtitle_block = (
            "<p class=MsoNormal align=center style='text-align:center;line-height:"
            f"{float(style['subtitle_line_height_pt']):.1f}pt;mso-line-height-rule:exactly'>"
            "<span style='font-size:"
            f"{float(style['subtitle_size_pt']):.1f}pt;font-family:{style['subtitle_font']}'>"
            f"{ascii_html(subtitle)}"
            "</span></p>"
        )

    metadata_block = ""
    if metadata_heading:
        metadata_block = (
            "<p class=MsoNormal style='margin-top:24.0pt'>"
            "<b><span style='font-size:"
            f"{float(style['metadata_heading_size_pt']):.1f}pt;font-family:{style['metadata_heading_font']};"
            f"line-height:{float(style['metadata_heading_line_height_pt']):.1f}pt'>"
            f"{ascii_html(metadata_heading)}"
            "</span></b></p>\n"
            f"{'{{METADATA_TABLE}}'}"
        )

    body_heading_block = ""
    if body_heading:
        body_heading_block = (
            "<p class=MsoNormal style='margin-bottom:12.0pt'>"
            "<b><span style='font-size:"
            f"{float(style['body_heading_size_pt']):.1f}pt;font-family:{style['body_heading_font']};"
            f"line-height:{float(style['body_heading_line_height_pt']):.1f}pt'>"
            f"{ascii_html(body_heading)}"
            "</span></b></p>"
        )

    return f"""<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<title>{ascii_html(title)}</title>
<style>
<!--
@page WordSection1 {{size:595.3pt 841.9pt;margin:72.0pt 72.0pt 72.0pt 72.0pt;}}
@page WordSection2 {{size:595.3pt 841.9pt;margin:72.0pt 72.0pt 72.0pt 72.0pt;}}
div.WordSection1 {{page:WordSection1;}}
div.WordSection2 {{page:WordSection2;}}
p.MsoNormal, li.MsoNormal, div.MsoNormal {{margin:0cm;font-size:12.0pt;font-family:SimSun;}}
p.MsoTitle {{margin:0cm;font-size:22.0pt;font-family:SimSun;}}
table.MsoTableGrid {{border-collapse:collapse;mso-table-layout-alt:fixed;}}
-->
</style>
</head>
<body lang=ZH-CN style='word-wrap:break-word'>
<div class=WordSection1>
{title_block}
{subtitle_block}
{render_note_paragraphs(cover_note, style)}
{metadata_block}
<p class=MsoNormal style='page-break-after:always'><span style='font-size:1.0pt'>&nbsp;</span></p>
</div>
<div class=WordSection2 style='layout-grid:15.6pt'>
{body_heading_block}
{{{{CONTENT}}}}
</div>
</body>
</html>
"""


def build_mht(html_text: str) -> bytes:
    root = EmailMessage()
    root.set_type("multipart/related")
    root.set_param("type", "text/html")

    html_part = EmailMessage()
    html_part.set_content(
        html_text,
        subtype="html",
        charset="us-ascii",
        cte="quoted-printable",
    )
    html_part["Content-Location"] = "file:///C:/generated-template.htm"
    root.attach(html_part)

    buffer = BytesIO()
    BytesGenerator(buffer).flatten(root)
    return buffer.getvalue()


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate a reusable MHT template from a simple spec or plain text description")
    parser.add_argument("--spec", "-s", required=True, help="Template spec or description markdown path")
    parser.add_argument("--output", "-o", required=True, help="Output MHT path")
    parser.add_argument(
        "--style-preset",
        default="",
        choices=["", *sorted(STYLE_PRESETS)],
        help="Optional style preset override for the generated template",
    )
    args = parser.parse_args()

    spec_text = Path(args.spec).read_text(encoding="utf-8")
    config = parse_spec_text(spec_text)
    if args.style_preset:
        config["style-preset"] = args.style_preset
    html_text = build_html(config)

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(build_mht(html_text))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
