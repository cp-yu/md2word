#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html
import json
import mimetypes
import os
import re
import shutil
import subprocess
import tempfile
from email import policy
from email.generator import BytesGenerator
from email.message import EmailMessage
from email.parser import BytesParser
from io import BytesIO
from pathlib import Path

from style_presets import DEFAULT_STYLE_PRESET, get_render_style, list_preset_names


BOLD_FIELD_LINE_RE = re.compile(r"^\*\*(.+?)：\*\*\s*(.*)$")
PLAIN_FIELD_LINE_RE = re.compile(r"^([A-Za-z0-9_\-\u4e00-\u9fff ]{1,40})[:：]\s*(.*)$")
ORDERED_ITEM_RE = re.compile(r"^(\d+)\.\s+(.*)$")
UNORDERED_ITEM_RE = re.compile(r"^([-*+])\s+(.*)$")
TABLE_RE = re.compile(r"(?is)<table\b.*?</table>")
TR_RE = re.compile(r"(?is)<tr\b.*?</tr>")
TD_RE = re.compile(r"(?is)(<td\b[^>]*>)(.*?)(</td>)")
TAG_RE = re.compile(r"(?is)<[^>]+>")
SECTION_RE = re.compile(
    r"(?is)(<div\b[^>]*class=(?P<quote>['\"]?)(?P<class>[^>]*\bWordSection\d+\b[^>]*)"
    r"(?P=quote)[^>]*>)(?P<body>.*?)(</div>)"
)
TITLE_RE = re.compile(r"(?is)<title>.*?</title>")
CONTENT_MARKERS = ("{{CONTENT}}", "<!--MD_CONTENT-->")
METADATA_TABLE_MARKER = "{{METADATA_TABLE}}"
TITLE_MARKERS = ("{{TITLE}}", "{{DOCUMENT_TITLE}}")
BODY_HINTS = ("正文", "main content", "document body", "report body", "content section")
IMAGE_LINE_RE = re.compile(r'^!\[(?P<alt>[^\]]*)\]\((?P<src>[^)]+)\)\s*$')
TABLE_SEPARATOR_CELL_RE = re.compile(r"^:?-{3,}:?$")
EXPLICIT_HEADING_PREFIX_RE = re.compile(
    r"^(?:\d+(?:\.\d+)*|[一二三四五六七八九十百千零〇两]+|[Ss]\d+)[、,.，．]?\s*.+$"
)
HEADING_MARKER_PREFIX = "[[MD2WORD_HEADING:"
HEADING_MARKER_SUFFIX = "]]"
MAX_WORD_HEADING_LEVEL = 6


def normalize_label(label: str) -> str:
    normalized = html.unescape(label.strip())
    normalized = normalized.replace(" ", "")
    normalized = normalized.replace("－", "-")
    normalized = normalized.replace("／", "/")
    normalized = normalized.replace("：", "")
    normalized = normalized.replace(":", "")
    return normalized


def strip_inline_markdown(text: str) -> str:
    text = text.strip()
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    text = re.sub(r"\*([^*]+)\*", r"\1", text)
    return re.sub(r"\s+", " ", text).strip()


def parse_image_line(text: str) -> tuple[str, str] | None:
    match = IMAGE_LINE_RE.match(text.strip())
    if not match:
        return None
    alt = strip_inline_markdown(match.group("alt")) or "插图"
    src = match.group("src").strip()
    if not src:
        return None
    if src.startswith("<") and src.endswith(">"):
        src = src[1:-1].strip()
    return alt, src


def split_table_row(line: str) -> list[str] | None:
    stripped = line.strip()
    if "|" not in stripped:
        return None
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        stripped = stripped[:-1]
    cells = [strip_inline_markdown(cell.strip()) for cell in stripped.split("|")]
    if len(cells) < 2:
        return None
    return cells


def parse_table_alignment(line: str) -> list[str] | None:
    cells = split_table_row(line)
    if cells is None:
        return None

    alignments: list[str] = []
    for cell in cells:
        compact = cell.replace(" ", "")
        if not TABLE_SEPARATOR_CELL_RE.fullmatch(compact):
            return None
        if compact.startswith(":") and compact.endswith(":"):
            alignments.append("center")
        elif compact.endswith(":"):
            alignments.append("right")
        else:
            alignments.append("left")
    return alignments


def parse_table_block(lines: list[str], start_index: int) -> tuple[dict[str, object], int] | None:
    if start_index + 1 >= len(lines):
        return None

    header = split_table_row(lines[start_index])
    alignments = parse_table_alignment(lines[start_index + 1])
    if header is None or alignments is None or len(header) != len(alignments):
        return None

    rows: list[list[str]] = []
    i = start_index + 2
    while i < len(lines):
        stripped = lines[i].strip()
        if not stripped:
            break
        row = split_table_row(lines[i])
        if row is None or len(row) != len(header):
            break
        rows.append(row)
        i += 1

    return (
        {
            "type": "table",
            "headers": header,
            "alignments": alignments,
            "rows": rows,
        },
        i,
    )


def ascii_html(text: str, *, quote: bool = False) -> str:
    escaped = html.escape(text, quote=quote)
    out: list[str] = []
    for ch in escaped:
        if ord(ch) < 128:
            out.append(ch)
        else:
            out.append(f"&#{ord(ch)};")
    return "".join(out)


def leading_indent(line: str) -> int:
    expanded = line.expandtabs(4)
    return len(expanded) - len(expanded.lstrip(" "))


def extract_text(fragment: str) -> str:
    text = TAG_RE.sub(" ", fragment)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def parse_list_item(line: str) -> dict[str, object] | None:
    expanded = line.expandtabs(4).rstrip("\n")
    indent = leading_indent(expanded)
    stripped = expanded[indent:]

    ordered = ORDERED_ITEM_RE.match(stripped)
    if ordered:
        return {
            "indent": indent,
            "ordered": True,
            "start": int(ordered.group(1)),
            "text": ordered.group(2),
        }

    unordered = UNORDERED_ITEM_RE.match(stripped)
    if unordered:
        return {
            "indent": indent,
            "ordered": False,
            "text": unordered.group(2),
        }

    return None


def parse_list_block(lines: list[str], start_index: int) -> tuple[dict[str, object], int]:
    first = parse_list_item(lines[start_index])
    if first is None:
        raise ValueError("parse_list_block called on non-list line")

    base_indent = int(first["indent"])
    ordered = bool(first["ordered"])
    items: list[dict[str, object]] = []
    i = start_index

    while i < len(lines):
        current = parse_list_item(lines[i])
        if current is None:
            break
        if int(current["indent"]) != base_indent or bool(current["ordered"]) != ordered:
            break

        item_text = strip_inline_markdown(str(current["text"]))
        i += 1
        children: list[dict[str, object]] = []
        continuation: list[str] = []

        while i < len(lines):
            raw = lines[i]
            if not raw.strip():
                break

            nested = parse_list_item(raw)
            if nested is not None:
                nested_indent = int(nested["indent"])
                if nested_indent > base_indent:
                    child_block, i = parse_list_block(lines, i)
                    children.append(child_block)
                    continue
                if nested_indent <= base_indent:
                    break

            if leading_indent(raw) > base_indent:
                continuation.append(strip_inline_markdown(raw.strip()))
                i += 1
                continue

            break

        full_text = " ".join(x for x in [item_text, *continuation] if x).strip()
        items.append({"text": full_text, "children": children})

    block: dict[str, object] = {
        "type": "list",
        "ordered": ordered,
        "items": items,
    }
    if ordered:
        block["start"] = int(first["start"])
    return block, i


def parse_header_field_line(line: str) -> tuple[str, str] | None:
    match = BOLD_FIELD_LINE_RE.match(line)
    if match:
        return normalize_label(match.group(1)), match.group(2).strip()

    match = PLAIN_FIELD_LINE_RE.match(line)
    if match:
        return normalize_label(match.group(1)), match.group(2).strip()

    return None


def has_header_fields_ahead(lines: list[str], start_index: int) -> bool:
    i = start_index
    while i < len(lines):
        stripped = lines[i].strip()
        if not stripped:
            i += 1
            continue
        return parse_header_field_line(stripped) is not None
    return False


def parse_markdown(md_text: str) -> tuple[list[tuple[str, str]], list[dict[str, object]]]:
    lines = md_text.splitlines()
    header_items: list[tuple[str, str]] = []
    blocks: list[dict[str, object]] = []
    in_header = True
    saw_header_items = False
    paragraph: list[str] = []
    i = 0

    def flush_paragraph() -> None:
        nonlocal paragraph
        if not paragraph:
            return
        text = strip_inline_markdown(" ".join(x.strip() for x in paragraph if x.strip()))
        paragraph = []
        if text:
            blocks.append({"type": "paragraph", "text": text})

    while i < len(lines):
        line = lines[i].rstrip("\n")
        stripped = line.strip()

        if in_header:
            header_field = parse_header_field_line(stripped)
            if header_field is not None:
                header_items.append(header_field)
                saw_header_items = True
                i += 1
                continue
            if stripped == "" and not saw_header_items and has_header_fields_ahead(lines, i + 1):
                i += 1
                continue
            if stripped.startswith("# ") and (saw_header_items or has_header_fields_ahead(lines, i + 1)):
                # Metadata may be preceded or followed by a duplicated H1 title.
                # Keep it as document title instead of repeating it in body.
                if saw_header_items:
                    in_header = False
                i += 1
                continue
            if saw_header_items and stripped == "---":
                in_header = False
                i += 1
                continue
            if saw_header_items and stripped == "":
                i += 1
                continue
            if saw_header_items:
                in_header = False
                continue
            in_header = False

        if not stripped:
            flush_paragraph()
            i += 1
            continue

        if stripped == "---":
            flush_paragraph()
            i += 1
            continue

        if stripped.startswith("```"):
            flush_paragraph()
            lang = stripped[3:].strip()
            i += 1
            code_lines: list[str] = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i].rstrip("\n"))
                i += 1
            if lang.lower() == "mermaid":
                blocks.append({"type": "mermaid", "lines": code_lines})
            else:
                blocks.append({"type": "code", "lang": lang, "lines": code_lines})
            if i < len(lines):
                i += 1
            continue

        if stripped == "$$":
            flush_paragraph()
            i += 1
            formula_lines: list[str] = []
            while i < len(lines) and lines[i].strip() != "$$":
                formula_lines.append(lines[i].rstrip("\n"))
                i += 1
            blocks.append({"type": "formula", "lines": formula_lines})
            if i < len(lines):
                i += 1
            continue

        image_line = parse_image_line(stripped)
        if image_line is not None:
            flush_paragraph()
            alt, src = image_line
            blocks.append({"type": "image", "src": src, "alt": alt})
            i += 1
            continue

        table_block = parse_table_block(lines, i)
        if table_block is not None:
            flush_paragraph()
            block, i = table_block
            blocks.append(block)
            continue

        if stripped.startswith("#"):
            flush_paragraph()
            level = len(stripped) - len(stripped.lstrip("#"))
            text = strip_inline_markdown(stripped[level:].strip())
            if text:
                blocks.append({"type": "heading", "level": level, "text": text})
            i += 1
            continue

        list_item = parse_list_item(line)
        if list_item is not None:
            flush_paragraph()
            block, i = parse_list_block(lines, i)
            blocks.append(block)
            continue

        paragraph.append(stripped)
        i += 1

    flush_paragraph()
    renumber_heading_blocks(blocks)
    return header_items, blocks


def renumber_heading_blocks(blocks: list[dict[str, object]]) -> None:
    heading_blocks = [block for block in blocks if block.get("type") == "heading"]
    if not heading_blocks:
        return

    base_level = min(max(int(block.get("level", 1)), 1) for block in heading_blocks)
    counters = [0] * MAX_WORD_HEADING_LEVEL

    for block in heading_blocks:
        text = str(block.get("text", "")).strip()
        if not text:
            continue

        normalized_level = min(max(int(block.get("level", 1)) - base_level + 1, 1), MAX_WORD_HEADING_LEVEL)
        for index in range(normalized_level, MAX_WORD_HEADING_LEVEL):
            counters[index] = 0

        if EXPLICIT_HEADING_PREFIX_RE.match(text):
            continue

        counters[normalized_level - 1] += 1
        if normalized_level == 1:
            block["text"] = f"{counters[0]}、{text}"
            continue

        prefix = ".".join(str(counters[index]) for index in range(normalized_level))
        block["text"] = f"{prefix} {text}"


def collect_values_by_label(header_items: list[tuple[str, str]]) -> dict[str, list[str]]:
    values_by_label: dict[str, list[str]] = {}
    for label, value in header_items:
        values_by_label.setdefault(label, []).append(value)
    return values_by_label


def pick_title(header_items: list[tuple[str, str]]) -> str:
    values_by_label = collect_values_by_label(header_items)
    for key in ("标题", "题目", "主题", "发明名称", "项目名称"):
        values = values_by_label.get(key, [])
        if values and values[0]:
            return values[0]
    return "Markdown Document"


def render_value_cell(text: str) -> str:
    content = ascii_html(text) if text else "&nbsp;"
    return (
        "<p class=MsoNormal style='margin:0;line-height:normal'>"
        "<span style='font-size:10.5pt;font-family:SimSun'>"
        f"{content}"
        "</span></p>"
    )


def render_metadata_table(header_items: list[tuple[str, str]]) -> str:
    if not header_items:
        return ""

    rows: list[str] = []
    for label, value in header_items:
        rows.append(
            "<tr>"
            "<td width=165 valign=top style='border:solid windowtext 1.0pt;"
            "padding:4.0pt 6.0pt;background:#F2F2F2'>"
            f"{render_value_cell(label)}"
            "</td>"
            "<td width=360 valign=top style='border:solid windowtext 1.0pt;"
            "border-left:none;padding:4.0pt 6.0pt'>"
            f"{render_value_cell(value)}"
            "</td>"
            "</tr>"
        )

    return (
        "<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 "
        "style='border-collapse:collapse;border:none;mso-border-alt:solid windowtext .5pt'>"
        + "".join(rows)
        + "</table>"
    )


def render_body_table_cell(
    text: str,
    *,
    style: dict[str, object],
    header: bool = False,
    alignment: str = "left",
) -> str:
    weight_prefix = "<b>" if header else ""
    weight_suffix = "</b>" if header else ""
    text_align = alignment if alignment in {"left", "center", "right"} else "left"
    return (
        "<p class=MsoNormal style='margin:0;"
        f"text-align:{text_align};line-height:{float(style['body_line_height_pt']):.1f}pt;"
        "mso-line-height-rule:exactly'>"
        f"{weight_prefix}<span style='font-size:{float(style['body_size_pt']):.1f}pt;"
        f"font-family:{style['body_font']}'>{ascii_html(text) if text else '&nbsp;'}</span>{weight_suffix}"
        "</p>"
    )


def render_body_table(block: dict[str, object], style: dict[str, object]) -> str:
    headers = [str(x) for x in block.get("headers", [])]
    rows = [[str(cell) for cell in row] for row in block.get("rows", [])]
    alignments = [str(x) for x in block.get("alignments", [])]
    if not headers:
        return ""

    column_count = len(headers)
    width_pct = 100.0 / column_count

    rendered_rows: list[str] = []
    header_cells = "".join(
        "<td valign=top width={width}% style='width:{width}%;border:solid windowtext 1.0pt;"
        "padding:4.0pt 6.0pt;background:#F2F2F2'>{content}</td>".format(
            width=int(round(width_pct)),
            content=render_body_table_cell(
                headers[index],
                style=style,
                header=True,
                alignment=alignments[index] if index < len(alignments) else "left",
            ),
        )
        for index in range(column_count)
    )
    rendered_rows.append(f"<tr>{header_cells}</tr>")

    for row in rows:
        body_cells = "".join(
            "<td valign=top width={width}% style='width:{width}%;border:solid windowtext 1.0pt;"
            "border-top:none;padding:4.0pt 6.0pt'>{content}</td>".format(
                width=int(round(width_pct)),
                content=render_body_table_cell(
                    row[index] if index < len(row) else "",
                    style=style,
                    alignment=alignments[index] if index < len(alignments) else "left",
                ),
            )
            for index in range(column_count)
        )
        rendered_rows.append(f"<tr>{body_cells}</tr>")

    return (
        "<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 "
        "style='margin-top:6.0pt;margin-bottom:12.0pt;border-collapse:collapse;"
        "border:none;mso-border-alt:solid windowtext .5pt;width:100%'>"
        + "".join(rendered_rows)
        + "</table>"
    )


def replace_nth_td(row_html: str, index: int, inner_html: str) -> str:
    out: list[str] = []
    last = 0
    for current, match in enumerate(TD_RE.finditer(row_html)):
        out.append(row_html[last:match.start()])
        if current == index:
            out.append(match.group(1))
            out.append(inner_html)
            out.append(match.group(3))
        else:
            out.append(match.group(0))
        last = match.end()
    out.append(row_html[last:])
    return "".join(out)


def match_header_label(cell_text: str, known_labels: list[str]) -> str | None:
    normalized = normalize_label(cell_text)
    if not normalized:
        return None

    for label in sorted(known_labels, key=len, reverse=True):
        if normalized == label:
            return label
        if normalized.startswith(label) and len(normalized) <= len(label) + 4:
            return label
        if label in normalized and len(normalized) <= len(label) + 6:
            return label
    return None


def fill_cover_tables(
    source_html: str,
    header_items: list[tuple[str, str]],
    report: dict[str, object],
) -> str:
    values_by_label = collect_values_by_label(header_items)
    if not values_by_label:
        report["table_replacements"] = []
        return source_html

    if values_by_label.get("专利类型") and values_by_label.get("PCT国际申请"):
        patent_type = values_by_label["专利类型"][0]
        pct_type = values_by_label["PCT国际申请"][0]
        if patent_type and pct_type:
            values_by_label["专利类型"][0] = f"{patent_type}；{pct_type}"

    counters = {label: 0 for label in values_by_label}
    replacements: list[dict[str, object]] = []
    new_html = source_html
    table_offset = 0

    for table_index, table_match in enumerate(TABLE_RE.finditer(source_html), start=1):
        start = table_match.start() + table_offset
        end = table_match.end() + table_offset
        table_html = new_html[start:end]
        rows = TR_RE.findall(table_html)
        updated_rows: list[str] = []

        for row_index, row_html in enumerate(rows, start=1):
            cells = list(TD_RE.finditer(row_html))
            updated_row = row_html
            for cell_index, cell in enumerate(cells):
                label = match_header_label(extract_text(cell.group(2)), list(values_by_label))
                if not label:
                    continue
                value_list = values_by_label[label]
                value_index = counters[label]
                if value_index >= len(value_list):
                    continue
                target_index = cell_index + 1 if cell_index + 1 < len(cells) else None
                if target_index is None:
                    continue
                value = value_list[value_index]
                updated_row = replace_nth_td(updated_row, target_index, render_value_cell(value))
                counters[label] += 1
                replacements.append(
                    {
                        "label": label,
                        "value": value,
                        "table": table_index,
                        "row": row_index,
                        "column": target_index + 1,
                    }
                )
                break
            updated_rows.append(updated_row)

        rebuilt = table_html
        for original_row, updated_row in zip(rows, updated_rows):
            rebuilt = rebuilt.replace(original_row, updated_row, 1)

        new_html = new_html[:start] + rebuilt + new_html[end:]
        table_offset += len(rebuilt) - (end - start)

    report["table_replacements"] = replacements
    return new_html


def paragraph_style_for(
    indent_pt: float,
    line_height_pt: float,
    *,
    single_spacing: bool = False,
) -> str:
    if single_spacing:
        return (
            f"text-indent:{indent_pt:.1f}pt;mso-char-indent-count:2.0;"
            "line-height:normal;mso-line-height-rule:auto"
        )
    return (
        f"text-indent:{indent_pt:.1f}pt;mso-char-indent-count:2.0;"
        f"line-height:{line_height_pt:.1f}pt;"
        "mso-line-height-rule:exactly"
    )


def heading_size_for(level: int, style: dict[str, object]) -> float:
    heading_sizes = dict(style["heading_sizes"])
    if level <= 3:
        return float(heading_sizes["h1_h3"])
    if level == 4:
        return float(heading_sizes["h4"])
    return float(heading_sizes["h5_plus"])


def heading_marker(level: int) -> str:
    normalized_level = max(1, min(level, MAX_WORD_HEADING_LEVEL))
    return f"{HEADING_MARKER_PREFIX}{normalized_level}{HEADING_MARKER_SUFFIX}"


def body_heading(text: str, level: int, size_pt: float, style: dict[str, object]) -> str:
    marker_html = (
        "<span style='font-size:1.0pt;color:white'>"
        f"{ascii_html(heading_marker(level))}"
        "</span>"
    )
    return (
        "<p class=MsoNormal style='"
        f"line-height:{float(style['heading_line_height_pt']):.1f}pt;"
        "mso-line-height-rule:exactly'>"
        f"{marker_html}"
        "<b><span style='font-size:"
        f"{size_pt:.1f}pt;mso-bidi-font-size:10.0pt;font-family:{style['heading_font']}'>"
        f"{ascii_html(text)}"
        "</span></b></p>"
    )


def body_paragraph(
    text: str,
    indent_pt: float = 24.0,
    monospace: bool = False,
    *,
    single_spacing: bool = False,
    style: dict[str, object],
) -> str:
    font = str(style["code_font"] if monospace else style["body_font"])
    size = float(style["code_size_pt"] if monospace else style["body_size_pt"])
    line_height_pt = float(
        style["code_line_height_pt"] if monospace else style["body_line_height_pt"]
    )
    return (
        "<p class=MsoNormal style='"
        f"{paragraph_style_for(indent_pt, line_height_pt, single_spacing=single_spacing)}'>"
        "<span"
        f" style='font-size:{size:.1f}pt;font-family:{font}'>"
        f"{ascii_html(text) if text else '&nbsp;'}"
        "</span></p>"
    )


def body_multiline_paragraph(
    lines: list[str],
    indent_pt: float = 24.0,
    monospace: bool = False,
    *,
    single_spacing: bool = False,
    style: dict[str, object],
) -> str:
    font = str(style["code_font"] if monospace else style["body_font"])
    size = float(style["code_size_pt"] if monospace else style["body_size_pt"])
    line_height_pt = float(
        style["code_line_height_pt"] if monospace else style["body_line_height_pt"]
    )
    content = "<br>".join(ascii_html(line) if line else "&nbsp;" for line in lines)
    return (
        "<p class=MsoNormal style='"
        f"{paragraph_style_for(indent_pt, line_height_pt, single_spacing=single_spacing)}'>"
        "<span"
        f" style='font-size:{size:.1f}pt;font-family:{font}'>"
        f"{content}"
        "</span></p>"
    )


def body_image_paragraph(src: str, alt: str) -> str:
    return (
        "<p class=MsoNormal align=center style='text-align:center;line-height:normal;"
        "margin-top:6.0pt;margin-bottom:6.0pt'>"
        f"<img src=\"{ascii_html(src, quote=True)}\" alt=\"{ascii_html(alt, quote=True)}\">"
        "</p>"
    )


def render_list_html(
    block: dict[str, object],
    style: dict[str, object],
    level: int = 0,
    lfo: int = 1,
) -> str:
    ordered = bool(block["ordered"])
    tag = "ol" if ordered else "ul"
    list_style = "decimal" if ordered else ("disc" if level == 0 else "circle" if level == 1 else "square")
    start_attr = ""
    if ordered:
        start = int(block.get("start", 1))
        if start > 1:
            start_attr = f" start={start}"

    rendered_items: list[str] = []
    for item in list(block["items"]):
        text = ascii_html(str(item.get("text", "")))
        child_html = "".join(
            render_list_html(child, style, level + 1, lfo=lfo)
            for child in list(item.get("children", []))
        )
        rendered_items.append(
            "<li style='margin-bottom:6.0pt;"
            f"line-height:{float(style['body_line_height_pt']):.1f}pt;"
            "mso-line-height-rule:exactly;"
            f"mso-list:l0 level{min(level + 1, 9)} lfo{lfo}'>"
            "<span style='font-size:"
            f"{float(style['body_size_pt']):.1f}pt;font-family:{style['body_font']}'>"
            f"{text or '&nbsp;'}"
            "</span>"
            f"{child_html}"
            "</li>"
        )

    margin_left = 36.0 + level * 18.0
    return (
        f"<{tag}{start_attr} style='margin-top:0;margin-bottom:0;"
        f"margin-left:{margin_left:.1f}pt;padding-left:18.0pt;"
        f"font-family:{style['body_font']};font-size:{float(style['body_size_pt']):.1f}pt;"
        f"line-height:{float(style['body_line_height_pt']):.1f}pt;"
        f"mso-line-height-rule:exactly;list-style-type:{list_style}'>"
        + "".join(rendered_items)
        + f"</{tag}>"
    )


def render_body_inner(blocks: list[dict[str, object]], style: dict[str, object]) -> str:
    parts: list[str] = []
    list_serial = 0
    for block in blocks:
        kind = block["type"]
        if kind == "heading":
            level = int(block["level"])
            text = str(block["text"])
            parts.append(body_heading(text, level, heading_size_for(level, style), style))
            continue

        if kind == "paragraph":
            parts.append(body_paragraph(str(block["text"]), style=style))
            continue

        if kind == "table":
            rendered = render_body_table(block, style)
            if rendered:
                parts.append(rendered)
            continue

        if kind == "list":
            list_serial += 1
            parts.append(render_list_html(block, style, lfo=list_serial))
            continue

        if kind == "formula":
            formula_lines = [str(x).rstrip() for x in block["lines"]]
            parts.append(
                body_multiline_paragraph(
                    ["$$", *(formula_lines or [""]), "$$"],
                    indent_pt=36.0,
                    monospace=True,
                    single_spacing=True,
                    style=style,
                )
            )
            continue

        if kind == "mermaid":
            src = str(block.get("content_location", "")).strip()
            alt = str(block.get("alt", "Mermaid 图")).strip() or "Mermaid 图"
            if src:
                parts.append(body_image_paragraph(src, alt))
            continue

        if kind == "image":
            src = str(block.get("content_location", "")).strip()
            alt = str(block.get("alt", "插图")).strip() or "插图"
            if src:
                parts.append(body_image_paragraph(src, alt))
            continue

        if kind == "code":
            lang = str(block["lang"]).strip()
            if lang:
                parts.append(
                    body_paragraph(f"[{lang}]", indent_pt=36.0, monospace=True, style=style)
                )
            code_lines = [str(x) for x in block["lines"]]
            for line in code_lines or [""]:
                parts.append(
                    body_paragraph(line, indent_pt=36.0, monospace=True, style=style)
                )
            continue

    return "\n".join(parts)


def replace_named_placeholders(
    html_text: str,
    header_items: list[tuple[str, str]],
    report: dict[str, object],
) -> str:
    values_by_label = collect_values_by_label(header_items)
    title = pick_title(header_items)
    metadata_table = render_metadata_table(header_items)
    placeholder_hits: list[str] = []
    new_html = html_text

    for marker in TITLE_MARKERS:
        if marker in new_html:
            new_html = new_html.replace(marker, ascii_html(title))
            placeholder_hits.append(marker)

    if METADATA_TABLE_MARKER in new_html:
        new_html = new_html.replace(METADATA_TABLE_MARKER, metadata_table)
        placeholder_hits.append(METADATA_TABLE_MARKER)

    for label, values in values_by_label.items():
        marker = "{{" + label + "}}"
        if marker not in new_html:
            continue
        for value in values:
            if marker not in new_html:
                break
            new_html = new_html.replace(marker, ascii_html(value), 1)
            placeholder_hits.append(marker)

    report["placeholder_replacements"] = placeholder_hits
    return new_html


def choose_section(sections: list[re.Match[str]]) -> int:
    hinted_indices: list[int] = []
    for index, match in enumerate(sections):
        section_text = extract_text(match.group("body")).lower()
        if any(hint in section_text for hint in BODY_HINTS):
            hinted_indices.append(index)
    if hinted_indices:
        return hinted_indices[-1]
    return len(sections) - 1


def inject_body_after_last_table(section_body: str, body_inner: str) -> str | None:
    tables = list(TABLE_RE.finditer(section_body))
    if not tables:
        return None

    last_table = tables[-1]
    prefix = section_body[: last_table.end()]
    suffix = section_body[last_table.end() :]
    suffix_text = extract_text(suffix)
    if len(suffix_text) > 160 and "{{CONTENT}}" not in suffix:
        return None
    return prefix + "\n" + body_inner + "\n"


def replace_body(
    source_html: str,
    blocks: list[dict[str, object]],
    report: dict[str, object],
    style: dict[str, object],
) -> str:
    body_inner = render_body_inner(blocks, style)
    for marker in CONTENT_MARKERS:
        if marker in source_html:
            report["body_strategy"] = f"marker:{marker}"
            return source_html.replace(marker, body_inner)

    sections = list(SECTION_RE.finditer(source_html))
    report["sections"] = [
        {
            "index": index + 1,
            "class": match.group("class"),
            "preview": extract_text(match.group("body"))[:120],
        }
        for index, match in enumerate(sections)
    ]

    if not sections:
        report["body_strategy"] = "append-before-body-close"
        block = "<div class=WordSection1 style='layout-grid:15.6pt'>\n" + body_inner + "\n</div>\n"
        return re.sub(r"(?is)</body>", block + "</body>", source_html, count=1)

    section_index = choose_section(sections)
    target = sections[section_index]
    open_tag = target.group(1)
    inner = target.group("body")
    close_tag = target.group(5)

    if len(sections) == 1:
        injected = inject_body_after_last_table(inner, body_inner)
        if injected is not None:
            report["body_strategy"] = "single-section-after-last-table"
            replacement = open_tag + injected + close_tag
        else:
            report["body_strategy"] = "single-section-replace"
            replacement = open_tag + "\n" + body_inner + "\n" + close_tag
    else:
        report["body_strategy"] = f"replace-section-{section_index + 1}"
        replacement = open_tag + "\n" + body_inner + "\n" + close_tag

    return source_html[: target.start()] + replacement + source_html[target.end() :]


def update_title(html_text: str, header_items: list[tuple[str, str]]) -> str:
    title = pick_title(header_items)
    replacement = f"<title>{ascii_html(title)}</title>"
    if TITLE_RE.search(html_text):
        return TITLE_RE.sub(replacement, html_text, count=1)
    return html_text


def build_report(
    report: dict[str, object],
    header_items: list[tuple[str, str]],
) -> str:
    lines = [
        "# Template Inference Report",
        "",
        f"- template: {report['template_path']}",
        f"- html_part_location: {report.get('html_part_location', '')}",
        f"- body_strategy: {report.get('body_strategy', 'unknown')}",
        f"- metadata_fields: {len(header_items)}",
        f"- image_count: {report.get('image_count', 0)}",
        f"- mermaid_count: {report.get('mermaid_count', 0)}",
        f"- mermaid_renderer: {report.get('mermaid_renderer', 'none')}",
        "",
        "## Placeholder Replacements",
    ]

    placeholder_hits = list(report.get("placeholder_replacements", []))
    if placeholder_hits:
        lines.extend(f"- {item}" for item in placeholder_hits)
    else:
        lines.append("- none")

    lines.extend(["", "## Table Replacements"])
    table_replacements = list(report.get("table_replacements", []))
    if table_replacements:
        for item in table_replacements:
            lines.append(
                "- table {table} row {row} col {column}: {label} -> {value}".format(**item)
            )
    else:
        lines.append("- none")

    lines.extend(["", "## Sections"])
    sections = list(report.get("sections", []))
    if sections:
        for item in sections:
            preview = item["preview"].replace("\n", " ").strip()
            lines.append(f"- #{item['index']} class={item['class']} preview={preview}")
    else:
        lines.append("- none")

    unreplaced: list[str] = []
    values_by_label = collect_values_by_label(header_items)
    matched_labels = {item["label"] for item in table_replacements}
    patent_type_merged = any(
        item["label"] == "专利类型" and "PCT" in str(item["value"])
        for item in table_replacements
    )
    for label in values_by_label:
        if label == "PCT国际申请" and patent_type_merged:
            continue
        marker = "{{" + label + "}}"
        if label not in matched_labels and marker not in placeholder_hits:
            unreplaced.append(label)

    lines.extend(["", "## Unreplaced Metadata Labels"])
    if unreplaced:
        lines.extend(f"- {label}" for label in unreplaced)
    else:
        lines.append("- none")
    lines.append("")
    return "\n".join(lines)


def mermaid_cli_command() -> list[str] | None:
    mmdc = shutil.which("mmdc")
    if mmdc:
        return [mmdc]

    cached_mmdc = sorted(
        (Path.home() / ".npm" / "_npx").glob("*/node_modules/.bin/mmdc"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    if cached_mmdc:
        return [str(cached_mmdc[0])]

    npx = shutil.which("npx")
    if npx:
        return [npx, "-y", "-p", "@mermaid-js/mermaid-cli", "mmdc"]

    return None


def mermaid_asset_base(html_part_location: str) -> str:
    location = html_part_location.strip()
    if not location:
        return "file:///C:/generated-output.files/"

    if re.search(r"(?i)\.html?$", location):
        return re.sub(r"(?i)\.html?$", ".files/", location)

    if location.endswith("/"):
        return location

    head, _, _tail = location.rpartition("/")
    if head:
        return head + "/"
    return location + "/"


def image_asset_type(image_path: Path) -> tuple[str, str]:
    mime_type, _encoding = mimetypes.guess_type(image_path.name)
    if mime_type and "/" in mime_type:
        maintype, subtype = mime_type.split("/", 1)
        if maintype == "image":
            return maintype, subtype
    suffix = image_path.suffix.lower()
    if suffix in {".jpg", ".jpeg"}:
        return "image", "jpeg"
    if suffix == ".gif":
        return "image", "gif"
    if suffix == ".bmp":
        return "image", "bmp"
    if suffix == ".webp":
        return "image", "webp"
    return "image", "png"


def prepare_local_image_assets(
    blocks: list[dict[str, object]],
    input_path: Path,
    html_part_location: str,
    report: dict[str, object],
) -> list[dict[str, object]]:
    image_blocks = [block for block in blocks if block.get("type") == "image"]
    report["image_count"] = len(image_blocks)
    if not image_blocks:
        return []

    assets: list[dict[str, object]] = []
    base_location = mermaid_asset_base(html_part_location)
    input_dir = input_path.resolve().parent

    for index, block in enumerate(image_blocks, start=1):
        src = str(block.get("src", "")).strip()
        if not src:
            raise RuntimeError(f"第 {index} 个图片块缺少图片路径")
        image_path = (input_dir / src).resolve()
        if not image_path.is_file():
            raise RuntimeError(f"图片文件不存在: {image_path}")
        suffix = image_path.suffix.lower() or ".png"
        file_name = f"image-{index:03d}{suffix}"
        content_location = base_location + file_name
        maintype, subtype = image_asset_type(image_path)
        block["content_location"] = content_location
        if not str(block.get("alt", "")).strip():
            block["alt"] = image_path.stem
        assets.append(
            {
                "content_location": content_location,
                "filename": file_name,
                "data": image_path.read_bytes(),
                "maintype": maintype,
                "subtype": subtype,
            }
        )

    return assets


def render_mermaid_png(command_prefix: list[str], mermaid_code: str, output_path: Path) -> None:
    input_path = output_path.with_suffix(".mmd")
    puppeteer_config_path = output_path.with_suffix(".puppeteer.json")
    input_path.write_text(mermaid_code, encoding="utf-8")
    puppeteer_config_path.write_text(
        json.dumps(
            {
                "headless": True,
                "args": [
                    "--no-sandbox",
                    "--disable-setuid-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-gpu",
                    "--single-process",
                    "--no-zygote",
                    "--disable-software-rasterizer",
                    "--no-proxy-server",
                ]
            }
        ),
        encoding="utf-8",
    )

    command = [
        *command_prefix,
        "-i",
        str(input_path),
        "-o",
        str(output_path),
        "-p",
        str(puppeteer_config_path),
        "-t",
        "neutral",
        "-b",
        "white",
    ]
    env = os.environ.copy()
    for key in ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy", "ALL_PROXY", "all_proxy"):
        env.pop(key, None)
    env["NO_PROXY"] = "*"
    env["no_proxy"] = "*"
    result = subprocess.run(command, capture_output=True, text=True, env=env)
    if result.returncode == 0 and output_path.exists():
        return

    detail = (result.stderr or result.stdout).strip()
    if len(detail) > 800:
        detail = detail[-800:]
    message = (
        "Mermaid 渲染失败。请确保本机可执行 `mmdc`，或允许 `npx -p @mermaid-js/mermaid-cli mmdc` 拉起 Mermaid CLI。"
        f"\n命令: {' '.join(command)}"
    )
    if detail:
        message += f"\n输出: {detail}"
    raise RuntimeError(message)


def prepare_mermaid_assets(
    blocks: list[dict[str, object]],
    html_part_location: str,
    report: dict[str, object],
) -> list[dict[str, object]]:
    mermaid_blocks = [block for block in blocks if block.get("type") == "mermaid"]
    if not mermaid_blocks:
        report["mermaid_count"] = 0
        return []

    command_prefix = mermaid_cli_command()
    if command_prefix is None:
        raise RuntimeError(
            "检测到 Mermaid 代码块，但当前环境没有 `mmdc`，也没有可用的 `npx`。"
        )

    assets: list[dict[str, object]] = []
    base_location = mermaid_asset_base(html_part_location)
    report["mermaid_count"] = len(mermaid_blocks)
    report["mermaid_renderer"] = " ".join(command_prefix)

    with tempfile.TemporaryDirectory(prefix="md_to_word_mermaid_") as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        for index, block in enumerate(mermaid_blocks, start=1):
            mermaid_code = "\n".join(str(line) for line in block.get("lines", []))
            file_name = f"mermaid-{index:03d}.png"
            output_path = temp_dir / file_name
            render_mermaid_png(command_prefix, mermaid_code, output_path)

            content_location = base_location + file_name
            block["content_location"] = content_location
            block["alt"] = f"Mermaid 图{index}"
            assets.append(
                {
                    "content_location": content_location,
                    "filename": file_name,
                    "data": output_path.read_bytes(),
                }
            )

    return assets


def attach_related_image_part(msg: EmailMessage, asset: dict[str, object]) -> None:
    part = EmailMessage()
    part.set_content(
        asset["data"],
        maintype=str(asset.get("maintype", "image")),
        subtype=str(asset.get("subtype", "png")),
        cte="base64",
    )
    part["Content-Location"] = asset["content_location"]
    part["Content-Disposition"] = f'inline; filename="{asset["filename"]}"'
    part["Content-ID"] = f'<{asset["filename"]}>'
    msg.attach(part)


def main() -> int:
    preset_names = list_preset_names()
    parser = argparse.ArgumentParser(description="Render markdown into a reusable Word-exported MHT template")
    parser.add_argument("--input", "-i", required=True, help="Input markdown path")
    parser.add_argument("--template", "-t", required=True, help="Template MHT path")
    parser.add_argument("--output", "-o", required=True, help="Output MHT path")
    parser.add_argument("--report", help="Optional template inference report output path")
    parser.add_argument(
        "--style-preset",
        default=DEFAULT_STYLE_PRESET,
        choices=preset_names,
        help="Rendering style preset for body content",
    )
    args = parser.parse_args()

    md_text = Path(args.input).read_text(encoding="utf-8")
    header_items, blocks = parse_markdown(md_text)
    style = get_render_style(args.style_preset)

    msg = BytesParser(policy=policy.default).parsebytes(Path(args.template).read_bytes())
    html_part = next(
        part
        for part in msg.walk()
        if part.get_content_type() == "text/html"
        and "header.htm" not in str(part.get("Content-Location", ""))
    )

    report: dict[str, object] = {
        "template_path": str(Path(args.template).resolve()),
        "html_part_location": str(html_part.get("Content-Location", "")),
    }

    local_image_assets = prepare_local_image_assets(
        blocks,
        Path(args.input),
        report["html_part_location"],
        report,
    )
    mermaid_assets = prepare_mermaid_assets(blocks, report["html_part_location"], report)
    html_source = html_part.get_content()
    html_source = update_title(html_source, header_items)
    html_source = replace_named_placeholders(html_source, header_items, report)
    html_source = fill_cover_tables(html_source, header_items, report)
    html_source = replace_body(html_source, blocks, report, style)

    content_location = html_part.get("Content-Location")
    html_part.set_content(
        html_source,
        subtype="html",
        charset="us-ascii",
        cte="quoted-printable",
    )
    if content_location:
        if html_part.get("Content-Location"):
            html_part.replace_header("Content-Location", content_location)
        else:
            html_part.add_header("Content-Location", content_location)

    for asset in [*local_image_assets, *mermaid_assets]:
        attach_related_image_part(msg, asset)

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    buffer = BytesIO()
    BytesGenerator(buffer, policy=policy.default.clone(linesep="\r\n")).flatten(msg)
    output_path.write_bytes(buffer.getvalue())

    if args.report:
        report_path = Path(args.report)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(build_report(report, header_items), encoding="utf-8")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
