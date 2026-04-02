#!/usr/bin/env python3
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from zipfile import BadZipFile, ZipFile


@dataclass(frozen=True)
class OutputPaths:
    output_mht: Path
    output_docx: Path
    processed_docx: Path
    template_report: Path | None


@dataclass(frozen=True)
class StagePaths:
    stage_dir: Path
    template_input_docx: Path
    template_normalized_mht: Path
    input_mht: Path
    raw_docx: Path
    processed_docx: Path


def resolve_path(path_value: str, *, allow_missing: bool) -> Path:
    path = Path(path_value).expanduser()
    if path.is_absolute():
        resolved = path.resolve(strict=False)
    else:
        resolved = (Path.cwd() / path).resolve(strict=False)
    if not allow_missing and not resolved.exists():
        raise FileNotFoundError(f"path not found: {resolved}")
    return resolved


def resolve_output_paths(
    input_abs: Path,
    output_mht: str,
    output_docx: str,
    processed_docx: str,
    template_report: str,
) -> OutputPaths:
    base_no_ext = input_abs.with_suffix("")
    resolved_report = resolve_path(template_report, allow_missing=True) if template_report else None
    return OutputPaths(
        output_mht=resolve_path(output_mht, allow_missing=True) if output_mht else base_no_ext.with_suffix(".mht"),
        output_docx=resolve_path(output_docx, allow_missing=True) if output_docx else base_no_ext.with_suffix(".docx"),
        processed_docx=resolve_path(processed_docx, allow_missing=True)
        if processed_docx
        else base_no_ext.with_suffix(".wordmath.docx"),
        template_report=resolved_report,
    )


def build_stage_paths(windows_temp_root: Path) -> StagePaths:
    stage_id = f"{datetime.now():%Y%m%d-%H%M%S}-{os.getpid()}"
    stage_dir = windows_temp_root / "repo2patent-md2word" / stage_id
    return StagePaths(
        stage_dir=stage_dir,
        template_input_docx=stage_dir / "template-input.docx",
        template_normalized_mht=stage_dir / "template-normalized.mht",
        input_mht=stage_dir / "input.mht",
        raw_docx=stage_dir / "raw.docx",
        processed_docx=stage_dir / "processed.docx",
    )


def ensure_parent_dirs(*paths: Path | None) -> None:
    for path in paths:
        if path is None:
            continue
        path.parent.mkdir(parents=True, exist_ok=True)


def validate_template_args(
    *,
    template_mht: str,
    template_docx: str,
    template_spec: str,
    template_out: str,
) -> None:
    template_sources = [value for value in (template_mht, template_docx, template_spec) if value]
    if len(template_sources) > 1:
        raise ValueError("only one of --template-mht / --template-docx / --template-spec can be used")
    if template_out and template_mht:
        raise ValueError("--template-out is only meaningful with --template-docx or --template-spec")


def validate_docx_zip(path: Path) -> bool:
    if not path.exists() or path.stat().st_size <= 0:
        return False
    try:
        with ZipFile(path) as archive:
            names = set(archive.namelist())
    except BadZipFile:
        return False
    return "[Content_Types].xml" in names and "word/document.xml" in names


def generated_template_path(template_spec_abs: Path) -> Path:
    if template_spec_abs.suffix:
        return template_spec_abs.with_suffix(".generated.mht")
    return template_spec_abs.parent / f"{template_spec_abs.name}.generated.mht"
