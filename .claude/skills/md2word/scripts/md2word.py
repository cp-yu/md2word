#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

from pipeline_common import (
    build_stage_paths,
    ensure_parent_dirs,
    generated_template_path,
    resolve_output_paths,
    resolve_path,
    validate_docx_zip,
    validate_template_args,
)
from style_presets import list_preset_names


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent


@dataclass(frozen=True)
class CommandResult:
    returncode: int
    stdout: str
    stderr: str
    timed_out: bool = False


def build_parser() -> argparse.ArgumentParser:
    preset_names = list_preset_names()
    parser = argparse.ArgumentParser(description="在 Windows + Word 流水线中将 Markdown 导出为 .mht、.docx 和 .wordmath.docx")
    parser.add_argument("--input", "-i", required=True, help="Input markdown path")
    parser.add_argument("--template-mht", "-t", default="", help="Use an existing MHT template")
    parser.add_argument("--template-docx", default="", help="Normalize a DOCX template to MHT first")
    parser.add_argument("--template-spec", default="", help="Generate a reusable MHT template from template spec")
    parser.add_argument("--template-out", default="", help="Output path for normalized/generated template MHT")
    parser.add_argument("--template-report", default="", help="Optional template inference report output path")
    parser.add_argument("--output-mht", default="", help="Intermediate rendered MHT path")
    parser.add_argument("--output-docx", default="", help="Raw DOCX output path")
    parser.add_argument("--processed-docx", default="", help="WordMath DOCX output path")
    parser.add_argument("--macro-name", default="", help="Optional Word macro to replace built-in formula conversion")
    parser.add_argument(
        "--style-preset",
        default="",
        choices=["", *preset_names],
        help="Optional rendering style preset",
    )
    parser.add_argument("--visible", action="store_true", help="Show Word window for debugging")
    parser.add_argument(
        "--timeout-seconds",
        type=int,
        default=int(os.environ.get("WORD_PIPELINE_TIMEOUT_SECONDS", "180")),
        help="Timeout for each Word stage",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        validate_template_args(
            template_mht=args.template_mht,
            template_docx=args.template_docx,
            template_spec=args.template_spec,
            template_out=args.template_out,
        )
        input_abs = resolve_path(args.input, allow_missing=False)
        output_paths = resolve_output_paths(
            input_abs=input_abs,
            output_mht=args.output_mht,
            output_docx=args.output_docx,
            processed_docx=args.processed_docx,
            template_report=args.template_report,
        )
        ensure_parent_dirs(
            output_paths.output_mht,
            output_paths.output_docx,
            output_paths.processed_docx,
            output_paths.template_report,
        )

        powershell_exe = pick_powershell_exe()
        stage_paths = build_stage_paths(temp_root())
        stage_paths.stage_dir.mkdir(parents=True, exist_ok=True)

        resolved_template_abs = resolve_template(args=args, powershell_exe=powershell_exe, stage_paths=stage_paths)
        if not resolved_template_abs.is_file():
            raise RuntimeError(f"resolved template not found: {resolved_template_abs}")

        render_args = [
            "--input",
            str(input_abs),
            "--template",
            str(resolved_template_abs),
            "--output",
            str(output_paths.output_mht),
        ]
        if args.style_preset:
            render_args.extend(["--style-preset", args.style_preset])
        if output_paths.template_report is not None:
            render_args.extend(["--report", str(output_paths.template_report)])
        run_python_script(SCRIPT_DIR / "render_mht.py", render_args, "render mht")

        shutil.copy2(output_paths.output_mht, stage_paths.input_mht)
        staged_pipeline = stage_powershell_worker(stage_paths.stage_dir, "word_mht_pipeline.ps1")
        pipeline_args = [
            "-InputMht",
            str(stage_paths.input_mht),
            "-OutputDocx",
            str(stage_paths.raw_docx),
            "-ProcessedDocx",
            str(stage_paths.processed_docx),
        ]
        if args.macro_name:
            pipeline_args.extend(["-MacroName", args.macro_name])
        if args.visible:
            pipeline_args.append("-Visible")
        result = run_powershell_script(
            staged_pipeline,
            pipeline_args,
            timeout_seconds=args.timeout_seconds,
            powershell_exe=powershell_exe,
        )
        handle_pipeline_result(result, stage_paths)

        if not validate_docx_zip(stage_paths.raw_docx):
            raise RuntimeError(f"invalid raw DOCX generated by Word pipeline: {stage_paths.raw_docx}")
        if not validate_docx_zip(stage_paths.processed_docx):
            raise RuntimeError(f"invalid processed DOCX generated by Word pipeline: {stage_paths.processed_docx}")

        shutil.copy2(stage_paths.raw_docx, output_paths.output_docx)
        shutil.copy2(stage_paths.processed_docx, output_paths.processed_docx)

        print(f"[ok] template: {resolved_template_abs}")
        if output_paths.template_report is not None:
            print(f"[ok] template report: {output_paths.template_report}")
        print(f"[ok] rendered mht: {output_paths.output_mht}")
        print(f"[ok] raw docx: {output_paths.output_docx}")
        print(f"[ok] wordmath docx: {output_paths.processed_docx}")
        return 0
    except (FileNotFoundError, OSError, RuntimeError, ValueError) as exc:
        print(f"[error] {exc}", file=sys.stderr)
        return 1


def resolve_template(*, args: argparse.Namespace, powershell_exe: str, stage_paths) -> Path:
    if args.template_mht:
        return resolve_path(args.template_mht, allow_missing=False)

    if args.template_docx:
        template_docx_abs = resolve_path(args.template_docx, allow_missing=False)
        shutil.copy2(template_docx_abs, stage_paths.template_input_docx)
        staged_worker = stage_powershell_worker(stage_paths.stage_dir, "word_template_to_mht.ps1")
        result = run_powershell_script(
            staged_worker,
            [
                "-InputDocx",
                str(stage_paths.template_input_docx),
                "-OutputMht",
                str(stage_paths.template_normalized_mht),
                *(["-Visible"] if args.visible else []),
            ],
            timeout_seconds=args.timeout_seconds,
            powershell_exe=powershell_exe,
        )
        if result.returncode != 0:
            raise RuntimeError(f"template docx -> mht failed with exit code: {result.returncode}")
        if not stage_paths.template_normalized_mht.is_file() or stage_paths.template_normalized_mht.stat().st_size <= 0:
            raise RuntimeError(
                f"template docx -> mht did not produce a valid mht: {stage_paths.template_normalized_mht}"
            )
        if args.template_out:
            template_out_abs = resolve_path(args.template_out, allow_missing=True)
            ensure_parent_dirs(template_out_abs)
            shutil.copy2(stage_paths.template_normalized_mht, template_out_abs)
            return template_out_abs
        return stage_paths.template_normalized_mht

    if args.template_spec:
        template_spec_abs = resolve_path(args.template_spec, allow_missing=False)
        template_out_abs = (
            resolve_path(args.template_out, allow_missing=True)
            if args.template_out
            else generated_template_path(template_spec_abs)
        )
        ensure_parent_dirs(template_out_abs)
        template_args = [
            "--spec",
            str(template_spec_abs),
            "--output",
            str(template_out_abs),
        ]
        if args.style_preset:
            template_args.extend(["--style-preset", args.style_preset])
        run_python_script(SCRIPT_DIR / "generate_template_mht.py", template_args, "generate template mht")
        return template_out_abs

    return SKILL_ROOT / "resources" / "专利交底书模板.mht"


def pick_powershell_exe() -> str:
    env_value = os.environ.get("POWERSHELL_EXE")
    if env_value:
        return env_value

    for command_name in ("powershell.exe", "pwsh.exe", "powershell", "pwsh"):
        command_path = shutil.which(command_name)
        if command_path:
            return command_path

    raise FileNotFoundError("PowerShell executable not found")


def temp_root() -> Path:
    env_value = os.environ.get("TEMP") or os.environ.get("TMP")
    if env_value:
        return Path(env_value)
    return Path(tempfile.gettempdir())


def run_process(command: Sequence[str], *, timeout_seconds: int | None = None) -> CommandResult:
    try:
        completed = subprocess.run(
            list(command),
            check=False,
            text=True,
            capture_output=True,
            errors="replace",
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired as exc:
        return CommandResult(
            returncode=124,
            stdout=exc.stdout or "",
            stderr=exc.stderr or "",
            timed_out=True,
        )
    return CommandResult(
        returncode=completed.returncode,
        stdout=completed.stdout or "",
        stderr=completed.stderr or "",
    )


def print_command_output(result: CommandResult) -> None:
    if result.stdout:
        sys.stdout.write(result.stdout)
        if not result.stdout.endswith("\n"):
            sys.stdout.write("\n")
    if result.stderr:
        sys.stderr.write(result.stderr)
        if not result.stderr.endswith("\n"):
            sys.stderr.write("\n")


def run_python_script(script_path: Path, script_args: list[str], label: str) -> None:
    result = run_process([sys.executable, str(script_path), *script_args])
    print_command_output(result)
    if result.returncode != 0:
        raise RuntimeError(f"{label} failed with exit code: {result.returncode}")


def stage_powershell_worker(stage_dir: Path, worker_name: str) -> Path:
    staged_worker = stage_dir / worker_name
    shutil.copy2(SCRIPT_DIR / worker_name, staged_worker)
    shutil.copy2(SCRIPT_DIR / "word_common.ps1", stage_dir / "word_common.ps1")
    return staged_worker


def run_powershell_script(
    script_path: Path,
    script_args: list[str],
    *,
    timeout_seconds: int,
    powershell_exe: str,
) -> CommandResult:
    result = run_process(
        [
            powershell_exe,
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            to_powershell_path(script_path),
            *script_args,
        ],
        timeout_seconds=timeout_seconds,
    )
    print_command_output(result)
    return result


def handle_pipeline_result(result: CommandResult, stage_paths) -> None:
    if result.returncode == 0:
        return
    if result.timed_out and stage_paths.raw_docx.is_file() and stage_paths.processed_docx.is_file():
        print("[warn] PowerShell wrapper timed out, but Word 已生成 DOCX，按成功处理。")
        return
    raise RuntimeError(f"Windows Word pipeline failed with exit code: {result.returncode}")


def to_powershell_path(path_value: Path | str) -> str:
    path_str = str(path_value)
    if os.name == "nt":
        return path_str
    parts = path_str.split("/", 4)
    if len(parts) >= 4 and parts[0] == "" and parts[1] == "mnt" and len(parts[2]) == 1:
        drive = parts[2].upper()
        rest = "\\".join(parts[3:])
        return f"{drive}:\\{rest}"
    return path_str


if __name__ == "__main__":
    raise SystemExit(main())
