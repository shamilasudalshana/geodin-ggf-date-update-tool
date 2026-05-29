from __future__ import annotations

from pathlib import Path
from datetime import datetime

from .backend import ProcessResult, ProcessSettings


def default_report_path(output_ggf_path: Path) -> Path:
    output_ggf_path = Path(output_ggf_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return output_ggf_path.with_name(f"{output_ggf_path.stem}_report_{timestamp}.txt")


def build_user_summary(result: ProcessResult) -> str:
    lines = [
        "GeoDIN GGF Date Tool - Processing Summary",
        "=" * 45,
        "",
        f"Input GGF: {result.input_ggf}",
        f"Input Excel: {result.input_xlsx}",
        f"Output GGF: {result.output_ggf}",
        "",
        f"GGF slots found: {result.total_slots_found}",
        f"Valid Excel dates written: {result.final_dates_written}",
        f"Dummy dates written: {result.dummy_dates_written}",
        f"File size unchanged: {result.file_size_unchanged}",
        "",
    ]

    if result.duplicate_dates:
        lines.append("Duplicate dates detected:")
        for d in result.duplicate_dates:
            lines.append(f"  - {d}")
        lines.append("")
    else:
        lines.append("No duplicate dates detected.")
        lines.append("")

    return "\n".join(lines)


def write_full_report(path: Path, result: ProcessResult) -> None:
    Path(path).write_text("\n".join(result.log_lines), encoding="utf-8")
