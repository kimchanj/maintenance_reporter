from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from app.services.excel_reader import read_maintenance_rows
from app.services.word_writer import (
    extract_template_fields,
    fill_template_and_save,
    validate_generated_report,
)


@dataclass
class ReportResult:
    output_path: Path
    validated_count: int


def generate_report(excel_path: Path, template_path: Path) -> ReportResult:
    """End-to-end pipeline."""
    template_fields = extract_template_fields(template_path)
    rows = read_maintenance_rows(excel_path, template_fields)
    if not rows:
        raise ValueError("No rows with 진행상태=완료 found in Excel file.")

    out_path = fill_template_and_save(template_path, rows)
    validation = validate_generated_report(template_path, out_path, rows)
    if not validation.is_match:
        details = "\n".join(validation.mismatches[:10])
        if len(validation.mismatches) > 10:
            details += f"\n... and {len(validation.mismatches) - 10} more mismatches."
        raise ValueError(
            "Cross validation failed between Excel and generated Word.\n"
            f"Excel rows: {validation.excel_count}, Word rows: {validation.word_count}\n"
            f"{details}"
        )

    return ReportResult(output_path=out_path, validated_count=validation.excel_count)
