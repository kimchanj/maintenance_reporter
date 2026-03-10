from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List

from app.services.excel_reader import MaintenanceRow, read_maintenance_rows
from app.services.vpn_reader import parse_vpn_report
from app.services.word_writer import (
    extract_template_fields,
    fill_template_and_save,
    validate_generated_report,
)


@dataclass
class ReportResult:
    output_path: Path
    validated_count: int
    warnings: List[str]


def _raise_on_validation_failure(template_path: Path, output_path: Path, rows: List[MaintenanceRow]) -> None:
    validation = validate_generated_report(template_path, output_path, rows)
    if validation.is_match:
        return

    details = "\n".join(validation.mismatches[:10])
    if len(validation.mismatches) > 10:
        details += f"\n... and {len(validation.mismatches) - 10} more mismatches."
    raise ValueError(
        "Cross validation failed between source data and generated Word.\n"
        f"Source rows: {validation.excel_count}, Word rows: {validation.word_count}\n"
        f"{details}"
    )


def generate_vpn_report(excel_path: Path, template_path: Path) -> ReportResult:
    """Create the initial VPN report from Excel data."""
    template_fields = extract_template_fields(template_path)
    rows = read_maintenance_rows(excel_path, template_fields)
    if not rows:
        raise ValueError("No rows with 진행상태=완료 found in the Excel file.")

    output_path = fill_template_and_save(template_path, rows)
    _raise_on_validation_failure(template_path, output_path, rows)
    return ReportResult(output_path=output_path, validated_count=len(rows), warnings=[])


def generate_general_report_from_vpn(vpn_report_path: Path, template_path: Path) -> ReportResult:
    """Create the general maintenance report from an edited VPN report."""
    parsed = parse_vpn_report(vpn_report_path)
    rows: List[MaintenanceRow] = []
    for item in parsed.rows:
        rows.append(
            {
                "요구일자": item.request_date,
                "처리내용": item.content_text,
                "진행상태": "완료",
            }
        )

    output_path = fill_template_and_save(template_path, rows)
    _raise_on_validation_failure(template_path, output_path, rows)
    return ReportResult(output_path=output_path, validated_count=len(rows), warnings=parsed.warnings)
