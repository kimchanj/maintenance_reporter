from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import List

from docx import Document


@dataclass
class ParsedVpnRow:
    request_date: str
    content_text: str
    raw_block: str


@dataclass
class ParsedVpnResult:
    rows: List[ParsedVpnRow]
    warnings: List[str]


def _compact(text: str) -> str:
    return "".join(text.split())


def _normalize_lines(text: str) -> List[str]:
    lines: List[str] = []
    for line in text.splitlines():
        cleaned = " ".join(line.split()).strip()
        if cleaned:
            lines.append(cleaned)
    return lines


def _extract_date_text(line: str) -> str | None:
    digits = re.findall(r"\d+", line)
    if len(digits) < 3:
        return None

    year = int(digits[0])
    month = int(digits[1])
    day = int(digits[2])
    if year < 100:
        year += 2000
    if not 1 <= month <= 12 or not 1 <= day <= 31:
        return None
    return f"{year:04d}.{month:02d}.{day:02d}"


def _find_work_history_text(doc: Document) -> str:
    for table in doc.tables:
        for row in table.rows:
            if len(row.cells) < 2:
                continue
            if _compact(row.cells[0].text) == "작업내역":
                return row.cells[1].text.strip()
    raise ValueError("Could not find the 작업내역 cell in the VPN document.")


def parse_vpn_report(vpn_report_path: Path) -> ParsedVpnResult:
    if vpn_report_path.suffix.lower() != ".docx":
        raise ValueError("VPN report must be .docx")

    doc = Document(str(vpn_report_path))
    work_history_text = _find_work_history_text(doc)
    if not work_history_text.strip():
        raise ValueError("The 작업내역 cell is empty in the VPN document.")

    blocks = [block.strip() for block in re.split(r"\n\s*\n+", work_history_text) if block.strip()]
    rows: List[ParsedVpnRow] = []
    warnings: List[str] = []

    for index, block in enumerate(blocks, start=1):
        lines = _normalize_lines(block)
        if len(lines) < 2:
            warnings.append(f"Block {index}: skipped because it does not have both a date line and a content line.")
            continue

        request_date = _extract_date_text(lines[0])
        if not request_date:
            warnings.append(f"Block {index}: skipped because a request date could not be parsed from '{lines[0]}'.")
            continue

        content_text = lines[1].strip()
        if not content_text:
            warnings.append(f"Block {index}: skipped because the content line is empty.")
            continue

        if len(lines) > 2:
            warnings.append(f"Block {index}: only the second line was used as 일반 내역 내용.")

        rows.append(
            ParsedVpnRow(
                request_date=request_date,
                content_text=content_text,
                raw_block=block,
            )
        )

    if not rows:
        details = "\n".join(warnings[:10]) if warnings else "No valid two-line blocks were found."
        raise ValueError(f"Could not parse any general-report rows from the VPN document.\n{details}")

    return ParsedVpnResult(rows=rows, warnings=warnings)
