#!/usr/bin/env python3
"""Prepare warehouse/customer/OD-link datasets from customer_cleaned.xlsx.

The script intentionally uses only the Python standard library so it can run in
minimal environments where Excel helper packages such as openpyxl are absent.
"""

from __future__ import annotations

import csv
import hashlib
import json
import math
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable
from zipfile import ZipFile
import xml.etree.ElementTree as ET

REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = REPO_ROOT / "customer_cleaned.xlsx"
DEFAULT_OUTPUT_DIR = REPO_ROOT / "data" / "processed"
DEFAULT_REPORT_DIR = REPO_ROOT / "reports"

MAIN_SHEET_NAME = "Sheet1"
HEADER_ALIASES = {
    "No": "source_row_no",
    "물류센터": "warehouse_name",
    "주소": "warehouse_address",
    "출발지위도": "warehouse_lat",
    "출발지경도": "warehouse_lon",
    "거래처명": "customer_name",
    "상세주소": "customer_address",
    "원주소(SAP)": "customer_original_address_sap",
    "위도": "customer_lat",
    "경도": "customer_lon",
}
WAREHOUSE_COLUMNS = [
    "warehouse_id",
    "warehouse_name",
    "warehouse_address",
    "warehouse_lat",
    "warehouse_lon",
]
CUSTOMER_COLUMNS = [
    "customer_id",
    "customer_name",
    "customer_address",
    "customer_original_address_sap",
    "customer_lat",
    "customer_lon",
]
OD_COLUMNS = [
    "link_id",
    "source_row_no",
    "warehouse_id",
    "customer_id",
    "warehouse_name",
    "customer_name",
    "origin_lat",
    "origin_lon",
    "destination_lat",
    "destination_lon",
]
KOREA_BOUNDS = {
    "lat_min": 33.0,
    "lat_max": 39.5,
    "lon_min": 124.0,
    "lon_max": 132.0,
}

XML_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass(frozen=True)
class SheetInfo:
    name: str
    path: str
    dimension: str | None


def normalize_text(value: object) -> str:
    """Normalize human-entered text without destroying Korean strings."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def parse_float(value: object) -> float | None:
    text = normalize_text(value)
    if not text:
        return None
    try:
        parsed = float(text)
    except ValueError:
        return None
    if not math.isfinite(parsed):
        return None
    return parsed


def coordinate_text(value: object) -> str:
    parsed = parse_float(value)
    if parsed is None:
        return ""
    return f"{parsed:.8f}"


def make_id(prefix: str, parts: Iterable[object]) -> str:
    normalized = "|".join(normalize_text(part).casefold() for part in parts)
    digest = hashlib.sha1(normalized.encode("utf-8")).hexdigest()[:10]
    return f"{prefix}_{digest}"


def column_letters(cell_ref: str) -> str:
    return "".join(ch for ch in cell_ref if ch.isalpha())


def column_index(cell_ref: str) -> int:
    result = 0
    for ch in column_letters(cell_ref):
        result = result * 26 + (ord(ch.upper()) - ord("A") + 1)
    return result - 1


def load_shared_strings(xlsx: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in xlsx.namelist():
        return []
    root = ET.fromstring(xlsx.read("xl/sharedStrings.xml"))
    shared_strings: list[str] = []
    for item in root.findall("main:si", XML_NS):
        text_parts = [node.text or "" for node in item.findall(".//main:t", XML_NS)]
        shared_strings.append("".join(text_parts))
    return shared_strings


def list_sheets(xlsx: ZipFile) -> list[SheetInfo]:
    workbook = ET.fromstring(xlsx.read("xl/workbook.xml"))
    relationships = ET.fromstring(xlsx.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in relationships}
    sheets: list[SheetInfo] = []
    for sheet in workbook.findall("main:sheets/main:sheet", XML_NS):
        rel_id = sheet.attrib[f"{{{REL_NS}}}id"]
        target = rel_map[rel_id]
        path = "xl/" + target.lstrip("/") if not target.startswith("xl/") else target
        root = ET.fromstring(xlsx.read(path))
        dimension = root.find("main:dimension", XML_NS)
        sheets.append(
            SheetInfo(
                name=sheet.attrib["name"],
                path=path,
                dimension=dimension.attrib.get("ref") if dimension is not None else None,
            )
        )
    return sheets


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    value_node = cell.find("main:v", XML_NS)
    inline_node = cell.find("main:is", XML_NS)
    if cell_type == "inlineStr" and inline_node is not None:
        return "".join(node.text or "" for node in inline_node.findall(".//main:t", XML_NS))
    if value_node is None:
        return ""
    raw_value = value_node.text or ""
    if cell_type == "s":
        return shared_strings[int(raw_value)]
    return raw_value


def read_sheet_rows(xlsx_path: Path, sheet_name: str) -> tuple[list[dict[str, str]], list[SheetInfo]]:
    with ZipFile(xlsx_path) as xlsx:
        shared_strings = load_shared_strings(xlsx)
        sheets = list_sheets(xlsx)
        matching = [sheet for sheet in sheets if sheet.name == sheet_name]
        if not matching:
            available = ", ".join(sheet.name for sheet in sheets)
            raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {available}")
        sheet = matching[0]
        root = ET.fromstring(xlsx.read(sheet.path))
        rows: list[list[str]] = []
        for row_node in root.findall("main:sheetData/main:row", XML_NS):
            row_values: list[str] = []
            for cell in row_node.findall("main:c", XML_NS):
                index = column_index(cell.attrib["r"])
                while len(row_values) <= index:
                    row_values.append("")
                row_values[index] = normalize_text(cell_value(cell, shared_strings))
            rows.append(row_values)
    if not rows:
        return [], sheets
    headers = [HEADER_ALIASES.get(header, header) for header in rows[0]]
    records: list[dict[str, str]] = []
    for raw_row in rows[1:]:
        padded = raw_row + [""] * (len(headers) - len(raw_row))
        record = {headers[index]: padded[index] for index in range(len(headers))}
        if any(normalize_text(value) for value in record.values()):
            records.append(record)
    return records, sheets


def in_korea(lat: float | None, lon: float | None) -> bool:
    if lat is None or lon is None:
        return False
    return (
        KOREA_BOUNDS["lat_min"] <= lat <= KOREA_BOUNDS["lat_max"]
        and KOREA_BOUNDS["lon_min"] <= lon <= KOREA_BOUNDS["lon_max"]
    )


def build_datasets(records: list[dict[str, str]]) -> tuple[list[dict[str, str]], list[dict[str, str]], list[dict[str, str]], dict[str, object]]:
    warehouses_by_id: dict[str, dict[str, str]] = {}
    customers_by_id: dict[str, dict[str, str]] = {}
    od_links: list[dict[str, str]] = []
    duplicate_link_counter: Counter[tuple[str, str]] = Counter()
    invalid_rows: list[dict[str, str]] = []
    warehouse_name_to_ids: defaultdict[str, set[str]] = defaultdict(set)
    customer_name_to_ids: defaultdict[str, set[str]] = defaultdict(set)

    for row_index, record in enumerate(records, start=2):
        warehouse_lat = parse_float(record.get("warehouse_lat"))
        warehouse_lon = parse_float(record.get("warehouse_lon"))
        customer_lat = parse_float(record.get("customer_lat"))
        customer_lon = parse_float(record.get("customer_lon"))

        validation_errors: list[str] = []
        if not normalize_text(record.get("warehouse_name")):
            validation_errors.append("missing_warehouse_name")
        if not normalize_text(record.get("customer_name")):
            validation_errors.append("missing_customer_name")
        if not in_korea(warehouse_lat, warehouse_lon):
            validation_errors.append("invalid_or_out_of_korea_warehouse_coordinate")
        if not in_korea(customer_lat, customer_lon):
            validation_errors.append("invalid_or_out_of_korea_customer_coordinate")
        if validation_errors:
            invalid_rows.append(
                {
                    "excel_row": str(row_index),
                    "source_row_no": normalize_text(record.get("source_row_no")),
                    "warehouse_name": normalize_text(record.get("warehouse_name")),
                    "customer_name": normalize_text(record.get("customer_name")),
                    "errors": ",".join(validation_errors),
                }
            )

        warehouse_id = make_id(
            "wh",
            [
                record.get("warehouse_name"),
                record.get("warehouse_address"),
                coordinate_text(record.get("warehouse_lat")),
                coordinate_text(record.get("warehouse_lon")),
            ],
        )
        customer_id = make_id(
            "cu",
            [
                record.get("customer_name"),
                record.get("customer_address"),
                record.get("customer_original_address_sap"),
                coordinate_text(record.get("customer_lat")),
                coordinate_text(record.get("customer_lon")),
            ],
        )
        warehouses_by_id.setdefault(
            warehouse_id,
            {
                "warehouse_id": warehouse_id,
                "warehouse_name": normalize_text(record.get("warehouse_name")),
                "warehouse_address": normalize_text(record.get("warehouse_address")),
                "warehouse_lat": coordinate_text(record.get("warehouse_lat")),
                "warehouse_lon": coordinate_text(record.get("warehouse_lon")),
            },
        )
        customers_by_id.setdefault(
            customer_id,
            {
                "customer_id": customer_id,
                "customer_name": normalize_text(record.get("customer_name")),
                "customer_address": normalize_text(record.get("customer_address")),
                "customer_original_address_sap": normalize_text(record.get("customer_original_address_sap")),
                "customer_lat": coordinate_text(record.get("customer_lat")),
                "customer_lon": coordinate_text(record.get("customer_lon")),
            },
        )
        warehouse_name_to_ids[normalize_text(record.get("warehouse_name"))].add(warehouse_id)
        customer_name_to_ids[normalize_text(record.get("customer_name"))].add(customer_id)
        duplicate_link_counter[(warehouse_id, customer_id)] += 1
        od_links.append(
            {
                "link_id": f"od_{row_index - 1:06d}",
                "source_row_no": normalize_text(record.get("source_row_no")) or str(row_index - 1),
                "warehouse_id": warehouse_id,
                "customer_id": customer_id,
                "warehouse_name": normalize_text(record.get("warehouse_name")),
                "customer_name": normalize_text(record.get("customer_name")),
                "origin_lat": coordinate_text(record.get("warehouse_lat")),
                "origin_lon": coordinate_text(record.get("warehouse_lon")),
                "destination_lat": coordinate_text(record.get("customer_lat")),
                "destination_lon": coordinate_text(record.get("customer_lon")),
            }
        )

    duplicate_pairs = sum(1 for count in duplicate_link_counter.values() if count > 1)
    duplicate_rows = sum(count - 1 for count in duplicate_link_counter.values() if count > 1)
    warehouses = sorted(warehouses_by_id.values(), key=lambda row: (row["warehouse_name"], row["warehouse_id"]))
    customers = sorted(customers_by_id.values(), key=lambda row: (row["customer_name"], row["customer_id"]))
    summary = {
        "raw_records": len(records),
        "warehouses": len(warehouses),
        "customers": len(customers),
        "od_links": len(od_links),
        "unique_od_pairs": len(duplicate_link_counter),
        "duplicate_od_pairs": duplicate_pairs,
        "duplicate_od_rows": duplicate_rows,
        "invalid_rows": len(invalid_rows),
        "warehouse_names_with_multiple_coordinate_sets": sum(1 for ids in warehouse_name_to_ids.values() if len(ids) > 1),
        "customer_names_with_multiple_coordinate_sets": sum(1 for ids in customer_name_to_ids.values() if len(ids) > 1),
        "invalid_row_samples": invalid_rows[:20],
    }
    return warehouses, customers, od_links, summary


def write_csv(path: Path, rows: list[dict[str, str]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as output_file:
        writer = csv.DictWriter(output_file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def write_report(path: Path, input_path: Path, sheets: list[SheetInfo], summary: dict[str, object]) -> None:
    generated_at = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    sheet_lines = "\n".join(f"| {sheet.name} | {sheet.dimension or 'unknown'} |" for sheet in sheets)
    invalid_samples = summary.get("invalid_row_samples", [])
    sample_lines = "\n".join(
        f"| {row['excel_row']} | {row['source_row_no']} | {row['warehouse_name']} | {row['customer_name']} | {row['errors']} |"
        for row in invalid_samples
    )
    if not sample_lines:
        sample_lines = "| - | - | - | - | - |"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        f"""# Customer Data Preparation Report

Generated at: {generated_at}

## Input Workbook

- File: `{input_path.name}`
- Main sheet used: `{MAIN_SHEET_NAME}`

| Sheet | Dimension |
| --- | --- |
{sheet_lines}

## Output Summary

| Metric | Value |
| --- | ---: |
| Raw data rows | {summary['raw_records']} |
| Unique warehouses | {summary['warehouses']} |
| Unique customers | {summary['customers']} |
| OD link rows | {summary['od_links']} |
| Unique warehouse-customer pairs | {summary['unique_od_pairs']} |
| Duplicate OD pair groups | {summary['duplicate_od_pairs']} |
| Duplicate OD extra rows | {summary['duplicate_od_rows']} |
| Rows with coordinate/name validation warnings | {summary['invalid_rows']} |
| Warehouse names with multiple coordinate/address sets | {summary['warehouse_names_with_multiple_coordinate_sets']} |
| Customer names with multiple coordinate/address sets | {summary['customer_names_with_multiple_coordinate_sets']} |

## Validation Rules

- Warehouse/customer names must be present.
- Warehouse and customer coordinates must parse as finite numbers.
- Coordinates are expected to fall within a broad Korea bounding box: latitude 33.0-39.5, longitude 124.0-132.0.
- Duplicate OD rows are retained in `od_links.csv` because they may represent separate operational rows; the unique pair count is reported for routing de-duplication.

## Invalid Row Samples

| Excel row | Source row no | Warehouse | Customer | Errors |
| ---: | ---: | --- | --- | --- |
{sample_lines}

## Generated Files

- `data/processed/warehouses.csv`
- `data/processed/customers.csv`
- `data/processed/od_links.csv`
- `data/processed/preparation_summary.json`
""",
        encoding="utf-8",
    )


def main() -> int:
    input_path = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else DEFAULT_INPUT
    if not input_path.exists():
        raise FileNotFoundError(f"Input workbook not found: {input_path}")
    records, sheets = read_sheet_rows(input_path, MAIN_SHEET_NAME)
    warehouses, customers, od_links, summary = build_datasets(records)
    write_csv(DEFAULT_OUTPUT_DIR / "warehouses.csv", warehouses, WAREHOUSE_COLUMNS)
    write_csv(DEFAULT_OUTPUT_DIR / "customers.csv", customers, CUSTOMER_COLUMNS)
    write_csv(DEFAULT_OUTPUT_DIR / "od_links.csv", od_links, OD_COLUMNS)
    (DEFAULT_OUTPUT_DIR / "preparation_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    write_report(DEFAULT_REPORT_DIR / "customer_data_preparation_report.md", input_path, sheets, summary)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
