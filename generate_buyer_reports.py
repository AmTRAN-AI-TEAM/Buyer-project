from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def numeric(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def clean_number(value: float):
    if abs(value - round(value)) < 1e-9:
        return int(round(value))
    return value


def excel_serial_date(value) -> str | None:
    if isinstance(value, dt.datetime):
        value = value.date()
    if isinstance(value, dt.date):
        return str((value - dt.date(1899, 12, 30)).days)
    if isinstance(value, (int, float)) and 40000 <= int(value) <= 50000:
        return str(int(value))
    if isinstance(value, str) and value.strip().isdigit():
        number = int(value.strip())
        if 40000 <= number <= 50000:
            return str(number)
    return None


def date_from_excel_serial(serial: str) -> dt.date:
    return dt.date(1899, 12, 30) + dt.timedelta(days=int(serial))


def find_header_row(ws, required: Iterable[str]) -> int:
    required_set = set(required)
    for row in ws.iter_rows():
        values = {str(cell.value).strip() for cell in row if cell.value is not None}
        if required_set.issubset(values):
            return row[0].row
    raise ValueError(f"Cannot find header row containing: {', '.join(required)}")


def column_map(ws, header_row: int) -> dict[str, int]:
    result = {}
    for cell in ws[header_row]:
        if cell.value is not None:
            result[str(cell.value).strip()] = cell.column
    return result


def autosize(ws) -> None:
    for col_idx, column_cells in enumerate(ws.columns, start=1):
        width = 10
        for cell in column_cells:
            if cell.value is not None:
                width = max(width, min(45, len(str(cell.value)) + 2))
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def style_sheet(ws) -> None:
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        if isinstance(cell.value, (dt.date, dt.datetime)):
            cell.number_format = "yyyy/m/d"
    ws.freeze_panes = "B2"
    autosize(ws)


def write_matrix_workbook(
    output_path: Path,
    sheet_name: str,
    headers: list[str],
    rows: list[list],
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(headers)
    for row in rows:
        ws.append(row)
    style_sheet(ws)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def generate_dps(dps_path: Path, output_path: Path, include_star_parts: bool = True) -> None:
    wb = load_workbook(dps_path, data_only=True, read_only=True)
    if "DPS原始" not in wb.sheetnames:
        raise ValueError("DPS file does not contain sheet: DPS原始")
    ws = wb["DPS原始"]

    header_row = find_header_row(ws, ["Line", "W/O", "AVTC P/N"])
    headers = [cell.value for cell in ws[header_row]]
    cols = column_map(ws, header_row)
    pn_col = cols["AVTC P/N"]

    date_columns: list[tuple[int, str]] = []
    for idx, value in enumerate(headers, start=1):
        date_label = excel_serial_date(value)
        if date_label is not None:
            date_columns.append((idx, date_label))

    if not date_columns:
        raise ValueError("No Excel serial date columns found in DPS原始")

    aggregate: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        pn = str(row[pn_col - 1]).strip() if row[pn_col - 1] is not None else ""
        if not pn:
            continue
        if not include_star_parts and pn.endswith("*"):
            continue
        for col_idx, date_label in date_columns:
            qty = numeric(row[col_idx - 1] if col_idx - 1 < len(row) else None)
            if qty:
                aggregate[pn][date_label] += qty

    dates = sorted({date for _col_idx, date in date_columns}, key=int)
    output_rows = []
    for pn in sorted(aggregate):
        values = [clean_number(aggregate[pn].get(date, 0.0)) for date in dates]
        total = clean_number(sum(numeric(value) for value in values))
        output_rows.append([pn, *values, total])

    date_headers = [date_from_excel_serial(date) for date in dates]
    write_matrix_workbook(output_path, "DPS整理后_auto", ["行标签", *date_headers, "total"], output_rows)


def read_shared_strings(zip_file: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []
    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    return [
        "".join(text.text or "" for text in item.findall(f".//{{{SPREADSHEET_NS}}}t"))
        for item in root.findall(f"{{{SPREADSHEET_NS}}}si")
    ]


def parse_pivot_cache(pp_path: Path) -> tuple[list[str], list[list[str]]]:
    with zipfile.ZipFile(pp_path) as xlsx:
        definition = ET.fromstring(xlsx.read("xl/pivotCache/pivotCacheDefinition1.xml"))
        fields: list[str] = []
        shared_by_field: list[list[str]] = []

        for cache_field in definition.find(f"{{{SPREADSHEET_NS}}}cacheFields"):
            fields.append(cache_field.attrib.get("name", ""))
            shared_items = []
            shared_root = cache_field.find(f"{{{SPREADSHEET_NS}}}sharedItems")
            if shared_root is not None:
                for item in shared_root:
                    tag = item.tag.split("}", 1)[-1]
                    shared_items.append("" if tag == "m" else item.attrib.get("v", ""))
            shared_by_field.append(shared_items)

        records_root = ET.fromstring(xlsx.read("xl/pivotCache/pivotCacheRecords1.xml"))
        records: list[list[str]] = []
        for record in records_root.findall(f"{{{SPREADSHEET_NS}}}r"):
            row = []
            for index, child in enumerate(record):
                tag = child.tag.split("}", 1)[-1]
                if tag == "x":
                    shared_index = int(child.attrib.get("v", "0"))
                    values = shared_by_field[index]
                    row.append(values[shared_index] if shared_index < len(values) else "")
                else:
                    row.append("" if tag == "m" else child.attrib.get("v", ""))
            records.append(row)
        return fields, records


def pp_output_mapping(fields: list[str]) -> list[tuple[str, list[int]]]:
    week_fields: dict[int, list[int]] = defaultdict(list)
    forecast_fields: list[tuple[str, int]] = []

    for idx, field in enumerate(fields):
        normalized = field.replace("_x000a_", " ").strip()
        week_match = re.match(r"WK(\d{2})\s+(\d{2})'([A-Za-z]+)$", normalized)
        if week_match:
            week = int(week_match.group(1))
            if 30 <= week <= 44:
                week_fields[week].append(idx)
            continue

        forecast_match = re.match(r"([A-Za-z]{3})'(\d{2})FCST$", normalized)
        if forecast_match:
            month, year = forecast_match.groups()
            if (year == "26" and month in {"Nov", "Dec"}) or year == "27":
                forecast_fields.append((f"{month}{year}FCST", idx))

    mapping: list[tuple[str, list[int]]] = []
    for week in sorted(week_fields):
        mapping.append((f"WK{week:02d}", sorted(week_fields[week])))

    month_order = {
        "Nov26FCST": 0,
        "Dec26FCST": 1,
        "Jan27FCST": 2,
        "Feb27FCST": 3,
        "Mar27FCST": 4,
        "Apr27FCST": 5,
        "May27FCST": 6,
        "Jun27FCST": 7,
        "Jul27FCST": 8,
    }
    for label, idx in sorted(forecast_fields, key=lambda item: month_order.get(item[0], 999)):
        if label.endswith("27FCST"):
            label = label[:3] + "'27 FCST"
        mapping.append((label, [idx]))
    return mapping


def generate_pp(pp_path: Path, output_path: Path) -> None:
    fields, records = parse_pivot_cache(pp_path)
    field_index = {name.strip(): idx for idx, name in enumerate(fields)}
    required = ["Customer", "Model", "AVTC FG Part Number", "Plan"]
    missing = [name for name in required if name not in field_index]
    if missing:
        raise ValueError(f"Missing pivot cache fields: {', '.join(missing)}")

    customer_idx = field_index["Customer"]
    model_idx = field_index["Model"]
    pn_idx = field_index["AVTC FG Part Number"]
    plan_idx = field_index["Plan"]
    mapping = pp_output_mapping(fields)
    if not mapping:
        raise ValueError("No PP output period fields were detected")

    aggregate: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    metadata: dict[str, tuple[str, str]] = {}

    for record in records:
        if len(record) <= max(customer_idx, model_idx, pn_idx, plan_idx):
            continue
        if record[plan_idx].strip() != "Production Input":
            continue
        pn = record[pn_idx].strip()
        if not pn:
            continue
        metadata.setdefault(pn, (record[customer_idx].strip(), record[model_idx].strip()))
        for output_label, source_indexes in mapping:
            qty = sum(numeric(record[i] if i < len(record) else None) for i in source_indexes)
            if qty:
                aggregate[pn][output_label] += qty

    periods = [label for label, _indexes in mapping]
    output_rows = []
    for pn in sorted(aggregate, key=lambda key: (metadata.get(key, ("", ""))[0], key)):
        customer, model = metadata.get(pn, ("", ""))
        values = [clean_number(aggregate[pn].get(period, 0.0)) for period in periods]
        total = clean_number(sum(numeric(value) for value in values))
        output_rows.append([customer, pn, model, *values, total])

    write_matrix_workbook(
        output_path,
        "整理后PP_auto",
        ["Customer", "AVTC FG Part Number", "Model", *periods, "total"],
        output_rows,
    )


def default_buyer_folder() -> Path:
    return Path(os.environ["USERPROFILE"]) / "Desktop" / "瑞軒資料" / "buyer"


def main() -> None:
    buyer_folder = default_buyer_folder()
    parser = argparse.ArgumentParser(description="Generate DPS and PP organized Excel files from raw source data.")
    parser.add_argument("--dps", type=Path, default=buyer_folder / "TV DPS Jul 31-Ver 2.xlsx")
    parser.add_argument("--pp", type=Path, default=buyer_folder / "AVTC TV MNT VC PP 20260730 update (002).xlsx")
    parser.add_argument("--out-dir", type=Path, default=buyer_folder)
    parser.add_argument("--exclude-star-parts", action="store_true", help="Exclude DPS part numbers ending with *.")
    args = parser.parse_args()

    dps_output = args.out_dir / "DPS整理后_auto.xlsx"
    pp_output = args.out_dir / "整理后PP_auto.xlsx"

    generate_dps(args.dps, dps_output, include_star_parts=not args.exclude_star_parts)
    generate_pp(args.pp, pp_output)

    print(f"DPS output: {dps_output}")
    print(f"PP output: {pp_output}")


if __name__ == "__main__":
    main()
