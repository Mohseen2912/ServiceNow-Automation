from __future__ import annotations

import argparse
import csv
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any


CHERWELL_TASK_PATTERN = re.compile(
    r"^\s*-\s*task:\s*['\"]?[^#\r\n]*cherwell[^#\r\n]*",
    re.IGNORECASE | re.MULTILINE,
)
SERVICENOW_JOB_PATTERN = re.compile(
    r"^\s*-\s*job:\s*['\"]?\bservicenow\b",
    re.IGNORECASE | re.MULTILINE,
)

DEFAULT_SERVICENOW_FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "authored_by": ("authored_by", "authored by", "authoredby", "author"),
    "service_offering": (
        "service_offering",
        "service offering",
        "serviceoffering",
    ),
    "cmdb_ci": (
        "cmdb_ci",
        "cmdb ci",
        "cmdbci",
        "configuration item",
        "configuration_item",
    ),
    "owned_by": (
        "owned_by",
        "owned by",
        "ownedby",
        "owner by",
        "owner_by",
        "owner",
    ),
}

_SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_RELATIONSHIP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_RELATIONSHIP_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _line_body(line: str) -> str:
    return line.rstrip("\r\n")


def _detect_newline(yaml_text: str) -> str:
    match = re.search(r"\r\n|\n|\r", yaml_text)
    return match.group(0) if match else "\n"


def _indent_width(line: str) -> int:
    return len(line) - len(line.lstrip(" "))


def _normalise_name(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def _yaml_scalar(value: object) -> str:
    """Return a safe, predictable YAML scalar for an Excel cell value."""

    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)

    text = str(value).strip()
    return "'" + text.replace("'", "''") + "'"


def extract_servicenow_fields(
    excel_row: Mapping[str, Any] | None,
    *,
    field_aliases: Mapping[str, Sequence[str]] | None = None,
) -> dict[str, Any]:
    """Map an Excel row's headers to canonical ServiceNow field names.

    Header matching ignores case, spaces, underscores, and punctuation. Empty cells
    are omitted. ``field_aliases`` extends or overrides the built-in mappings.
    """

    if not excel_row:
        return {}

    aliases: dict[str, tuple[str, ...]] = dict(DEFAULT_SERVICENOW_FIELD_ALIASES)
    if field_aliases:
        aliases.update(
            {
                field_name: tuple(column_names)
                for field_name, column_names in field_aliases.items()
            }
        )

    row_by_normalised_header = {
        _normalise_name(header): value for header, value in excel_row.items()
    }
    fields: dict[str, Any] = {}
    for field_name, column_names in aliases.items():
        for column_name in column_names:
            normalised_column = _normalise_name(column_name)
            if normalised_column not in row_by_normalised_header:
                continue

            value = row_by_normalised_header[normalised_column]
            if value is not None and str(value).strip():
                fields[field_name] = value
            break

    return fields


def _render_servicenow_job(
    job_indent: str,
    newline: str,
    servicenow_fields: Mapping[str, Any],
) -> list[str]:
    lines = [f"{job_indent}- job: servicenow{newline}"]
    if servicenow_fields:
        lines.append(f"{job_indent}  variables:{newline}")
        lines.extend(
            f"{job_indent}    {field_name}: {_yaml_scalar(value)}{newline}"
            for field_name, value in servicenow_fields.items()
        )
    lines.extend(
        [
            f"{job_indent}  steps:{newline}",
            f"{job_indent}    - script: echo ServiceNow job{newline}",
        ]
    )
    return lines


def _enrich_existing_servicenow_job(
    block_lines: list[str],
    servicenow_fields: Mapping[str, Any],
    newline: str,
) -> list[str]:
    if not servicenow_fields:
        return block_lines

    job_index = next(
        (
            index
            for index, line in enumerate(block_lines)
            if SERVICENOW_JOB_PATTERN.match(_line_body(line))
        ),
        None,
    )
    if job_index is None:
        return block_lines

    job_body = _line_body(block_lines[job_index])
    job_indent_width = _indent_width(job_body)
    job_indent = job_body[:job_indent_width]
    job_end = len(block_lines)
    for index in range(job_index + 1, len(block_lines)):
        body = _line_body(block_lines[index])
        if not body.strip() or body.lstrip().startswith("#"):
            continue
        if _indent_width(body) <= job_indent_width:
            job_end = index
            break

    variables_index = next(
        (
            index
            for index in range(job_index + 1, job_end)
            if re.match(
                rf"^{re.escape(job_indent)}  variables:\s*$",
                _line_body(block_lines[index]),
            )
        ),
        None,
    )

    variable_lines = [
        f"{job_indent}    {field_name}: {_yaml_scalar(value)}{newline}"
        for field_name, value in servicenow_fields.items()
    ]
    if variables_index is None:
        return (
            block_lines[: job_index + 1]
            + [f"{job_indent}  variables:{newline}"]
            + variable_lines
            + block_lines[job_index + 1 :]
        )

    variables_indent_width = job_indent_width + 2
    variables_end = job_end
    for index in range(variables_index + 1, job_end):
        body = _line_body(block_lines[index])
        if not body.strip() or body.lstrip().startswith("#"):
            continue
        if _indent_width(body) <= variables_indent_width:
            variables_end = index
            break

    canonical_names = {_normalise_name(name) for name in servicenow_fields}
    retained_variables: list[str] = []
    for line in block_lines[variables_index + 1 : variables_end]:
        body = _line_body(line)
        match = re.match(r"^\s+(?P<key>[^:#]+):", body)
        if match and _normalise_name(match.group("key")) in canonical_names:
            continue
        retained_variables.append(line)

    return (
        block_lines[: variables_index + 1]
        + variable_lines
        + retained_variables
        + block_lines[variables_end:]
    )


def insert_servicenow_job(
    yaml_text: str,
    excel_row: Mapping[str, Any] | None = None,
    *,
    field_aliases: Mapping[str, Sequence[str]] | None = None,
) -> str:
    """Insert a ServiceNow job at the top of each stage jobs list when a Cherwell task exists."""

    if not CHERWELL_TASK_PATTERN.search(yaml_text):
        return yaml_text

    newline = _detect_newline(yaml_text)
    servicenow_fields = extract_servicenow_fields(
        excel_row,
        field_aliases=field_aliases,
    )
    lines = yaml_text.splitlines(keepends=True)
    result: list[str] = []
    index = 0

    while index < len(lines):
        line = lines[index]
        result.append(line)

        match = re.match(r"^(?P<indent>[ \t]*)jobs:\s*$", _line_body(line))
        if match:
            jobs_indent = len(match.group("indent"))
            block_start = index + 1
            block_end = block_start

            while block_end < len(lines):
                candidate = _line_body(lines[block_end])
                if not candidate.strip() or candidate.lstrip().startswith("#"):
                    block_end += 1
                    continue

                current_indent = len(candidate) - len(candidate.lstrip(" "))
                if current_indent <= jobs_indent:
                    break
                block_end += 1

            block_lines = lines[block_start:block_end]
            has_cherwell = any(
                CHERWELL_TASK_PATTERN.search(_line_body(item)) for item in block_lines
            )
            has_servicenow = any(
                SERVICENOW_JOB_PATTERN.search(_line_body(item)) for item in block_lines
            )

            if has_cherwell and has_servicenow:
                block_lines = _enrich_existing_servicenow_job(
                    block_lines,
                    servicenow_fields,
                    newline,
                )
            elif has_cherwell:
                insert_at = 0
                job_indent = " " * (jobs_indent + 2)
                for item_index, item in enumerate(block_lines):
                    item_body = _line_body(item)
                    if not item_body.strip() or item_body.lstrip().startswith("#"):
                        continue

                    item_indent = _indent_width(item_body)
                    if item_indent > jobs_indent and item_body.lstrip().startswith("-"):
                        insert_at = item_index
                        job_indent = item_body[:item_indent]
                        break

                block_lines = (
                    block_lines[:insert_at]
                    + _render_servicenow_job(
                        job_indent,
                        newline,
                        servicenow_fields,
                    )
                    + block_lines[insert_at:]
                )

            result.extend(block_lines)
            index = block_end
            continue

        index += 1

    return "".join(result)


def _xlsx_cell_value(
    cell: ET.Element,
    shared_strings: Sequence[str],
) -> str:
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        return "".join(
            node.text or "" for node in cell.findall(f".//{{{_SPREADSHEET_NS}}}t")
        )

    value_node = cell.find(f"{{{_SPREADSHEET_NS}}}v")
    raw_value = "" if value_node is None or value_node.text is None else value_node.text
    if cell_type == "s" and raw_value:
        return shared_strings[int(raw_value)]
    if cell_type == "b":
        return "true" if raw_value == "1" else "false"
    return raw_value


def _read_xlsx_row(
    path: Path,
    *,
    row_number: int,
    header_row: int,
    sheet_name: str | None,
) -> dict[str, str]:
    with zipfile.ZipFile(path) as archive:
        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships_root = ET.fromstring(
            archive.read("xl/_rels/workbook.xml.rels")
        )
        relationship_targets = {
            relationship.get("Id"): relationship.get("Target")
            for relationship in relationships_root.findall(
                f"{{{_PACKAGE_RELATIONSHIP_NS}}}Relationship"
            )
        }
        sheets = workbook_root.findall(f".//{{{_SPREADSHEET_NS}}}sheet")
        selected_sheet = next(
            (
                sheet
                for sheet in sheets
                if sheet_name is not None and sheet.get("name") == sheet_name
            ),
            None,
        )
        if selected_sheet is None and sheet_name is None and sheets:
            selected_sheet = sheets[0]
        if selected_sheet is None:
            raise ValueError(f"Worksheet not found: {sheet_name!r}")

        relationship_id = selected_sheet.get(f"{{{_RELATIONSHIP_NS}}}id")
        target = relationship_targets.get(relationship_id)
        if not target:
            raise ValueError("The selected worksheet has no workbook relationship")
        worksheet_path = posixpath.normpath(posixpath.join("xl", target.lstrip("/")))
        if worksheet_path.startswith("xl/xl/"):
            worksheet_path = worksheet_path[3:]

        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared_strings = [
                "".join(
                    node.text or ""
                    for node in item.findall(f".//{{{_SPREADSHEET_NS}}}t")
                )
                for item in shared_root.findall(f"{{{_SPREADSHEET_NS}}}si")
            ]

        worksheet_root = ET.fromstring(archive.read(worksheet_path))
        requested_rows: dict[int, dict[int, str]] = {}
        for row in worksheet_root.findall(f".//{{{_SPREADSHEET_NS}}}row"):
            current_row = int(row.get("r", "0"))
            if current_row not in {header_row, row_number}:
                continue
            cells: dict[int, str] = {}
            for cell in row.findall(f"{{{_SPREADSHEET_NS}}}c"):
                reference = cell.get("r", "")
                column_letters = re.match(r"[A-Za-z]+", reference)
                if not column_letters:
                    continue
                column_number = 0
                for letter in column_letters.group(0).upper():
                    column_number = column_number * 26 + ord(letter) - ord("A") + 1
                cells[column_number] = _xlsx_cell_value(cell, shared_strings)
            requested_rows[current_row] = cells

    headers = requested_rows.get(header_row, {})
    values = requested_rows.get(row_number, {})
    if not headers:
        raise ValueError(f"Excel header row {header_row} is empty")
    if not values:
        raise ValueError(f"Excel data row {row_number} is empty")
    return {
        header.strip(): values.get(column_number, "")
        for column_number, header in headers.items()
        if header.strip()
    }


def read_excel_row(
    path: Path,
    *,
    row_number: int = 2,
    header_row: int = 1,
    sheet_name: str | None = None,
) -> dict[str, str]:
    """Read one data row from an XLSX or CSV file using its header row as keys."""

    if row_number <= header_row:
        raise ValueError("The data row must come after the header row")

    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return _read_xlsx_row(
            path,
            row_number=row_number,
            header_row=header_row,
            sheet_name=sheet_name,
        )
    if suffix == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as csv_file:
            for _ in range(header_row - 1):
                next(csv_file, None)
            reader = csv.DictReader(csv_file)
            for current_row, row in enumerate(reader, start=header_row + 1):
                if current_row == row_number:
                    return {str(key): value or "" for key, value in row.items() if key}
        raise ValueError(f"CSV data row {row_number} is empty or does not exist")
    raise ValueError("Only .xlsx, .xlsm, and .csv Excel sources are supported")


def update_yaml_file(
    path: Path,
    *,
    excel_row: Mapping[str, Any] | None = None,
    field_aliases: Mapping[str, Sequence[str]] | None = None,
    dry_run: bool = False,
) -> bool:
    yaml_text = path.read_bytes().decode("utf-8")
    updated = insert_servicenow_job(
        yaml_text,
        excel_row,
        field_aliases=field_aliases,
    )

    if updated == yaml_text:
        return False

    if not dry_run:
        path.write_bytes(updated.encode("utf-8"))

    return True


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Insert a ServiceNow job at the top of each jobs block that contains a Cherwell task."
    )
    parser.add_argument("yaml_files", nargs="+", type=Path)
    parser.add_argument(
        "--excel",
        type=Path,
        help="XLSX, XLSM, or CSV file containing the ServiceNow values.",
    )
    parser.add_argument(
        "--sheet",
        help="Worksheet name (defaults to the first sheet).",
    )
    parser.add_argument(
        "--header-row",
        type=int,
        default=1,
        help="One-based Excel header row number (default: 1).",
    )
    parser.add_argument(
        "--row",
        type=int,
        default=2,
        help="One-based Excel data row number (default: 2).",
    )
    parser.add_argument(
        "--field",
        action="append",
        default=[],
        metavar="SERVICENOW_FIELD=EXCEL_COLUMN",
        help="Map an additional ServiceNow field to an Excel column; repeat as needed.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Report files that would change without writing updates.",
    )
    args = parser.parse_args(argv)

    if args.sheet and not args.excel:
        parser.error("--sheet requires --excel")

    custom_field_aliases: dict[str, tuple[str, ...]] = {}
    for field_mapping in args.field:
        field_name, separator, excel_column = field_mapping.partition("=")
        if not separator or not field_name.strip() or not excel_column.strip():
            parser.error("--field must use SERVICENOW_FIELD=EXCEL_COLUMN")
        custom_field_aliases[field_name.strip()] = (excel_column.strip(),)

    try:
        excel_row = (
            read_excel_row(
                args.excel,
                row_number=args.row,
                header_row=args.header_row,
                sheet_name=args.sheet,
            )
            if args.excel
            else None
        )
    except (OSError, ValueError, zipfile.BadZipFile, ET.ParseError) as error:
        parser.error(str(error))

    changed_files = []
    for yaml_file in args.yaml_files:
        changed = update_yaml_file(
            yaml_file,
            excel_row=excel_row,
            field_aliases=custom_field_aliases,
            dry_run=args.dry_run,
        )
        if changed:
            changed_files.append(yaml_file)
            action = "Would update" if args.dry_run else "Updated"
            print(f"{action}: {yaml_file}")
        else:
            print(f"No change: {yaml_file}")

    return 1 if args.dry_run and changed_files else 0


if __name__ == "__main__":
    sys.exit(main())
