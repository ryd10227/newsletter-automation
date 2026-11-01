#!/usr/bin/env python3
"""Export the final newsletter worksheet for Photoshop JSX.

This script intentionally uses only Python's standard library.  An .xlsx file is
a ZIP package, so the values, cached formula results, and embedded pictures can
be read without installing Excel-specific Python packages.
"""

from __future__ import annotations

import argparse
import csv
import os
import posixpath
import re
import sys
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DRAW_MAIN_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

NS = {"m": MAIN_NS, "r": REL_NS, "pr": PKG_REL_NS, "xdr": DRAW_NS, "a": DRAW_MAIN_NS}

SECTION_NAMES = {
    "사업안내형": "business",
    "사업안내": "business",
    "배너형": "banner",
    "배너": "banner",
    "아카데미": "academy",
}
CAPACITY = {"business": 8, "banner": 3, "academy": 4}
WEEKDAYS_KO = ("월", "화", "수", "목", "금", "토", "일")


def clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).replace("\u00a0", " ").replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", text).strip()


def is_input_guide(value: object) -> bool:
    """Return True for the yellow placeholder rows in the blank Excel template."""
    return clean_text(value).startswith("입력 안내:")


def normalize_url(value: object) -> str:
    text = clean_text(value)
    match = re.search(r"https?://[^\s\]\[()<>]+", text, flags=re.IGNORECASE)
    if match:
        text = match.group(0).rstrip(".,;")
    return re.sub(
        r"^(https?://)(kotra\.or\.kr)(?=[:/?#]|$)",
        r"\1www.\2",
        text,
        flags=re.IGNORECASE,
    )


def column_number(cell_ref: str) -> int:
    letters = re.match(r"[A-Z]+", cell_ref)
    if not letters:
        raise ValueError("잘못된 셀 주소: %s" % cell_ref)
    value = 0
    for char in letters.group(0):
        value = value * 26 + ord(char) - 64
    return value


def relationship_map(xml_bytes: bytes) -> Dict[str, str]:
    root = ET.fromstring(xml_bytes)
    return {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in root.findall("pr:Relationship", NS)
    }


def resolve_part(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def rels_part(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def load_shared_strings(package: zipfile.ZipFile) -> List[str]:
    part = "xl/sharedStrings.xml"
    if part not in package.namelist():
        return []
    root = ET.fromstring(package.read(part))
    strings: List[str] = []
    for item in root.findall("m:si", NS):
        strings.append("".join(node.text or "" for node in item.findall(".//m:t", NS)))
    return strings


def target_sheet_part(package: zipfile.ZipFile) -> Tuple[str, str]:
    workbook_part = "xl/workbook.xml"
    workbook = ET.fromstring(package.read(workbook_part))
    sheets = workbook.find("m:sheets", NS)
    if sheets is None or not list(sheets):
        raise ValueError("통합 문서에 시트가 없습니다.")
    rels = relationship_map(package.read(rels_part(workbook_part)))
    sheet_items = list(sheets)
    selected = next(
        (sheet for wanted in ("최종", "최종 업로드") for sheet in sheet_items if sheet.attrib.get("name") == wanted),
        sheet_items[0],
    )
    sheet_name = selected.attrib.get("name", "대상 시트")
    rel_id = selected.attrib["{%s}id" % REL_NS]
    return sheet_name, resolve_part(workbook_part, rels[rel_id])


def read_target_sheet(package: zipfile.ZipFile) -> Tuple[str, str, Dict[int, Dict[int, object]]]:
    shared = load_shared_strings(package)
    sheet_name, sheet_part = target_sheet_part(package)
    root = ET.fromstring(package.read(sheet_part))
    rows: Dict[int, Dict[int, object]] = {}
    for cell in root.findall(".//m:sheetData/m:row/m:c", NS):
        ref = cell.attrib["r"]
        row_match = re.search(r"\d+$", ref)
        if not row_match:
            continue
        row_no = int(row_match.group(0))
        col_no = column_number(ref)
        cell_type = cell.attrib.get("t", "n")
        value_node = cell.find("m:v", NS)
        value: object = None
        if cell_type == "inlineStr":
            value = "".join(node.text or "" for node in cell.findall(".//m:t", NS))
        elif value_node is not None:
            raw = value_node.text or ""
            if cell_type == "s":
                value = shared[int(raw)]
            elif cell_type in ("str", "e"):
                value = raw
            elif cell_type == "b":
                value = raw == "1"
            else:
                try:
                    value = float(raw)
                    if value.is_integer():
                        value = int(value)
                except ValueError:
                    value = raw
        rows.setdefault(row_no, {})[col_no] = value
    return sheet_name, sheet_part, rows


def excel_date(value: object) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
    text = clean_text(value)
    for pattern in (r"^(\d{4})[-./](\d{1,2})[-./](\d{1,2})",):
        match = re.match(pattern, text)
        if match:
            return date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    return None


def format_date(value: date, include_year: bool = False) -> str:
    prefix = ("%02d." % (value.year % 100)) if include_year else ""
    return "%s%d.%d.(%s)" % (prefix, value.month, value.day, WEEKDAYS_KO[value.weekday()])


def format_period(start_value: object, end_value: object, issue_date: Optional[date]) -> str:
    start = excel_date(start_value)
    end = excel_date(end_value)
    if not start:
        return clean_text(start_value)
    show_start_year = bool(issue_date and start.year != issue_date.year)
    start_text = format_date(start, show_start_year)
    if not end or end == start:
        return start_text
    show_end_year = end.year != start.year or bool(issue_date and end.year != issue_date.year)
    return "%s  -  %s" % (start_text, format_date(end, show_end_year))


def normalize_section(value: object) -> Optional[str]:
    compact = clean_text(value).replace(" ", "")
    return SECTION_NAMES.get(compact)


def find_sections(rows: Dict[int, Dict[int, object]]) -> Dict[str, Tuple[int, int]]:
    labels: List[Tuple[int, str]] = []
    for row_no in sorted(rows):
        section = normalize_section(rows[row_no].get(1))
        if section:
            labels.append((row_no, section))
    sections: Dict[str, Tuple[int, int]] = {}
    for index, (label_row, section) in enumerate(labels):
        end_row = labels[index + 1][0] - 1 if index + 1 < len(labels) else max(rows)
        header_row = 0
        for row_no in range(label_row + 1, end_row + 1):
            if clean_text(rows.get(row_no, {}).get(1)) == "연번":
                header_row = row_no
                break
        if not header_row:
            raise ValueError("'%s' 구역에서 '연번' 머리글을 찾지 못했습니다." % section)
        sections[section] = (header_row + 1, end_row)
    missing = [name for name in ("business", "banner", "academy") if name not in sections]
    if missing:
        raise ValueError("'최종' 시트에서 필요한 구역을 찾지 못했습니다: %s" % ", ".join(missing))
    return sections


def numeric_slot(value: object) -> Optional[int]:
    if isinstance(value, (int, float)):
        return int(value)
    match = re.search(r"\d+", clean_text(value))
    return int(match.group(0)) if match else None


def records_from_rows(
    rows: Dict[int, Dict[int, object]],
    issue_date: Optional[date],
    best_effort: bool = False,
    warnings: Optional[List[str]] = None,
) -> List[Dict[str, str]]:
    warnings = warnings if warnings is not None else []
    sections = find_sections(rows)
    output: List[Dict[str, str]] = []

    start_row, end_row = sections["business"]
    business_header = rows.get(start_row - 1, {})
    compact_business = (
        clean_text(business_header.get(3)).replace(" ", "") == "사업기간"
        and clean_text(business_header.get(4)).replace(" ", "") in ("신청마감", "접수마감")
        and clean_text(business_header.get(5)).upper() == "URL"
    )
    for row_no in range(start_row, end_row + 1):
        row = rows.get(row_no, {})
        slot = numeric_slot(row.get(1))
        title = clean_text(row.get(2))
        if slot is None or not title or is_input_guide(title):
            continue
        if not 1 <= slot <= CAPACITY["business"]:
            message = "사업안내 %d번은 PSD 수용 한도(%d)를 넘어 건너뛰었습니다." % (slot, CAPACITY["business"])
            if best_effort:
                warnings.append(message)
                continue
            raise ValueError(message)
        if compact_business:
            period_text = clean_text(row.get(3))
            deadline_text = clean_text(row.get(4))
            url_text = normalize_url(row.get(5))
        else:
            deadline = excel_date(row.get(5))
            period_text = format_period(row.get(3), row.get(4), issue_date)
            deadline_text = (
                format_date(deadline, bool(issue_date and deadline and deadline.year != issue_date.year))
                if deadline
                else clean_text(row.get(9))
            )
            url_text = normalize_url(row.get(6))
        output.append(
            {
                "section": "business",
                "slot": str(slot),
                "title": title,
                "period": period_text,
                "deadline": deadline_text,
                "url": url_text,
                "image_path": "",
            }
        )

    start_row, end_row = sections["academy"]
    academy_header = rows.get(start_row - 1, {})
    compact_academy = (
        clean_text(academy_header.get(3)).replace(" ", "") == "교육기간"
        and clean_text(academy_header.get(4)).replace(" ", "") in ("신청마감", "접수마감", "접수마감일")
        and clean_text(academy_header.get(5)).upper() == "URL"
        and not clean_text(academy_header.get(6))
    )
    for row_no in range(start_row, end_row + 1):
        row = rows.get(row_no, {})
        slot = numeric_slot(row.get(1))
        title = clean_text(row.get(2))
        if slot is None or not title or is_input_guide(title):
            continue
        if not 1 <= slot <= CAPACITY["academy"]:
            message = "아카데미 %d번은 PSD 수용 한도(%d)를 넘어 건너뛰었습니다." % (slot, CAPACITY["academy"])
            if best_effort:
                warnings.append(message)
                continue
            raise ValueError(message)
        if compact_academy:
            period = clean_text(row.get(3))
            deadline = clean_text(row.get(4))
        else:
            period = clean_text(row.get(6)) or clean_text(row.get(3))
            deadline = clean_text(row.get(7)) or clean_text(row.get(4))
        output.append(
            {
                "section": "academy",
                "slot": str(slot),
                "title": title,
                "period": period,
                "deadline": deadline,
                "url": normalize_url(row.get(5)),
                "image_path": "",
            }
        )

    start_row, end_row = sections["banner"]
    for row_no in range(start_row, end_row + 1):
        row = rows.get(row_no, {})
        slot = numeric_slot(row.get(1))
        title = clean_text(row.get(2))
        if slot is None or not title or is_input_guide(title):
            continue
        url = normalize_url(row.get(4))
        # URL이 없는 dimmed 대기행은 Excel 안내용이며 PSD/CSV 데이터가 아닙니다.
        if not url:
            continue
        if not 1 <= slot <= CAPACITY["banner"]:
            message = "배너 %d번은 PSD 수용 한도(%d)를 넘어 건너뛰었습니다." % (slot, CAPACITY["banner"])
            if best_effort:
                warnings.append(message)
                continue
            raise ValueError(message)
        output.append(
            {
                "section": "banner",
                "slot": str(slot),
                "title": title,
                "period": "",
                "deadline": clean_text(row.get(3)),
                "url": url,
                "image_path": "",
            }
        )
    return output


def target_sheet_images(package: zipfile.ZipFile, sheet_part: str) -> List[Tuple[str, bytes]]:
    sheet_rels_name = rels_part(sheet_part)
    if sheet_rels_name not in package.namelist():
        return []
    sheet_rels_root = ET.fromstring(package.read(sheet_rels_name))
    drawing_target = None
    for rel in sheet_rels_root.findall("pr:Relationship", NS):
        if rel.attrib.get("Type", "").endswith("/drawing"):
            drawing_target = rel.attrib["Target"]
            break
    if not drawing_target:
        return []
    drawing_part = resolve_part(sheet_part, drawing_target)
    drawing_rels = relationship_map(package.read(rels_part(drawing_part)))
    root = ET.fromstring(package.read(drawing_part))
    pictures: List[Tuple[int, int, str]] = []
    anchors: Iterable[ET.Element] = list(root.findall("xdr:twoCellAnchor", NS)) + list(
        root.findall("xdr:oneCellAnchor", NS)
    )
    for anchor in anchors:
        from_node = anchor.find("xdr:from", NS)
        blip = anchor.find(".//a:blip", NS)
        if from_node is None or blip is None:
            continue
        rel_id = blip.attrib.get("{%s}embed" % REL_NS)
        if not rel_id or rel_id not in drawing_rels:
            continue
        col = int(from_node.findtext("xdr:col", "0", NS))
        col_offset = int(from_node.findtext("xdr:colOff", "0", NS))
        row = int(from_node.findtext("xdr:row", "0", NS))
        row_offset = int(from_node.findtext("xdr:rowOff", "0", NS))
        x_key = col * 10_000_000 + col_offset
        y_key = row * 10_000_000 + row_offset
        pictures.append((x_key, y_key, resolve_part(drawing_part, drawing_rels[rel_id])))
    result: List[Tuple[str, bytes]] = []
    for _, _, image_part in sorted(pictures):
        result.append((Path(image_part).suffix.lower() or ".png", package.read(image_part)))
    return result


def infer_issue_date(xlsx_path: Path, explicit: Optional[str]) -> date:
    match = re.fullmatch(
        r"newsletter_DB_(20\d{6})\.(?:xlsm|xlsx)",
        xlsx_path.name,
        flags=re.IGNORECASE,
    )
    if not match:
        raise ValueError(
            "Excel 파일명은 newsletter_DB_YYYYMMDD.xlsm 또는 .xlsx 형식이어야 합니다: %s"
            % xlsx_path.name
        )
    filename_date = datetime.strptime(match.group(1), "%Y%m%d").date()
    if explicit:
        explicit_date = datetime.strptime(explicit, "%Y-%m-%d").date()
        if explicit_date != filename_date:
            raise ValueError("Excel 파일명의 날짜와 --newsletter-date가 일치하지 않습니다.")
    return filename_date


def write_export(
    xlsx_path: Path,
    output_dir: Path,
    explicit_date: Optional[str],
    assets_dir: Optional[Path] = None,
    best_effort: bool = False,
    warnings: Optional[List[str]] = None,
) -> Path:
    warnings = warnings if warnings is not None else []
    issue_date = infer_issue_date(xlsx_path, explicit_date)
    output_dir.mkdir(parents=True, exist_ok=True)
    assets_dir = assets_dir or (output_dir / "images")
    assets_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(xlsx_path) as package:
        sheet_name, sheet_part, rows = read_target_sheet(package)
        records = records_from_rows(rows, issue_date, best_effort, warnings)
        images = target_sheet_images(package, sheet_part)

    banner_records = sorted(
        (record for record in records if record["section"] == "banner"),
        key=lambda record: int(record["slot"]),
    )
    for index, record in enumerate(banner_records):
        if index >= len(images):
            break
        suffix, data = images[index]
        if suffix != ".png":
            message = (
                "배너 %s번 이미지는 PNG 형식이어야 합니다. 현재 형식: %s"
                % (record["slot"], suffix or "알 수 없음")
            )
            if best_effort:
                warnings.append(message)
                continue
            raise ValueError(message)
        image_name = "banner_%02d.png" % int(record["slot"])
        image_file = assets_dir / image_name
        image_file.write_bytes(data)
        record["image_path"] = os.path.relpath(image_file, output_dir).replace(os.sep, "/")

    csv_path = output_dir / "newsletter.csv"
    fieldnames = ["section", "slot", "title", "period", "deadline", "url", "image_path", "newsletter_date"]
    with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for index, record in enumerate(records):
            row = dict(record)
            row["newsletter_date"] = issue_date.isoformat() if issue_date and index == 0 else ""
            writer.writerow(row)

    print("시트: %s" % sheet_name)
    print("CSV: %s" % csv_path)
    print(
        "레코드: 사업안내 %d, 아카데미 %d, 배너 %d"
        % (
            sum(r["section"] == "business" for r in records),
            sum(r["section"] == "academy" for r in records),
            sum(r["section"] == "banner" for r in records),
        )
    )
    print("배너 이미지: %d개 추출" % min(len(images), len(banner_records)))
    if len(images) < len(banner_records):
        message = "배너 행보다 포함 이미지가 적습니다. CSV의 image_path를 확인하세요."
        warnings.append(message)
    return csv_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="뉴스레터 Excel의 '최종' 시트를 Photoshop용 CSV로 내보냅니다.")
    parser.add_argument("--xlsx", required=True, type=Path, help="원본 .xlsx 경로")
    parser.add_argument("--newsletter-date", help="발행일(YYYY-MM-DD). 생략 시 Excel 파일명에서 추론")
    parser.add_argument("--output-dir", required=True, type=Path, help="newsletter.csv를 저장할 폴더")
    parser.add_argument("--assets-dir", type=Path, help="배너 이미지를 저장할 폴더. 생략하면 output-dir/images")
    parser.add_argument("--best-effort", action="store_true", help="일부 오류가 있어도 가능한 CSV와 이미지를 저장")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    warnings: List[str] = []
    try:
        write_export(
            args.xlsx,
            args.output_dir,
            args.newsletter_date,
            args.assets_dir,
            args.best_effort,
            warnings,
        )
    except (OSError, ValueError, KeyError, zipfile.BadZipFile) as exc:
        print("오류: %s" % exc, file=sys.stderr)
        return 1
    if warnings:
        for warning in warnings:
            print("경고: %s" % warning, file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
