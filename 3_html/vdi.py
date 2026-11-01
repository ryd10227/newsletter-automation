#!/usr/bin/env python3
"""Fill managed newsletter image-map links from the Photoshop CSV export."""

from __future__ import annotations

import argparse
import csv
import html
import re
import sys
from pathlib import Path
from typing import Dict, List, Match, Optional, Tuple


UNSUBSCRIBE_URL = "https://www.kotra.or.kr/membership/mberEmailRecive.api?type="
SECTION_ALT_PREFIX = {
    "business": "사업안내",
    "academy": "아카데미",
    "banner": "홍보배너",
}
CAPACITY = {"business": 8, "academy": 4, "banner": 3}
LOCKED_SLOTS = {("banner", 2)}

AREA_RE = re.compile(r"<area\b[^>]*>", re.IGNORECASE)
ALT_RE = re.compile(r"\balt\s*=\s*([\"'])(.*?)\1", re.IGNORECASE)
HREF_RE = re.compile(r"\s+href\s*=\s*([\"']).*?\1", re.IGNORECASE)
MAIN_IMG_RE = re.compile(
    r"<img\b(?=[^>]*\busemap\s*=\s*([\"'])#Map\1)[^>]*>",
    re.IGNORECASE,
)
SRC_RE = re.compile(r"\bsrc\s*=\s*([\"']).*?\1", re.IGNORECASE)


def read_csv_links(
    csv_path: Path,
    best_effort: bool = False,
    warnings: Optional[List[str]] = None,
) -> Dict[Tuple[str, int], str]:
    warnings = warnings if warnings is not None else []
    links: Dict[Tuple[str, int], str] = {}
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        required = {"section", "slot", "url"}
        if not reader.fieldnames or not required.issubset(reader.fieldnames):
            raise ValueError("CSV에 section, slot, url 열이 필요합니다.")

        for line_number, row in enumerate(reader, start=2):
            section = (row.get("section") or "").strip().lower()
            if section not in SECTION_ALT_PREFIX:
                continue
            try:
                slot = int((row.get("slot") or "").strip())
            except ValueError:
                message = f"CSV {line_number}행의 slot이 숫자가 아니라 건너뛰었습니다."
                if best_effort:
                    warnings.append(message)
                    continue
                raise ValueError(message)
            if slot < 1 or slot > CAPACITY[section]:
                message = (
                    f"CSV {line_number}행의 {section} {slot}번은 HTML 수용 한도를 넘습니다."
                )
                if best_effort:
                    warnings.append(message + " 이 행을 건너뛰었습니다.")
                    continue
                raise ValueError(message)
            raw_url = (row.get("url") or "").strip()
            url = raw_url.splitlines()[0].strip() if raw_url else ""
            if not re.match(r"^https?://", url, re.IGNORECASE):
                message = f"CSV {line_number}행의 URL이 올바르지 않아 건너뛰었습니다: {url}"
                if best_effort:
                    warnings.append(message)
                    continue
                raise ValueError(message)
            key = (section, slot)
            if key in links:
                message = f"CSV에 {section} {slot}번이 중복되어 뒤의 행을 건너뛰었습니다."
                if best_effort:
                    warnings.append(message)
                    continue
                raise ValueError(message)
            links[key] = url
    return links


def area_alt(area_tag: str) -> str:
    match = ALT_RE.search(area_tag)
    return html.unescape(match.group(2)).strip() if match else ""


def managed_key(alt: str) -> Optional[Tuple[str, int]]:
    for section, prefix in SECTION_ALT_PREFIX.items():
        match = re.fullmatch(re.escape(prefix) + r"(\d+)", alt)
        if match:
            slot = int(match.group(1))
            if 1 <= slot <= CAPACITY[section]:
                return section, slot
    return None


def replace_href(area_tag: str, url: Optional[str]) -> str:
    if url:
        attribute = f' href="{html.escape(url, quote=True)}"'
        if HREF_RE.search(area_tag):
            return HREF_RE.sub(attribute, area_tag, count=1)
        return area_tag[:-1] + attribute + ">"
    return HREF_RE.sub("", area_tag, count=1)


def replace_main_image_src(html_text: str, image_name: str) -> Tuple[str, int]:
    def update_image(match: Match[str]) -> str:
        image_tag = match.group(0)
        attribute = f'src="{image_name}"'
        if SRC_RE.search(image_tag):
            return SRC_RE.sub(attribute, image_tag, count=1)
        return image_tag[:-1] + " " + attribute + ">"

    return MAIN_IMG_RE.subn(update_image, html_text)


def newsletter_image_name(html_path: Path) -> str:
    match = re.fullmatch(r"newsletter_(20\d{6})\.html", html_path.name, re.IGNORECASE)
    if not match:
        raise ValueError(
            "HTML 파일명은 newsletter_YYYYMMDD.html 형식이어야 합니다: "
            + html_path.name
        )
    return f"images/newsletter_{match.group(1)}.png"


def fill_html(
    html_text: str,
    links: Dict[Tuple[str, int], str],
    require_unsubscribe: bool = True,
) -> Tuple[str, Dict[str, int]]:
    counts = {"business": 0, "academy": 0, "banner": 0, "preserved": 0, "unsubscribe": 0}

    def update_area(match: Match[str]) -> str:
        area_tag = match.group(0)
        alt = area_alt(area_tag)
        if alt == "수신거부":
            counts["unsubscribe"] += 1
            return replace_href(area_tag, UNSUBSCRIBE_URL)

        key = managed_key(alt)
        if key is None:
            return area_tag

        # 홍보배너 2는 HTML에 들어 있는 고정 링크를 항상 사용합니다.
        # CSV에 값이 들어오더라도 이 슬롯은 절대 덮어쓰지 않습니다.
        if key in LOCKED_SLOTS:
            counts["preserved"] += 1
            return area_tag

        section, _ = key
        url = links.get(key)
        if url:
            counts[section] += 1
            return replace_href(area_tag, url)

        # CSV에 없는 슬롯은 기존 링크를 보존합니다. 배너 2처럼 이번
        # 회차의 데이터 대상이 아닌 고정 슬롯을 실수로 지우지 않습니다.
        counts["preserved"] += 1
        return area_tag

    output = AREA_RE.sub(update_area, html_text)
    if require_unsubscribe and counts["unsubscribe"] != 1:
        raise ValueError("HTML에서 alt=\"수신거부\" 영역을 정확히 1개 찾지 못했습니다.")
    return output, counts


def hrefs_outside_managed(html_text: str) -> Dict[str, str]:
    result: Dict[str, str] = {}
    for index, match in enumerate(AREA_RE.finditer(html_text), start=1):
        tag = match.group(0)
        alt = area_alt(tag)
        if managed_key(alt) is not None or alt == "수신거부":
            continue
        href = HREF_RE.search(tag)
        result[f"{index}:{alt}"] = href.group(0) if href else ""
    return result


def default_output_path(html_path: Path) -> Path:
    match = re.fullmatch(r"newsletter_(20\d{6})\.html", html_path.name, re.IGNORECASE)
    if not match:
        raise ValueError(
            "HTML 파일명은 newsletter_YYYYMMDD.html 형식이어야 합니다: "
            + html_path.name
        )
    return html_path.with_name(f"vdi_{match.group(1)}.html")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="뉴스레터 CSV URL을 HTML 이미지맵의 사업안내·아카데미·배너 영역에 넣습니다."
    )
    parser.add_argument("--csv", required=True, type=Path, help="1_excel_to_csv/export.py가 만든 CSV")
    parser.add_argument("--html", required=True, type=Path, help="원본 뉴스레터 HTML")
    parser.add_argument("--output", type=Path, help="결과 HTML. 생략하면 vdi_YYYYMMDD.html")
    parser.add_argument("--best-effort", action="store_true", help="잘못된 CSV 행을 건너뛰고 가능한 링크를 저장")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    output_path = args.output or default_output_path(args.html)
    warnings: List[str] = []
    try:
        # newline="" keeps the source file's CRLF/LF convention unchanged so
        # the generated HTML differs only at the managed href attributes.
        with args.html.open("r", encoding="utf-8-sig", newline="") as handle:
            source = handle.read()
        protected_before = hrefs_outside_managed(source)
        if args.best_effort:
            try:
                links = read_csv_links(args.csv, True, warnings)
            except (OSError, ValueError) as exc:
                links = {}
                warnings.append(f"CSV를 읽지 못해 기존 링크를 유지했습니다: {exc}")
        else:
            links = read_csv_links(args.csv)
        output, counts = fill_html(source, links, require_unsubscribe=not args.best_effort)
        image_name = newsletter_image_name(args.html)
        output, image_count = replace_main_image_src(output, image_name)
        if image_count != 1:
            message = 'HTML에서 usemap="#Map"을 사용하는 메인 이미지를 정확히 1개 찾지 못했습니다.'
            if args.best_effort:
                warnings.append(message)
            else:
                raise ValueError(message)
        if args.best_effort and counts["unsubscribe"] != 1:
            warnings.append('HTML에서 alt="수신거부" 영역을 정확히 1개 찾지 못해 해당 URL은 변경하지 못했습니다.')
        protected_after = hrefs_outside_managed(output)
        if protected_before != protected_after:
            message = "관리 대상 밖의 링크가 변경되었습니다."
            if args.best_effort:
                warnings.append(message)
            else:
                raise ValueError(message + " 저장을 중단했습니다.")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
            handle.write(output)
    except (OSError, ValueError) as exc:
        print(f"오류: {exc}", file=sys.stderr)
        return 1

    print(f"결과 HTML: {output_path}")
    print(
        "링크 적용: 사업안내 {business}, 아카데미 {academy}, 배너 {banner}; "
        "기존 링크 유지 {preserved}; 수신거부 변경 {unsubscribe}".format(**counts)
    )
    if warnings:
        for warning in warnings:
            print(f"경고: {warning}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
