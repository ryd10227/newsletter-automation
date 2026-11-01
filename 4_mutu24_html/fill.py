#!/usr/bin/env python3
"""Build the Mootu24 newsletter HTML from PSD text, CSV data, and image-map URLs."""

from __future__ import annotations

import argparse
import csv
import html
import re
import struct
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Dict, Iterable, List, Match, Optional, Sequence, Tuple


TAG_RE = re.compile(r"<[^>]+>", re.DOTALL)
AREA_RE = re.compile(r"<area\b[^>]*>", re.IGNORECASE | re.DOTALL)
ATTR_RE_TEMPLATE = r"\b{name}\s*=\s*([\"'])(.*?)\1"
IMG_RE = re.compile(r"<img\b[^>]*>", re.IGNORECASE | re.DOTALL)
ANCHOR_RE = re.compile(r"<a\b[^>]*>.*?</a\s*>", re.IGNORECASE | re.DOTALL)
SECTION_CAPACITY = {"business": 8, "academy": 4, "banner": 3}
# PSD/무투24 목록은 왼쪽 열 3개 뒤 오른쪽 열 3개 순서지만,
# 뉴스레터 이미지맵의 리스트 번호는 위에서부터 좌우 교차 순서입니다.
MARKET_LINK_SLOTS = (1, 3, 5, 2, 4, 6)


@dataclass(frozen=True)
class Layer:
    index: int
    top: int
    left: int
    bottom: int
    right: int
    name: str
    is_text: bool


@dataclass(frozen=True)
class MarketItem:
    country: str
    office: str
    title: str
    top: int
    left: int


def clean_text(value: object) -> str:
    text = str(value or "").replace("\u00a0", " ").replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", text).strip()


def attr_value(tag: str, name: str) -> str:
    match = re.search(ATTR_RE_TEMPLATE.format(name=re.escape(name)), tag, re.IGNORECASE | re.DOTALL)
    return html.unescape(match.group(2)).strip() if match else ""


def set_attr(tag: str, name: str, value: str) -> str:
    escaped = html.escape(value, quote=True)
    pattern = re.compile(ATTR_RE_TEMPLATE.format(name=re.escape(name)), re.IGNORECASE | re.DOTALL)
    replacement = f'{name}="{escaped}"'
    if pattern.search(tag):
        return pattern.sub(replacement, tag, count=1)
    return tag[:-1].rstrip() + " " + replacement + ">"


def read_csv_records(
    csv_path: Path,
    best_effort: bool = False,
    warnings: Optional[List[str]] = None,
) -> Tuple[Dict[Tuple[str, int], Dict[str, str]], Optional[date]]:
    warnings = warnings if warnings is not None else []
    records: Dict[Tuple[str, int], Dict[str, str]] = {}
    issue_date: Optional[date] = None
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        required = {"section", "slot", "title", "period", "deadline", "url"}
        if not reader.fieldnames or not required.issubset(reader.fieldnames):
            raise ValueError("CSV에 section, slot, title, period, deadline, url 열이 필요합니다.")
        for line_no, row in enumerate(reader, start=2):
            date_text = clean_text(row.get("newsletter_date"))
            if date_text and issue_date is None:
                try:
                    issue_date = date.fromisoformat(date_text)
                except ValueError:
                    message = f"CSV {line_no}행의 newsletter_date 형식이 올바르지 않습니다: {date_text}"
                    if best_effort:
                        warnings.append(message)
                    else:
                        raise ValueError(message)
            section = clean_text(row.get("section")).lower()
            if section not in SECTION_CAPACITY:
                continue
            try:
                slot = int(clean_text(row.get("slot")))
            except ValueError:
                message = f"CSV {line_no}행의 slot이 숫자가 아니라 건너뛰었습니다."
                if best_effort:
                    warnings.append(message)
                    continue
                raise ValueError(message)
            if not 1 <= slot <= SECTION_CAPACITY[section]:
                message = f"CSV {line_no}행의 {section} {slot}번은 수용 한도를 넘어 건너뛰었습니다."
                if best_effort:
                    warnings.append(message)
                    continue
                raise ValueError(message)
            key = (section, slot)
            if key in records:
                message = f"CSV에 {section} {slot}번이 중복되어 뒤의 행을 건너뛰었습니다."
                if best_effort:
                    warnings.append(message)
                    continue
                raise ValueError(message)
            record = {name: clean_text(row.get(name)) for name in reader.fieldnames}
            url = record.get("url", "")
            if url and not re.match(r"^https?://", url, re.IGNORECASE):
                message = f"CSV {line_no}행의 URL이 올바르지 않아 빈 링크로 처리했습니다: {url}"
                if best_effort:
                    warnings.append(message)
                    record["url"] = ""
                else:
                    raise ValueError(message)
            records[key] = record
    return records, issue_date


def read_map_links(html_path: Path) -> Dict[str, str]:
    with html_path.open("r", encoding="utf-8-sig", newline="") as handle:
        source = handle.read()
    links: Dict[str, str] = {}
    for match in AREA_RE.finditer(source):
        tag = match.group(0)
        alt = attr_value(tag, "alt")
        href = attr_value(tag, "href")
        if alt and href:
            links[alt] = href
    return links


def primary_section_order(html_path: Path) -> str:
    """Return the visual order encoded by the input newsletter image-map HTML."""
    with html_path.open("r", encoding="utf-8-sig", newline="") as handle:
        source = handle.read()

    # Prefer the y-coordinate because it describes the actual image layout.
    # Fall back to source order for templates whose area tags have no coords.
    market_key: Optional[Tuple[int, int]] = None
    business_key: Optional[Tuple[int, int]] = None
    for match in AREA_RE.finditer(source):
        tag = match.group(0)
        alt = attr_value(tag, "alt")
        coords = attr_value(tag, "coords")
        numbers = [int(value) for value in re.findall(r"-?\d+", coords)]
        top = numbers[1] if len(numbers) >= 2 else match.start()
        key = (top, match.start())
        if market_key is None and alt in ("해외시장뉴스_메인", "헤드라인", "리스트1"):
            market_key = key
        if business_key is None and (alt == "사업안내_메인" or alt == "사업안내1"):
            business_key = key
        if market_key is not None and business_key is not None:
            break
    if market_key is None or business_key is None:
        raise ValueError("이미지맵 HTML에서 해외시장뉴스와 사업안내의 배치 순서를 찾지 못했습니다.")
    return "business_first" if business_key < market_key else "market_first"


def _u16(data: bytes, offset: int) -> int:
    return struct.unpack_from(">H", data, offset)[0]


def _i16(data: bytes, offset: int) -> int:
    return struct.unpack_from(">h", data, offset)[0]


def _u32(data: bytes, offset: int) -> int:
    return struct.unpack_from(">I", data, offset)[0]


def read_psd_layers(psd_path: Path) -> List[Layer]:
    """Read layer names, bounds, and text-layer flags from a regular PSD file."""
    data = psd_path.read_bytes()
    if len(data) < 30 or data[:4] != b"8BPS":
        raise ValueError("올바른 PSD 파일이 아닙니다.")
    version = _u16(data, 4)
    if version != 1:
        raise ValueError("PSB는 지원하지 않습니다. PSD 형식으로 저장해 주세요.")

    pos = 26
    for _ in range(2):  # color-mode data, image resources
        length = _u32(data, pos)
        pos += 4 + length
    layer_mask_length = _u32(data, pos)
    pos += 4
    if layer_mask_length == 0:
        raise ValueError("PSD에 레이어 정보가 없습니다.")
    layer_info_length = _u32(data, pos)
    pos += 4
    if layer_info_length < 2:
        raise ValueError("PSD에 레이어 레코드가 없습니다.")
    layer_count = abs(_i16(data, pos))
    pos += 2

    layers: List[Layer] = []
    for index in range(layer_count):
        top, left, bottom, right = struct.unpack_from(">iiii", data, pos)
        pos += 16
        channel_count = _u16(data, pos)
        pos += 2 + channel_count * 6
        pos += 12  # blend signature/mode plus opacity, clipping, flags, filler
        extra_length = _u32(data, pos)
        pos += 4
        extra_end = pos + extra_length

        mask_length = _u32(data, pos)
        pos += 4 + mask_length
        ranges_length = _u32(data, pos)
        pos += 4 + ranges_length

        pascal_length = data[pos]
        pos += 1
        pascal_name = data[pos : pos + pascal_length].decode("macroman", "replace")
        pos += pascal_length
        pos += (4 - ((1 + pascal_length) % 4)) % 4

        unicode_name: Optional[str] = None
        is_text = False
        while pos + 12 <= extra_end:
            signature = data[pos : pos + 4]
            key = data[pos + 4 : pos + 8]
            block_length = _u32(data, pos + 8)
            pos += 12
            if signature not in (b"8BIM", b"8B64") or pos + block_length > extra_end:
                break
            if key == b"luni" and block_length >= 4:
                char_count = _u32(data, pos)
                raw = data[pos + 4 : pos + 4 + char_count * 2]
                unicode_name = raw.decode("utf-16be", "replace")
            elif key == b"TySh":
                is_text = True
            pos += block_length
            if pos % 2:
                pos += 1
        pos = extra_end
        layers.append(
            Layer(index, top, left, bottom, right, clean_text(unicode_name or pascal_name), is_text)
        )
    return layers


def _text_layers(layers: Iterable[Layer]) -> List[Layer]:
    return [layer for layer in layers if layer.is_text and clean_text(layer.name)]


def extract_psd_content(layers: Sequence[Layer]) -> Tuple[str, List[MarketItem], List[Tuple[str, str]], str, Optional[date]]:
    texts = _text_layers(layers)
    market_heading = next((x for x in texts if "해외시장뉴스" in x.name), None)
    business_heading = next((x for x in texts if x.name == "KOTRA 사업안내"), None)
    if market_heading is None or business_heading is None:
        raise ValueError("PSD에서 해외시장뉴스 또는 사업안내 제목 레이어를 찾지 못했습니다.")

    # The Tuesday layout puts market news first, while the Friday layout puts
    # business notices first. Bound the market section by the next heading
    # below it instead of assuming that the business heading is always below.
    section_names = {
        "KOTRA 사업안내",
        "아카데미 교육·연수 과정",
        "경제외교 소식",
    }
    following_headings = [
        x.top for x in texts if x.name in section_names and x.top > market_heading.top
    ]
    market_bottom = min(following_headings) if following_headings else max(x.bottom for x in texts) + 1
    market_zone = [x for x in texts if market_heading.top < x.top < market_bottom]
    locations = [x for x in market_zone if "무역관" in x.name and ("ㅣ" in x.name or "|" in x.name)]
    if len(locations) != 6:
        raise ValueError(f"PSD에서 해외시장뉴스 무역관 레이어를 6개 대신 {len(locations)}개 찾았습니다.")

    market_items: List[MarketItem] = []
    used_titles = set()
    for location in locations:
        candidates = [
            x
            for x in market_zone
            if x.index not in used_titles
            and x.index != location.index
            and "무역관" not in x.name
            and "해외시장뉴스" not in x.name
            and clean_text(x.name).strip("…·. ")
            and -5 <= x.top - location.top <= 80
            and abs(x.left - location.left) <= 180
        ]
        if not candidates:
            raise ValueError(f"'{location.name}'에 대응하는 기사 제목 레이어를 찾지 못했습니다.")
        title_layer = min(candidates, key=lambda x: (abs(x.top - location.top), abs(x.left - location.left)))
        used_titles.add(title_layer.index)
        parts = re.split(r"[ㅣ|]", location.name, maxsplit=1)
        if len(parts) != 2:
            raise ValueError(f"국가/무역관 형식을 해석할 수 없습니다: {location.name}")
        market_items.append(
            MarketItem(clean_text(parts[0]), clean_text(parts[1]), clean_text(title_layer.name), location.top, location.left)
        )

    # Mootu24's list is column-major: left column 1-3, then right column 4-6.
    # Text alignment makes x coordinates differ by a few pixels, so classify
    # the columns first instead of sorting on the raw x coordinate.
    left_edge = min(item.left for item in market_items)
    right_edge = max(item.left for item in market_items)
    column_split = (left_edge + right_edge) / 2
    market_items.sort(key=lambda item: (0 if item.left < column_split else 1, item.top))

    first_list_top = min(item.top for item in market_items)
    headline_candidates = [
        x
        for x in market_zone
        if x.top < first_list_top
        and "해외시장뉴스" not in x.name
        and "무역관" not in x.name
        and clean_text(x.name).strip("…·. ")
    ]
    if not headline_candidates:
        raise ValueError("PSD에서 해외시장뉴스 헤드라인 제목을 찾지 못했습니다.")
    headline = max(headline_candidates, key=lambda x: len(x.name)).name

    economy_heading = next((x for x in texts if x.name == "경제외교 소식"), None)
    economy_items: List[Tuple[str, str]] = []
    if economy_heading is not None:
        economy_layers = [
            x for x in texts if x.top > economy_heading.top and x.left < 560 and ("ㅣ" in x.name or "|" in x.name)
        ]
        economy_layers.sort(key=lambda x: x.top)
        for layer in economy_layers[:2]:
            parts = re.split(r"[ㅣ|]", layer.name, maxsplit=1)
            economy_items.append((clean_text(parts[0]), clean_text(parts[1])))
    if len(economy_items) != 2:
        raise ValueError(f"PSD에서 경제외교 소식을 2개 대신 {len(economy_items)}개 찾았습니다.")

    economy_banner = next((x.name for x in texts if "Help Desk" in x.name), "")
    psd_date: Optional[date] = None
    for layer in texts:
        match = re.fullmatch(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일", layer.name)
        if match:
            psd_date = date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
            break
    return headline, market_items, economy_items, economy_banner, psd_date


def find_container(html_text: str, class_name: str, start: int = 0) -> Tuple[int, int, int, int]:
    open_pattern = re.compile(
        r'<div\b(?=[^>]*\bclass\s*=\s*(["\'])[^"\']*\b'
        + re.escape(class_name)
        + r'\b[^"\']*\1)[^>]*>',
        re.IGNORECASE | re.DOTALL,
    )
    match = open_pattern.search(html_text, start)
    if not match:
        raise ValueError(f'HTML에서 class="{class_name}" 영역을 찾지 못했습니다.')
    depth = 1
    tag_pattern = re.compile(r"<div\b[^>]*>|</div\s*>", re.IGNORECASE | re.DOTALL)
    for tag in tag_pattern.finditer(html_text, match.end()):
        if tag.group(0).lower().startswith("</div"):
            depth -= 1
            if depth == 0:
                return match.start(), match.end(), tag.start(), tag.end()
        else:
            depth += 1
    raise ValueError(f'class="{class_name}" 영역의 닫는 </div>를 찾지 못했습니다.')


def replace_container_inner(html_text: str, class_name: str, inner: str, start: int = 0) -> str:
    _, open_end, close_start, _ = find_container(html_text, class_name, start)
    return html_text[:open_end] + inner + html_text[close_start:]


def reorder_primary_sections(html_text: str, order: str) -> str:
    headings = {
        "market": "<h2>KOTRA 해외시장뉴스</h2>",
        "business": "<h2>KOTRA 사업안내</h2>",
    }
    spans: Dict[str, Tuple[int, int]] = {}
    start = 0
    while True:
        try:
            block_start, _, _, block_end = find_container(html_text, "nlCont", start)
        except ValueError:
            break
        block = html_text[block_start:block_end]
        for name, heading in headings.items():
            if heading in block:
                spans[name] = (block_start, block_end)
        start = block_end
    if set(spans) != {"market", "business"}:
        raise ValueError("무투24 템플릿에서 해외시장뉴스와 사업안내 영역을 찾지 못했습니다.")

    market_start, market_end = spans["market"]
    business_start, business_end = spans["business"]
    is_business_first = business_start < market_start
    wants_business_first = order == "business_first"
    if is_business_first == wants_business_first:
        return html_text

    if market_start < business_start:
        first_start, first_end = market_start, market_end
        second_start, second_end = business_start, business_end
    else:
        first_start, first_end = business_start, business_end
        second_start, second_end = market_start, market_end
    between = html_text[first_end:second_start]
    if re.search(r'<div\b[^>]*\bclass\s*=\s*(["\'])[^"\']*\bnlCont\b', between, re.IGNORECASE):
        raise ValueError("무투24 템플릿에서 해외시장뉴스와 사업안내 사이에 다른 영역이 있어 순서를 바꿀 수 없습니다.")
    return (
        html_text[:first_start]
        + html_text[second_start:second_end]
        + between
        + html_text[first_start:first_end]
        + html_text[second_end:]
    )


def replace_project_section(html_text: str, heading: str, records: Sequence[Dict[str, str]], period_label: str) -> str:
    heading_pos = html_text.find(f"<h2>{heading}</h2>")
    if heading_pos < 0:
        raise ValueError(f"HTML에서 '{heading}' 제목을 찾지 못했습니다.")
    boxes = []
    for record in records:
        boxes.append(
            '\n                <a href="{url}" class="projectBox">\n'
            '                    <ul>\n'
            '                        <li class="period">\n'
            '                            <span>{period_label}</span>\n'
            '                            <p>{period}</p>\n'
            '                        </li>\n'
            '                        <li class="deadline">\n'
            '                            <span>신청마감</span>\n'
            '                            <p>{deadline}</p>\n'
            '                        </li>\n'
            '                    </ul>\n'
            '                    <p>{title}</p>\n'
            '                </a>'.format(
                url=html.escape(record.get("url", ""), quote=True),
                period_label=period_label,
                period=html.escape(record.get("period", "")),
                deadline=html.escape(record.get("deadline", "")),
                title=html.escape(record.get("title", "")),
            )
        )
    inner = "".join(boxes) + "\n            "
    return replace_container_inner(html_text, "projectBoxes", inner, heading_pos)


def replace_named_anchor(html_text: str, class_name: str, href: str, image_alt: str) -> str:
    pattern = re.compile(
        r'<a\b(?=[^>]*\bclass\s*=\s*(["\'])[^"\']*\b'
        + re.escape(class_name)
        + r'\b[^"\']*\1)[^>]*>.*?</a\s*>',
        re.IGNORECASE | re.DOTALL,
    )
    match = pattern.search(html_text)
    if not match:
        raise ValueError(f'HTML에서 class="{class_name}" 링크를 찾지 못했습니다.')
    anchor = match.group(0)
    open_end = anchor.find(">") + 1
    open_tag = set_attr(anchor[:open_end], "href", href)
    anchor = open_tag + anchor[open_end:]
    image_match = IMG_RE.search(anchor)
    if image_match:
        image_tag = set_attr(image_match.group(0), "alt", image_alt)
        anchor = anchor[: image_match.start()] + image_tag + anchor[image_match.end() :]
    return html_text[: match.start()] + anchor + html_text[match.end() :]


def replace_banners(
    html_text: str,
    records: Dict[Tuple[str, int], Dict[str, str]],
    map_links: Dict[str, str],
) -> str:
    _, open_end, close_start, _ = find_container(html_text, "bannerCont")
    inner = html_text[open_end:close_start]
    anchors = list(ANCHOR_RE.finditer(inner))
    if len(anchors) != 3:
        raise ValueError(f"HTML 배너 영역에서 링크를 3개 대신 {len(anchors)}개 찾았습니다.")
    replacements: List[str] = []
    for slot, match in enumerate(anchors, start=1):
        anchor = match.group(0)
        record = records.get(("banner", slot))
        if slot == 2:
            href = map_links.get("홍보배너2", attr_value(anchor, "href"))
            alt = "글로벌 방위산업 트렌드"
        elif record:
            href = record.get("url", "")
            alt = record.get("title", "")
        else:
            href = attr_value(anchor, "href")
            image_match = IMG_RE.search(anchor)
            alt = attr_value(image_match.group(0), "alt") if image_match else ""
        open_tag_end = anchor.find(">") + 1
        anchor = set_attr(anchor[:open_tag_end], "href", href) + anchor[open_tag_end:]
        image_match = IMG_RE.search(anchor)
        if image_match:
            image_tag = set_attr(image_match.group(0), "alt", alt)
            anchor = anchor[: image_match.start()] + image_tag + anchor[image_match.end() :]
        replacements.append(anchor)
    for match, replacement in reversed(list(zip(anchors, replacements))):
        inner = inner[: match.start()] + replacement + inner[match.end() :]
    return html_text[:open_end] + inner + html_text[close_start:]


def update_header_date(html_text: str, issue_date: date) -> str:
    pattern = re.compile(
        r"(<div\s+class=\"lt\">\s*<h1>kotra 뉴스레터</h1>\s*<p>).*?(</p>)",
        re.IGNORECASE | re.DOTALL,
    )
    replacement = rf"\g<1>{issue_date.year}년 {issue_date.month}월 {issue_date.day}일\g<2>"
    output, count = pattern.subn(replacement, html_text, count=1)
    if count != 1:
        raise ValueError("HTML에서 뉴스레터 발행일을 찾지 못했습니다.")
    return output


def image_sources(html_text: str) -> List[str]:
    return [attr_value(match.group(0), "src") for match in IMG_RE.finditer(html_text)]


def build_html(
    template: str,
    records: Dict[Tuple[str, int], Dict[str, str]],
    issue_date: date,
    map_links: Dict[str, str],
    headline: str,
    market_items: Sequence[MarketItem],
    economy_items: Sequence[Tuple[str, str]],
    economy_banner: str,
    section_order: str,
) -> str:
    required_links = ["헤드라인"] + [f"리스트{i}" for i in range(1, 7)] + [
        "홍보배너2",
        "경제외교1",
        "경제외교2",
        "경제외교 우측배너",
    ]
    missing = [name for name in required_links if not map_links.get(name)]
    if missing:
        raise ValueError("이미지맵 HTML에서 링크를 찾지 못했습니다: " + ", ".join(missing))

    output = update_header_date(template, issue_date)
    output = replace_named_anchor(output, "titleNews", map_links["헤드라인"], headline)

    news_boxes = []
    for slot, item in enumerate(market_items, start=1):
        map_slot = MARKET_LINK_SLOTS[slot - 1]
        news_boxes.append(
            '\n                <a href="{url}" class="newsBox" target="_blank">\n'
            '                    <div class="local">\n'
            '                        <span>{country}</span>\n'
            '                        <span>{office}</span>\n'
            '                    </div>\n'
            '                    <p>{title}</p>\n'
            '                </a>'.format(
                url=html.escape(map_links[f"리스트{map_slot}"], quote=True),
                country=html.escape(item.country),
                office=html.escape(item.office),
                title=html.escape(item.title),
            )
        )
    output = replace_container_inner(output, "newsBoxes", "".join(news_boxes) + "\n            ")

    business = [records[("business", slot)] for slot in range(1, 9) if ("business", slot) in records]
    academy = [records[("academy", slot)] for slot in range(1, 5) if ("academy", slot) in records]
    output = replace_project_section(output, "KOTRA 사업안내", business, "사업기간")
    output = replace_project_section(output, "아카데미 교육·연수 과정", academy, "교육기간")
    output = replace_banners(output, records, map_links)

    example_boxes = []
    for slot, (label, title) in enumerate(economy_items, start=1):
        example_boxes.append(
            '\n                    <a href="{url}" class="exampleBox">\n'
            '                        <span>{label}</span>\n'
            '                        <p>{title}</p>\n'
            '                    </a>'.format(
                url=html.escape(map_links[f"경제외교{slot}"], quote=True),
                label=html.escape(label),
                title=html.escape(title),
            )
        )
    output = replace_container_inner(output, "exampleBoxes", "".join(example_boxes) + "\n                ")
    output = replace_named_anchor(
        output,
        "exampleImg",
        map_links["경제외교 우측배너"],
        economy_banner or "경제외교 소식",
    )
    output = reorder_primary_sections(output, section_order)
    return output


def template_section_order(html_text: str) -> str:
    market = html_text.find("<h2>KOTRA 해외시장뉴스</h2>")
    business = html_text.find("<h2>KOTRA 사업안내</h2>")
    if market < 0 or business < 0:
        return "market_first"
    return "business_first" if business < market else "market_first"


def build_html_best_effort(
    template: str,
    records: Dict[Tuple[str, int], Dict[str, str]],
    issue_date: Optional[date],
    map_links: Dict[str, str],
    headline: str,
    market_items: Sequence[MarketItem],
    economy_items: Sequence[Tuple[str, str]],
    economy_banner: str,
    section_order: str,
    warnings: List[str],
) -> str:
    """Fill independent sections and preserve the template wherever one section fails."""
    output = template

    def apply(label: str, operation) -> None:
        nonlocal output
        try:
            output = operation(output)
        except (KeyError, ValueError) as exc:
            warnings.append(f"{label}: {exc}")

    if issue_date is not None:
        apply("발행일", lambda value: update_header_date(value, issue_date))
    else:
        warnings.append("발행일을 찾지 못해 템플릿 값을 유지했습니다.")

    if headline and map_links.get("헤드라인"):
        apply(
            "해외시장뉴스 헤드라인",
            lambda value: replace_named_anchor(value, "titleNews", map_links["헤드라인"], headline),
        )
    else:
        warnings.append("해외시장뉴스 헤드라인 텍스트 또는 링크가 없어 템플릿 값을 유지했습니다.")

    market_link_names = [f"리스트{slot}" for slot in range(1, 7)]
    if len(market_items) == 6 and all(map_links.get(name) for name in market_link_names):
        news_boxes = []
        for slot, item in enumerate(market_items, start=1):
            map_slot = MARKET_LINK_SLOTS[slot - 1]
            news_boxes.append(
                '\n                <a href="{url}" class="newsBox" target="_blank">\n'
                '                    <div class="local">\n'
                '                        <span>{country}</span>\n'
                '                        <span>{office}</span>\n'
                '                    </div>\n'
                '                    <p>{title}</p>\n'
                '                </a>'.format(
                    url=html.escape(map_links[f"리스트{map_slot}"], quote=True),
                    country=html.escape(item.country),
                    office=html.escape(item.office),
                    title=html.escape(item.title),
                )
            )
        apply(
            "해외시장뉴스 6건",
            lambda value: replace_container_inner(value, "newsBoxes", "".join(news_boxes) + "\n            "),
        )
    else:
        warnings.append("해외시장뉴스 6건의 PSD 텍스트 또는 링크가 부족해 템플릿 값을 유지했습니다.")

    business = [records[("business", slot)] for slot in range(1, 9) if ("business", slot) in records]
    academy = [records[("academy", slot)] for slot in range(1, 5) if ("academy", slot) in records]
    if business:
        apply(
            "사업안내",
            lambda value: replace_project_section(value, "KOTRA 사업안내", business, "사업기간"),
        )
    else:
        warnings.append("사업안내 CSV 데이터가 없어 템플릿 값을 유지했습니다.")
    if academy:
        apply(
            "아카데미",
            lambda value: replace_project_section(value, "아카데미 교육·연수 과정", academy, "교육기간"),
        )
    else:
        warnings.append("아카데미 CSV 데이터가 없어 템플릿 값을 유지했습니다.")

    if any(key[0] == "banner" for key in records) or map_links.get("홍보배너2"):
        apply("홍보배너", lambda value: replace_banners(value, records, map_links))
    else:
        warnings.append("홍보배너 데이터와 링크가 없어 템플릿 값을 유지했습니다.")

    if len(economy_items) == 2 and all(map_links.get(f"경제외교{slot}") for slot in (1, 2)):
        example_boxes = []
        for slot, (label, title) in enumerate(economy_items, start=1):
            example_boxes.append(
                '\n                    <a href="{url}" class="exampleBox">\n'
                '                        <span>{label}</span>\n'
                '                        <p>{title}</p>\n'
                '                    </a>'.format(
                    url=html.escape(map_links[f"경제외교{slot}"], quote=True),
                    label=html.escape(label),
                    title=html.escape(title),
                )
            )
        apply(
            "경제외교 소식",
            lambda value: replace_container_inner(value, "exampleBoxes", "".join(example_boxes) + "\n                "),
        )
    else:
        warnings.append("경제외교 소식 텍스트 또는 링크가 부족해 템플릿 값을 유지했습니다.")

    if economy_banner and map_links.get("경제외교 우측배너"):
        apply(
            "경제외교 우측배너",
            lambda value: replace_named_anchor(
                value,
                "exampleImg",
                map_links["경제외교 우측배너"],
                economy_banner,
            ),
        )
    else:
        warnings.append("경제외교 우측배너 텍스트 또는 링크가 없어 템플릿 값을 유지했습니다.")

    apply("해외시장뉴스/사업안내 순서", lambda value: reorder_primary_sections(value, section_order))
    return output


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="PSD·CSV·이미지맵 HTML의 데이터를 무투24 HTML 템플릿에 자동 입력합니다."
    )
    parser.add_argument("--template", required=True, type=Path, help="무투24 HTML 템플릿(.txt 또는 .html)")
    parser.add_argument("--csv", required=True, type=Path, help="1_excel_to_csv/export.py가 만든 newsletter.csv")
    parser.add_argument("--psd", required=True, type=Path, help="텍스트가 확정된 뉴스레터 PSD")
    parser.add_argument("--map-html", required=True, type=Path, help="링크가 들어 있는 이미지맵 뉴스레터 HTML")
    parser.add_argument("--output", required=True, type=Path, help="생성할 무투24 HTML")
    parser.add_argument("--issue-date", help="발행일 YYYY-MM-DD. CSV/PSD보다 우선")
    parser.add_argument("--best-effort", action="store_true", help="일부 입력 오류가 있어도 가능한 영역을 저장")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    warnings: List[str] = []
    try:
        template_bytes = args.template.read_bytes()
        template_has_bom = template_bytes.startswith(b"\xef\xbb\xbf")
        template = template_bytes.decode("utf-8-sig")
        if args.best_effort:
            try:
                records, csv_date = read_csv_records(args.csv, True, warnings)
            except (OSError, ValueError) as exc:
                records, csv_date = {}, None
                warnings.append(f"CSV 읽기: {exc}")
            try:
                links = read_map_links(args.map_html)
            except OSError as exc:
                links = {}
                warnings.append(f"이미지맵 HTML 읽기: {exc}")
            try:
                section_order = primary_section_order(args.map_html)
            except (OSError, ValueError) as exc:
                section_order = template_section_order(template)
                warnings.append(f"영역 순서 판별: {exc}; 템플릿 순서를 유지했습니다.")
            try:
                layers = read_psd_layers(args.psd)
                headline, market_items, economy_items, economy_banner, psd_date = extract_psd_content(layers)
            except (OSError, ValueError, struct.error) as exc:
                headline, market_items, economy_items, economy_banner, psd_date = "", [], [], "", None
                warnings.append(f"PSD 텍스트 추출: {exc}")
            explicit_date: Optional[date] = None
            if args.issue_date:
                try:
                    explicit_date = date.fromisoformat(args.issue_date)
                except ValueError:
                    warnings.append(f"명령행 발행일 형식이 올바르지 않습니다: {args.issue_date}")
            issue_date = explicit_date or csv_date or psd_date
            output = build_html_best_effort(
                template,
                records,
                issue_date,
                links,
                headline,
                market_items,
                economy_items,
                economy_banner,
                section_order,
                warnings,
            )
        else:
            records, csv_date = read_csv_records(args.csv)
            links = read_map_links(args.map_html)
            section_order = primary_section_order(args.map_html)
            layers = read_psd_layers(args.psd)
            headline, market_items, economy_items, economy_banner, psd_date = extract_psd_content(layers)
            explicit_date = date.fromisoformat(args.issue_date) if args.issue_date else None
            issue_date = explicit_date or csv_date or psd_date
            if issue_date is None:
                raise ValueError("CSV 또는 PSD에서 뉴스레터 발행일을 찾지 못했습니다.")
            output = build_html(
                template,
                records,
                issue_date,
                links,
                headline,
                market_items,
                economy_items,
                economy_banner,
                section_order,
            )
        newline = "\r\n" if "\r\n" in template else "\n"
        output = output.replace("\r\n", "\n").replace("\r", "\n")
        if newline == "\r\n":
            output = output.replace("\n", "\r\n")
        if image_sources(template) != image_sources(output):
            raise ValueError("이미지 src가 변경되어 저장을 중단했습니다.")
        args.output.parent.mkdir(parents=True, exist_ok=True)
        output_bytes = output.encode("utf-8")
        if template_has_bom:
            output_bytes = b"\xef\xbb\xbf" + output_bytes
        args.output.write_bytes(output_bytes)
    except (OSError, ValueError, struct.error) as exc:
        print(f"오류: {exc}", file=sys.stderr)
        return 1

    print(f"결과 HTML: {args.output}")
    print(
        f"입력 완료: 헤드라인 1, 해외시장뉴스 {len(market_items)}, "
        f"사업안내 {sum(1 for key in records if key[0] == 'business')}, "
        f"아카데미 {sum(1 for key in records if key[0] == 'academy')}, "
        f"배너 3, 경제외교 {len(economy_items)}"
    )
    print("이미지 src: 변경 없음")
    print("상단 순서: " + ("사업안내 → 해외시장뉴스" if section_order == "business_first" else "해외시장뉴스 → 사업안내"))
    if warnings:
        for warning in warnings:
            print(f"경고: {warning}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
