#!/usr/bin/env python3
"""Create the final transfer ZIP from completed newsletter outputs."""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from datetime import datetime
from pathlib import Path


def issue_date_from_output(output_dir: Path) -> str:
    issue_text = output_dir.resolve().parent.name
    if not re.fullmatch(r"20\d{6}", issue_text):
        raise ValueError("output의 상위 폴더명이 YYYYMMDD 형식이 아닙니다: " + issue_text)
    datetime.strptime(issue_text, "%Y%m%d")
    return issue_text


def create_zip(output_dir: Path) -> Path:
    if not output_dir.is_dir():
        raise ValueError("output 폴더를 찾을 수 없습니다: " + str(output_dir))

    issue_text = issue_date_from_output(output_dir)
    images_dir = output_dir / "images"
    image_names = [
        "bg.png",
        "banner_01.png",
        "banner_02.png",
        "banner_03.png",
        "우측배너.png",
        f"newsletter_{issue_text}.png",
    ]
    vdi_html = output_dir / f"vdi_{issue_text}.html"
    mutu_html = output_dir / f"mutu24_{issue_text}.html"
    zip_path = output_dir / f"{issue_text}_반입.zip"

    if zip_path.exists():
        zip_path.unlink()

    required = [images_dir / name for name in image_names] + [vdi_html, mutu_html]
    missing = [str(path) for path in required if not path.is_file()]
    if missing:
        raise ValueError("ZIP 필수 파일이 없습니다: " + ", ".join(missing))

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for image_name in image_names:
            archive.write(images_dir / image_name, f"images/{image_name}")
        archive.write(vdi_html, vdi_html.name)
        archive.write(mutu_html, mutu_html.name)
    return zip_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="완료된 이미지와 HTML을 반입용 ZIP으로 묶습니다.")
    parser.add_argument("--output-dir", required=True, type=Path, help="날짜 작업 폴더 아래의 output 폴더")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        zip_path = create_zip(args.output_dir)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        print("오류: " + str(exc), file=sys.stderr)
        return 1
    print("반입 ZIP: " + str(zip_path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
