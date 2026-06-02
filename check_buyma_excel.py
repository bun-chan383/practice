#!/usr/bin/env python3
"""Inspect a BUYMA Excel workbook without modifying it.

This script intentionally uses only the Python standard library so it can run in
minimal environments. It reads the .xlsx file as a ZIP archive and reports basic
workbook structure and sheet dimensions.
"""

from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Read an .xlsx file and print a non-mutating workbook summary."
    )
    parser.add_argument("workbook", help="Path to the .xlsx workbook to inspect")
    return parser.parse_args()


def read_xml(archive: zipfile.ZipFile, name: str) -> ET.Element:
    with archive.open(name) as fp:
        return ET.parse(fp).getroot()


def sheet_path_from_target(target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return f"xl/{target}"


def count_non_empty_rows(sheet_root: ET.Element) -> tuple[int, int]:
    rows = sheet_root.findall("main:sheetData/main:row", NS)
    non_empty = 0
    cells = 0
    for row in rows:
        row_cells = row.findall("main:c", NS)
        if row_cells:
            non_empty += 1
            cells += len(row_cells)
    return non_empty, cells


def inspect_workbook(path: Path) -> list[str]:
    lines: list[str] = []
    lines.append(f"対象ファイル: {path}")

    if not path.exists():
        raise FileNotFoundError(f"ファイルが存在しません: {path}")
    if not path.is_file():
        raise ValueError(f"通常ファイルではありません: {path}")
    if path.suffix.lower() != ".xlsx":
        raise ValueError(f".xlsx ファイルではありません: {path}")

    lines.append(f"ファイルサイズ: {path.stat().st_size} bytes")

    with zipfile.ZipFile(path) as archive:
        workbook_root = read_xml(archive, "xl/workbook.xml")
        rels_root = read_xml(archive, "xl/_rels/workbook.xml.rels")

        rel_targets = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels_root.findall("pkgrel:Relationship", NS)
            if "Id" in rel.attrib and "Target" in rel.attrib
        }

        sheets = workbook_root.findall("main:sheets/main:sheet", NS)
        lines.append(f"シート数: {len(sheets)}")

        for index, sheet in enumerate(sheets, start=1):
            name = sheet.attrib.get("name", f"Sheet{index}")
            rel_id = sheet.attrib.get(f"{{{NS['rel']}}}id")
            target = rel_targets.get(rel_id or "")
            if not target:
                lines.append(f"[{index}] {name}: シート定義が見つかりません")
                continue

            sheet_xml_path = sheet_path_from_target(target)
            sheet_root = read_xml(archive, sheet_xml_path)
            dimension = sheet_root.find("main:dimension", NS)
            dimension_ref = dimension.attrib.get("ref", "不明") if dimension is not None else "不明"
            row_count, cell_count = count_non_empty_rows(sheet_root)
            merge_cells = sheet_root.find("main:mergeCells", NS)
            merge_count = int(merge_cells.attrib.get("count", "0")) if merge_cells is not None else 0

            lines.append(
                f"[{index}] {name}: dimension={dimension_ref}, "
                f"non_empty_rows={row_count}, cells={cell_count}, merged_ranges={merge_count}"
            )

    lines.append("結果: 読み取りチェック完了（Excelファイルは編集していません）")
    return lines


def main() -> int:
    args = parse_args()
    try:
        for line in inspect_workbook(Path(args.workbook)):
            print(line)
    except (OSError, ValueError, zipfile.BadZipFile, ET.ParseError) as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
