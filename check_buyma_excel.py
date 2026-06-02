#!/usr/bin/env python3
"""Inspect and validate a BUYMA Excel workbook without modifying it.

This script intentionally uses only the Python standard library so it can run in
minimal environments. It reads the .xlsx file as a ZIP archive and reports basic
workbook structure and selected formula-validation results.
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

MYUS_FEE_COUNTRIES = (
    "イギリス",
    "アメリカ",
    "イタリア",
    "フランス",
    "ドイツ",
    "スペイン",
    "オランダ",
    "ベルギー",
    "ポルトガル",
    "オーストリア",
)
CELL_REF_RE = re.compile(
    r"(?<![A-Za-z0-9_])(?P<col>\$?[A-Z]{1,3})(?P<abs_row>\$?)(?P<row>\d+)"
)


@dataclass(frozen=True)
class SheetInfo:
    index: int
    name: str
    xml_path: str
    root: ET.Element


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


def load_sheets(archive: zipfile.ZipFile) -> list[SheetInfo]:
    workbook_root = read_xml(archive, "xl/workbook.xml")
    rels_root = read_xml(archive, "xl/_rels/workbook.xml.rels")

    rel_targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels_root.findall("pkgrel:Relationship", NS)
        if "Id" in rel.attrib and "Target" in rel.attrib
    }

    loaded_sheets: list[SheetInfo] = []
    sheets = workbook_root.findall("main:sheets/main:sheet", NS)
    for index, sheet in enumerate(sheets, start=1):
        name = sheet.attrib.get("name", f"Sheet{index}")
        rel_id = sheet.attrib.get(f"{{{NS['rel']}}}id")
        target = rel_targets.get(rel_id or "")
        if not target:
            continue

        sheet_xml_path = sheet_path_from_target(target)
        loaded_sheets.append(
            SheetInfo(
                index=index,
                name=name,
                xml_path=sheet_xml_path,
                root=read_xml(archive, sheet_xml_path),
            )
        )
    return loaded_sheets


def myus_fee_formula(row: int) -> str:
    country_checks = ",".join(f'J{row}="{country}"' for country in MYUS_FEE_COUNTRIES)
    return (
        f'IF(Q{row}="","",'
        f'IF(OR({country_checks},ISNUMBER(SEARCH("MyUS",K{row}))),'
        f'ROUND(R{row}*0.08,0),0))'
    )


def translate_shared_formula(formula: str, row_delta: int) -> str:
    if row_delta == 0:
        return formula

    def replace(match: re.Match[str]) -> str:
        if match.group("abs_row"):
            return match.group(0)
        row = int(match.group("row")) + row_delta
        return f'{match.group("col")}{row}'

    return CELL_REF_RE.sub(replace, formula)


def formula_for_cell(
    cell: ET.Element, shared_formulas: dict[str, tuple[str, int]]
) -> str | None:
    formula = cell.find("main:f", NS)
    if formula is None:
        return None
    if formula.text:
        return formula.text

    shared_index = formula.attrib.get("si")
    if shared_index not in shared_formulas:
        return None

    master_formula, master_row = shared_formulas[shared_index]
    cell_row = int(re.search(r"\d+", cell.attrib["r"]).group(0))
    return translate_shared_formula(master_formula, cell_row - master_row)


def collect_shared_formulas(sheet_root: ET.Element) -> dict[str, tuple[str, int]]:
    shared_formulas: dict[str, tuple[str, int]] = {}
    for cell in sheet_root.findall(".//main:c", NS):
        formula = cell.find("main:f", NS)
        if formula is None or not formula.text:
            continue
        shared_index = formula.attrib.get("si")
        if not shared_index:
            continue
        row = int(re.search(r"\d+", cell.attrib["r"]).group(0))
        shared_formulas[shared_index] = (formula.text, row)
    return shared_formulas


def validate_production_myus_fee(sheet_root: ET.Element) -> list[str]:
    """Validate 本番管理!S4:S203 MyUS fee formulas.

    The MyUS text check intentionally references column K. Column M is unrelated
    to the current S-column rule and must not be used in the expected formula.
    """
    shared_formulas = collect_shared_formulas(sheet_root)
    problems: list[str] = []
    for row_number in range(4, 204):
        row = sheet_root.find(f"main:sheetData/main:row[@r='{row_number}']", NS)
        cell = row.find(f"main:c[@r='S{row_number}']", NS) if row is not None else None
        actual = formula_for_cell(cell, shared_formulas) if cell is not None else None
        expected = myus_fee_formula(row_number)
        if actual != expected:
            problems.append(
                f"本番管理!S{row_number}: expected={expected} actual={actual or '（式なし）'}"
            )
    return problems


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
        sheets = load_sheets(archive)
        lines.append(f"シート数: {len(sheets)}")

        for sheet in sheets:
            dimension = sheet.root.find("main:dimension", NS)
            dimension_ref = dimension.attrib.get("ref", "不明") if dimension is not None else "不明"
            row_count, cell_count = count_non_empty_rows(sheet.root)
            merge_cells = sheet.root.find("main:mergeCells", NS)
            merge_count = int(merge_cells.attrib.get("count", "0")) if merge_cells is not None else 0

            lines.append(
                f"[{sheet.index}] {sheet.name}: dimension={dimension_ref}, "
                f"non_empty_rows={row_count}, cells={cell_count}, merged_ranges={merge_count}"
            )

        production_sheet = next((sheet for sheet in sheets if sheet.name == "本番管理"), None)
        if production_sheet is None:
            lines.append("検証: 本番管理 シートが見つかりません")
        else:
            problems = validate_production_myus_fee(production_sheet.root)
            lines.append(
                "検証: 本番管理!S4:S203 MyUS手数料円 期待式チェック "
                f"問題 {len(problems)} 件"
            )
            lines.extend(problems[:20])
            if len(problems) > 20:
                lines.append(f"... ほか {len(problems) - 20} 件")

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
