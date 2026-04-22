#!/usr/bin/env python3

import json
import re
import zipfile
from collections import OrderedDict, defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parent.parent
WORKBOOK_PATH = ROOT / "DFM 2026.xlsx"
OUTPUT_PATH = ROOT / "seed-data.js"

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def cell_value(cell, shared_strings):
    cell_type = cell.attrib.get("t")
    value = cell.find("main:v", NS)
    if cell_type == "s" and value is not None:
        return shared_strings[int(value.text)]
    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.iterfind(".//main:t", NS))
    if value is not None:
        return value.text or ""
    return ""


def read_shared_strings(archive):
    shared_strings_path = "xl/sharedStrings.xml"
    if shared_strings_path not in archive.namelist():
        return []

    strings = []
    root = ET.fromstring(archive.read(shared_strings_path))
    for item in root.findall("main:si", NS):
        text = "".join(node.text or "" for node in item.iterfind(".//main:t", NS))
        strings.append(text)
    return strings


def workbook_sheet_paths(archive):
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {
        rel.attrib["Id"]: "xl/" + rel.attrib["Target"]
        for rel in rels
    }

    sheet_paths = OrderedDict()
    for sheet in workbook.find("main:sheets", NS):
        rel_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        sheet_paths[sheet.attrib["name"]] = rel_map[rel_id]
    return sheet_paths


def read_sheet_rows(archive, shared_strings, sheet_path):
    root = ET.fromstring(archive.read(sheet_path))
    rows = []
    sheet_data = root.find("main:sheetData", NS)
    for row in sheet_data.findall("main:row", NS):
        record = {}
        for cell in row.findall("main:c", NS):
            ref = cell.attrib.get("r", "")
            match = re.match(r"([A-Z]+)", ref)
            if not match:
                continue
            record[match.group(1)] = cell_value(cell, shared_strings)
        rows.append(record)
    return rows


def clean_text(value):
    return " ".join((value or "").replace("\n", " ").split())


def maybe_number(value):
    text = clean_text(value)
    if not text:
        return None
    try:
        number = float(text)
    except ValueError:
        return text
    if number.is_integer():
        return int(number)
    return round(number, 4)


def build_seed_data():
    with zipfile.ZipFile(str(WORKBOOK_PATH)) as archive:
        shared_strings = read_shared_strings(archive)
        sheet_paths = workbook_sheet_paths(archive)

        collection_rows = read_sheet_rows(
            archive,
            shared_strings,
            sheet_paths["Data Collection"],
        )
        defect_rows = read_sheet_rows(
            archive,
            shared_strings,
            sheet_paths["Sheet2"],
        )

    defect_catalog = OrderedDict()
    for row in defect_rows[1:]:
        code = clean_text(row.get("D", ""))
        if not code:
            continue
        entry = defect_catalog.setdefault(
            code,
            {
                "code": code,
                "feature": clean_text(row.get("A", "")),
                "description": clean_text(row.get("B", "")),
                "productClass": clean_text(row.get("C", "")),
                "defects": [],
            },
        )
        defect_name = clean_text(row.get("E", ""))
        intensity = maybe_number(row.get("F", ""))
        if defect_name and defect_name.lower() != "no":
            entry["defects"].append(
                {
                    "name": defect_name,
                    "intensity": intensity if isinstance(intensity, (int, float)) else None,
                }
            )

    style_fg_values = defaultdict(list)
    records = []
    for index, row in enumerate(collection_rows[1:], start=2):
        season = clean_text(row.get("B", ""))
        style = clean_text(row.get("E", ""))
        style_key = "{}__{}".format(season, style)

        fg_qty = maybe_number(row.get("K", ""))
        fg_number = fg_qty if isinstance(fg_qty, (int, float)) else None
        if season and style and fg_number is not None:
            style_fg_values[style_key].append(fg_number)

        type_code = clean_text(row.get("H", "")) or clean_text(row.get("G", ""))
        defect_info = defect_catalog.get(type_code, {"defects": []})
        total_intensity = 0
        for defect in defect_info.get("defects", []):
            total_intensity += defect.get("intensity") or 0

        records.append(
            {
                "id": "row-{}".format(index),
                "sourceRow": index,
                "season": season,
                "category": clean_text(row.get("C", "")),
                "protoStage": clean_text(row.get("D", "")),
                "style": style,
                "styleKey": style_key,
                "constructionCode": clean_text(row.get("G", "")),
                "typeCode": type_code,
                "modification": clean_text(row.get("I", "")),
                "remark": clean_text(row.get("J", "")),
                "fgQty": fg_number,
                "fgAnchor": maybe_number(row.get("L", "")),
                "feature": defect_info.get("feature", ""),
                "description": defect_info.get("description", ""),
                "productClass": defect_info.get("productClass", ""),
                "defects": defect_info.get("defects", []),
                "defectCount": len(defect_info.get("defects", [])),
                "totalIntensity": round(total_intensity, 4),
            }
        )

    warnings = []
    for style_key, values in sorted(style_fg_values.items()):
        unique_values = list(OrderedDict.fromkeys(values))
        if len(unique_values) > 1:
            season, style = style_key.split("__", 1)
            warnings.append(
                {
                    "type": "fg_conflict",
                    "styleKey": style_key,
                    "season": season,
                    "style": style,
                    "values": unique_values,
                    "message": "Conflicting FG Qty values detected for {} / {}.".format(season, style),
                }
            )

    payload = {
        "meta": {
            "sourceFile": WORKBOOK_PATH.name,
            "recordCount": len(records),
            "styleSeasonCount": len(style_fg_values),
            "generatedBy": "scripts/build_seed_data.py",
        },
        "records": records,
        "defectCatalog": list(defect_catalog.values()),
        "warnings": warnings,
    }

    js = "window.DFM_SEED_DATA = {};\n".format(
        json.dumps(payload, ensure_ascii=True, separators=(",", ":"))
    )
    OUTPUT_PATH.write_text(js, encoding="utf-8")


if __name__ == "__main__":
    build_seed_data()
