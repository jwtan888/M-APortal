#!/usr/bin/env python3
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import urllib.error
from datetime import datetime, timezone, timedelta
from pathlib import Path


BASE_ID = os.environ.get("AIRTABLE_BASE_ID", "appIgoUvdgdhmf24V")
TABLE_ID = os.environ.get("AIRTABLE_TABLE_ID", "tblFoONq2PzRybmz3")
TOKEN = os.environ.get("AIRTABLE_TOKEN", "").strip()

ROOT = Path(__file__).resolve().parents[1]
ASSET_DIR = ROOT / "construction-assets"
MASTER_PATH = ROOT / "construction-master.js"

FIELDS = [
    "Construction Code",
    "Construction Sketch",
    "Nike Value rating",
    "Complexity Level",
    "Recommended End use",
]


def die(message):
    print(message, file=sys.stderr)
    sys.exit(1)


def request_json(url):
    req = urllib.request.Request(
        url,
        headers={
            "Authorization": f"Bearer {TOKEN}",
            "Content-Type": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        if error.code == 401:
            die("Airtable returned 401 Unauthorized. Update AIRTABLE_TOKEN with a current Airtable PAT.")
        if error.code == 403:
            die("Airtable returned 403 Forbidden. The PAT needs access to this base/table.")
        raise


def fetch_all_records():
    if not TOKEN:
        die("AIRTABLE_TOKEN is required")

    records = []
    offset = None
    while True:
        query = {"pageSize": "100"}
        if offset:
            query["offset"] = offset
        url = f"https://api.airtable.com/v0/{BASE_ID}/{TABLE_ID}?{urllib.parse.urlencode(query)}"
        payload = request_json(url)
        records.extend(payload.get("records") or [])
        offset = payload.get("offset")
        if not offset:
            break
        time.sleep(0.21)
    return records


def text_value(value):
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, list):
        return ", ".join(text_value(item) for item in value if text_value(item)).strip()
    if isinstance(value, dict):
        for key in ("name", "text", "value", "url"):
            if value.get(key):
                return text_value(value[key])
    return str(value).strip()


def sketch_attachment_url(value):
    if isinstance(value, list) and value:
        first = value[0]
        if isinstance(first, dict):
            return first.get("url") or ""
    if isinstance(value, dict):
        return value.get("url") or ""
    if isinstance(value, str) and value.startswith("http"):
        return value
    return ""


def asset_for_code(code):
    matches = sorted(ASSET_DIR.glob(f"{code}_*.png"))
    if matches:
        return f"./construction-assets/{matches[0].name}"
    return ""


def normalize(records):
    rows = []
    seen = set()
    missing_images = []

    for record in records:
        fields = record.get("fields") or {}
        code = text_value(fields.get("Construction Code"))
        if not code or code in seen:
            continue
        seen.add(code)

        sketch = asset_for_code(code)
        if not sketch:
            sketch = sketch_attachment_url(fields.get("Construction Sketch"))
        if not sketch:
            missing_images.append(code)

        rows.append(
            {
                "Construction Code": code,
                "Construction Sketch": sketch,
                "Nike Value rating": text_value(fields.get("Nike Value rating")),
                "Complexity Level": text_value(fields.get("Complexity Level")),
                "Recommended End use": text_value(fields.get("Recommended End use")),
            }
        )

    rows.sort(key=lambda item: item["Construction Code"])
    return rows, missing_images


def write_master(rows):
    updated = datetime.now(timezone(timedelta(hours=8))).isoformat(timespec="seconds")
    payload = {
        "source": f"Airtable construction master {BASE_ID}/{TABLE_ID}",
        "updatedAt": updated,
        "records": rows,
    }
    MASTER_PATH.write_text(
        "window.DFM_CONSTRUCTION_MASTER = "
        + json.dumps(payload, ensure_ascii=True, separators=(",", ":"))
        + ";\n",
        encoding="utf-8",
    )


def main():
    records = fetch_all_records()
    rows, missing_images = normalize(records)
    write_master(rows)
    print(json.dumps({
        "airtableRecords": len(records),
        "masterRows": len(rows),
        "missingImages": missing_images[:25],
        "missingImageCount": len(missing_images),
        "firstCode": rows[0]["Construction Code"] if rows else "",
        "lastCode": rows[-1]["Construction Code"] if rows else "",
        "masterPath": str(MASTER_PATH),
    }, indent=2))


if __name__ == "__main__":
    main()
