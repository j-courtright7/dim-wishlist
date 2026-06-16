#!/usr/bin/env python3
"""
Build DIM wishlist files directly from the current Aegis Endgame Analysis Excel workbook.

What it does:
- Reads data/endgame.xlsx by default.
- Finds rows with Name, Tier, Rank, Perk 1, Perk 2, and Notes-style columns.
- Downloads the current Destiny 2 manifest from Bungie.
- Maps weapon/perk names to Destiny Inventory Item hashes.
- Writes strict DIM wishlist lines:
    dimwishlist:item=<weapon_hash>&perks=<perk_1_hash>,<perk_2_hash>#notes:...
- Outputs S-tier, A-tier, and A+S-tier files.

This intentionally does NOT include barrels, magazines, or origin traits in the DIM match.
That keeps the wishlist useful for vault cleaning without requiring perfect 5/5 rolls.
"""

from __future__ import annotations

import argparse
import csv
import itertools
import json
import os
import re
import sqlite3
import sys
import tempfile
import urllib.request
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

try:
    import openpyxl
except ImportError as exc:
    raise SystemExit(
        "Missing dependency: openpyxl. Install with: pip install openpyxl"
    ) from exc


BUNGIE_ROOT = "https://www.bungie.net"


@dataclass(frozen=True)
class AegisRow:
    sheet: str
    row_number: int
    name: str
    tier: str
    rank: str
    perk1_names: tuple[str, ...]
    perk2_names: tuple[str, ...]
    barrel: str = ""
    mag: str = ""
    origin_trait: str = ""
    notes: str = ""


@dataclass
class ManifestIndex:
    # Normalized display name -> inventory item hashes
    weapons_by_name: dict[str, list[int]]
    plugs_by_name: dict[str, list[int]]
    all_item_names_by_hash: dict[int, str]


def norm_name(value: Any) -> str:
    """Normalize names for matching Aegis sheet text to Bungie manifest display names."""
    if value is None:
        return ""
    text = str(value).strip()
    text = text.replace("\u2019", "'").replace("\u2018", "'")
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = re.sub(r"\s+", " ", text)
    return text.casefold()


def clean_cell(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = text.replace("\u2019", "'").replace("\u2018", "'")
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def split_perks(value: Any) -> tuple[str, ...]:
    """
    Split perk cells like:
      Reconstruction, Demolitionist, Discord
      Fourth Time's the Charm\nKinetic Tremors
    """
    text = clean_cell(value)
    if not text:
        return tuple()

    # Remove common annotations.
    text = re.sub(r"\([^)]*enhanced[^)]*\)", "", text, flags=re.I)
    text = re.sub(r"\bEnhanced\b", "", text, flags=re.I)

    parts: list[str] = []
    for chunk in re.split(r"[\n\r,;]+", text):
        chunk = clean_cell(chunk)
        if not chunk:
            continue
        # Avoid obvious non-perk placeholders.
        if chunk.casefold() in {"none", "n/a", "na", "-", "?", "need testing"}:
            continue
        parts.append(chunk)

    # Deduplicate while preserving order.
    seen: set[str] = set()
    out: list[str] = []
    for p in parts:
        key = norm_name(p)
        if key not in seen:
            seen.add(key)
            out.append(p)
    return tuple(out)


def find_header_row(ws: Any, max_scan_rows: int = 20) -> tuple[int, dict[str, int]] | None:
    """
    Locate a header row containing Name, Perk 1, Perk 2, Tier.
    Returns (row_number, normalized_header -> 1-based column index).
    """
    max_row = ws.max_row or 0
    max_col = ws.max_column or 0

    if max_row < 1 or max_col < 1:
        return None

    for row in range(1, min(max_scan_rows, max_row) + 1):
        headers: dict[str, int] = {}

        for col in range(1, max_col + 1):
            text = clean_cell(ws.cell(row=row, column=col).value)
            if text:
                headers[norm_name(text)] = col

        if "name" in headers and "perk 1" in headers and "perk 2" in headers and "tier" in headers:
            return row, headers

    return None


def pick_col(headers: dict[str, int], *names: str) -> int | None:
    for name in names:
        key = norm_name(name)
        if key in headers:
            return headers[key]
    return None


def read_aegis_rows(excel_path: Path) -> list[AegisRow]:
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    rows: list[AegisRow] = []

    for ws in wb.worksheets:
        header = find_header_row(ws)
        if not header:
            continue

        header_row, headers = header
        name_col = pick_col(headers, "Name")
        tier_col = pick_col(headers, "Tier")
        rank_col = pick_col(headers, "Rank")
        perk1_col = pick_col(headers, "Perk 1")
        perk2_col = pick_col(headers, "Perk 2")
        barrel_col = pick_col(headers, "Barrel")
        mag_col = pick_col(headers, "Mag")
        origin_col = pick_col(headers, "Origin Trait", "Origin Trai")
        notes_col = pick_col(headers, "Notes")

        if not all([name_col, tier_col, perk1_col, perk2_col]):
            continue

        for r in range(header_row + 1, ws.max_row + 1):
            name = clean_cell(ws.cell(row=r, column=name_col).value)
            tier = clean_cell(ws.cell(row=r, column=tier_col).value).upper()
            if not name or tier not in {"S", "A"}:
                continue

            perk1 = split_perks(ws.cell(row=r, column=perk1_col).value)
            perk2 = split_perks(ws.cell(row=r, column=perk2_col).value)
            if not perk1 or not perk2:
                continue

            rank = clean_cell(ws.cell(row=r, column=rank_col).value) if rank_col else ""
            barrel = clean_cell(ws.cell(row=r, column=barrel_col).value) if barrel_col else ""
            mag = clean_cell(ws.cell(row=r, column=mag_col).value) if mag_col else ""
            origin = clean_cell(ws.cell(row=r, column=origin_col).value) if origin_col else ""
            notes = clean_cell(ws.cell(row=r, column=notes_col).value) if notes_col else ""

            rows.append(
                AegisRow(
                    sheet=ws.title,
                    row_number=r,
                    name=name,
                    tier=tier,
                    rank=rank,
                    perk1_names=perk1,
                    perk2_names=perk2,
                    barrel=barrel,
                    mag=mag,
                    origin_trait=origin,
                    notes=notes,
                )
            )

    return rows


def get_json_url(url: str, api_key: str | None = None) -> Any:
    headers = {"User-Agent": "aegis-dim-wishlist-builder/1.0"}
    if api_key:
        headers["X-API-Key"] = api_key
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req) as response:
        return json.loads(response.read().decode("utf-8"))


def download_file(url: str, dest: Path, api_key: str | None = None) -> None:
    headers = {"User-Agent": "aegis-dim-wishlist-builder/1.0"}
    if api_key:
        headers["X-API-Key"] = api_key
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req) as response:
        dest.write_bytes(response.read())


def download_manifest_db(cache_dir: Path, api_key: str | None = None) -> Path:
    cache_dir.mkdir(parents=True, exist_ok=True)
    manifest = get_json_url(f"{BUNGIE_ROOT}/Platform/Destiny2/Manifest/", api_key=api_key)
    response = manifest.get("Response", manifest)

    paths = response.get("mobileWorldContentPaths", {}).get("en")
    if not paths:
        # Some responses use jsonWorldComponentContentPaths, but mobile sqlite is much easier.
        raise RuntimeError("Could not find mobileWorldContentPaths['en'] in Bungie manifest response.")

    zip_url = BUNGIE_ROOT + paths
    zip_path = cache_dir / "destiny_manifest.zip"
    download_file(zip_url, zip_path, api_key=api_key)

    with zipfile.ZipFile(zip_path, "r") as zf:
        db_names = [n for n in zf.namelist() if n.endswith((".content", ".sqlite", ".db"))]
        if not db_names:
            db_names = zf.namelist()
        db_name = db_names[0]
        zf.extract(db_name, cache_dir)

    return cache_dir / db_name


def iter_inventory_item_json(db_path: Path) -> Iterable[dict[str, Any]]:
    conn = sqlite3.connect(db_path)
    try:
        # Table can be named DestinyInventoryItemDefinition in current manifests.
        table_names = {
            row[0] for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
        }
        table = "DestinyInventoryItemDefinition"
        if table not in table_names:
            candidates = [t for t in table_names if "InventoryItemDefinition" in t]
            if not candidates:
                raise RuntimeError("Could not find DestinyInventoryItemDefinition table.")
            table = candidates[0]

        # Usually has columns id, json.
        for (json_text,) in conn.execute(f"SELECT json FROM {table}"):
            try:
                yield json.loads(json_text)
            except Exception:
                continue
    finally:
        conn.close()


def build_manifest_index(db_path: Path) -> ManifestIndex:
    weapons_by_name: dict[str, list[int]] = defaultdict(list)
    plugs_by_name: dict[str, list[int]] = defaultdict(list)
    all_item_names_by_hash: dict[int, str] = {}

    for item in iter_inventory_item_json(db_path):
        display = item.get("displayProperties") or {}
        name = clean_cell(display.get("name"))
        if not name:
            continue

        h = item.get("hash")
        if h is None:
            continue
        try:
            h = int(h)
        except Exception:
            continue

        all_item_names_by_hash[h] = name
        key = norm_name(name)

        item_type = int(item.get("itemType", -1))
        item_subtype = int(item.get("itemSubType", -1))
        inventory = item.get("inventory") or {}
        bucket_hash = inventory.get("bucketTypeHash")

        # Weapon items are itemType 3. Bucket can be kinetic/energy/power, but itemType is usually enough.
        if item_type == 3:
            weapons_by_name[key].append(h)

        # Perks, traits, barrels, mags, origin traits, etc. are plug items.
        plug = item.get("plug")
        if plug is not None:
            # DIM docs say use InventoryItem versions, not SandboxPerkDefinition.
            plugs_by_name[key].append(h)

    def dedupe_sorted(d: dict[str, list[int]]) -> dict[str, list[int]]:
        return {k: sorted(set(v)) for k, v in d.items()}

    return ManifestIndex(
        weapons_by_name=dedupe_sorted(weapons_by_name),
        plugs_by_name=dedupe_sorted(plugs_by_name),
        all_item_names_by_hash=all_item_names_by_hash,
    )


def safe_note(text: str) -> str:
    # DIM notes live after #notes:. Keep it one line.
    text = text.replace("\n", "\\n").replace("\r", "")
    text = text.replace("#", "")
    return text.strip()


def row_note(row: AegisRow) -> str:
    bits = [
        f"{row.name} - Rank {row.rank}, Tier {row.tier} by TheAegisRelic".strip(),
        f"Sheet: {row.sheet}",
        f"Column 1: {', '.join(row.perk1_names)}",
        f"Column 2: {', '.join(row.perk2_names)}",
    ]
    if row.barrel:
        bits.append(f"Barrel: {row.barrel}")
    if row.mag:
        bits.append(f"Mag: {row.mag}")
    if row.origin_trait:
        bits.append(f"Origin Trait: {row.origin_trait}")
    if row.notes:
        bits.append(f"Notes: {row.notes}")
    return safe_note("\\n".join(bits))


def generate_dim_lines(
    rows: list[AegisRow],
    index: ManifestIndex,
    tier_filter: set[str],
    include_origin_trait: bool = False,
    max_hashes_per_name: int = 20,
) -> tuple[list[str], list[dict[str, str]]]:
    lines: list[str] = []
    unresolved: list[dict[str, str]] = []
    seen_lines: set[str] = set()

    selected = [r for r in rows if r.tier in tier_filter]

    for row in selected:
        weapon_hashes = index.weapons_by_name.get(norm_name(row.name), [])
        if not weapon_hashes:
            unresolved.append({
                "sheet": row.sheet,
                "row": str(row.row_number),
                "weapon": row.name,
                "type": "weapon",
                "missing": row.name,
                "tier": row.tier,
                "rank": row.rank,
            })
            continue

        perk1_hashes_by_name: list[tuple[str, list[int]]] = []
        perk2_hashes_by_name: list[tuple[str, list[int]]] = []
        origin_hashes: list[int] = []

        for p in row.perk1_names:
            hashes = index.plugs_by_name.get(norm_name(p), [])
            if not hashes:
                unresolved.append({
                    "sheet": row.sheet,
                    "row": str(row.row_number),
                    "weapon": row.name,
                    "type": "perk_1",
                    "missing": p,
                    "tier": row.tier,
                    "rank": row.rank,
                })
            else:
                perk1_hashes_by_name.append((p, hashes[:max_hashes_per_name]))

        for p in row.perk2_names:
            hashes = index.plugs_by_name.get(norm_name(p), [])
            if not hashes:
                unresolved.append({
                    "sheet": row.sheet,
                    "row": str(row.row_number),
                    "weapon": row.name,
                    "type": "perk_2",
                    "missing": p,
                    "tier": row.tier,
                    "rank": row.rank,
                })
            else:
                perk2_hashes_by_name.append((p, hashes[:max_hashes_per_name]))

        if include_origin_trait and row.origin_trait:
            origin_hashes = index.plugs_by_name.get(norm_name(row.origin_trait), [])[:max_hashes_per_name]
            if not origin_hashes:
                unresolved.append({
                    "sheet": row.sheet,
                    "row": str(row.row_number),
                    "weapon": row.name,
                    "type": "origin_trait",
                    "missing": row.origin_trait,
                    "tier": row.tier,
                    "rank": row.rank,
                })

        if not perk1_hashes_by_name or not perk2_hashes_by_name:
            continue

        note = row_note(row)

        lines.append("")
        lines.append(f"//notes:{note}")

        for weapon_hash in weapon_hashes[:max_hashes_per_name]:
            for _, p1_hashes in perk1_hashes_by_name:
                for _, p2_hashes in perk2_hashes_by_name:
                    for p1_hash, p2_hash in itertools.product(p1_hashes, p2_hashes):
                        perk_hashes = [p1_hash, p2_hash]
                        if include_origin_trait and origin_hashes:
                            for origin_hash in origin_hashes:
                                candidate = f"dimwishlist:item={weapon_hash}&perks={','.join(map(str, perk_hashes + [origin_hash]))}"
                                if candidate not in seen_lines:
                                    seen_lines.add(candidate)
                                    lines.append(candidate)
                        else:
                            candidate = f"dimwishlist:item={weapon_hash}&perks={','.join(map(str, perk_hashes))}"
                            if candidate not in seen_lines:
                                seen_lines.add(candidate)
                                lines.append(candidate)

    return lines, unresolved


def write_wishlist(path: Path, title: str, description: str, body_lines: list[str]) -> None:
    header = [
        f"title:{title}",
        f"description:{description}",
        "// Generated from the current Aegis Endgame Analysis Excel file.",
        "// Strict matching mode: weapon + one recommended Column 1 perk + one recommended Column 2 perk.",
        "// Barrels, mags, masterworks, and origin traits are intentionally not required by default.",
        "",
    ]
    path.write_text("\n".join(header + body_lines).strip() + "\n", encoding="utf-8")


def write_unresolved(path: Path, unresolved: list[dict[str, str]]) -> None:
    fieldnames = ["sheet", "row", "weapon", "type", "missing", "tier", "rank"]
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in unresolved:
            writer.writerow(row)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", default="data/endgame.xlsx", help="Path to Aegis Endgame Analysis .xlsx")
    parser.add_argument("--outdir", default=".", help="Output directory")
    parser.add_argument("--cache-dir", default=".cache", help="Cache directory for Bungie manifest DB")
    parser.add_argument("--include-origin-trait", action="store_true", help="Require origin trait in DIM matches")
    parser.add_argument("--max-hashes-per-name", type=int, default=20, help="Safety cap for ambiguous name hashes")
    args = parser.parse_args()

    excel_path = Path(args.excel)
    outdir = Path(args.outdir)
    cache_dir = Path(args.cache_dir)

    if not excel_path.exists():
        raise SystemExit(f"Excel file not found: {excel_path}")

    outdir.mkdir(parents=True, exist_ok=True)

    api_key = os.environ.get("BUNGIE_API_KEY") or None

    print(f"Reading workbook: {excel_path}")
    rows = read_aegis_rows(excel_path)
    print(f"Found {len(rows)} A/S-tier rows with perk recommendations.")

    print("Downloading/loading Destiny manifest...")
    db_path = download_manifest_db(cache_dir, api_key=api_key)
    print(f"Manifest DB: {db_path}")

    print("Building manifest name index...")
    index = build_manifest_index(db_path)
    print(f"Indexed {len(index.weapons_by_name)} weapon names and {len(index.plugs_by_name)} plug names.")

    outputs = [
        ("Aegis-Endgame-S.txt", {"S"}, "Aegis Endgame S-tier", "Current Aegis Endgame Analysis S-tier weapons."),
        ("Aegis-Endgame-A.txt", {"A"}, "Aegis Endgame A-tier", "Current Aegis Endgame Analysis A-tier weapons."),
        ("Aegis-Endgame-A-and-S.txt", {"A", "S"}, "Aegis Endgame A and S-tier", "Current Aegis Endgame Analysis A/S-tier weapons."),
    ]

    all_unresolved: list[dict[str, str]] = []

    for filename, tier_filter, title, description in outputs:
        body, unresolved = generate_dim_lines(
            rows,
            index,
            tier_filter=tier_filter,
            include_origin_trait=args.include_origin_trait,
            max_hashes_per_name=args.max_hashes_per_name,
        )
        write_wishlist(outdir / filename, title, description, body)
        all_unresolved.extend(unresolved)
        dim_line_count = sum(1 for line in body if line.startswith("dimwishlist:"))
        print(f"Wrote {filename}: {dim_line_count} DIM wishlist lines; {len(unresolved)} unresolved references.")

    write_unresolved(outdir / "Aegis-Unresolved-Names.csv", all_unresolved)
    print(f"Wrote Aegis-Unresolved-Names.csv with {len(all_unresolved)} unresolved references.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
