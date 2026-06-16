#!/usr/bin/env python3
from __future__ import annotations

import argparse, csv, itertools, json, os, re, sqlite3, urllib.request, zipfile
from collections import defaultdict
from pathlib import Path

import openpyxl

ROOT = "https://www.bungie.net"


def normalize_quotes(x):
    if x is None:
        return ""
    return str(x).strip().replace("\u2019", "'").replace("\u2018", "'").replace("\u201c", '"').replace("\u201d", '"')


def clean(x):
    """Clean display text for notes/matching, collapsing whitespace."""
    s = normalize_quotes(x)
    return re.sub(r"\s+", " ", s).strip()


def norm(x):
    return clean(x).casefold()


def split_perks(x):
    """
    Split perk cells while preserving Excel line breaks.

    The previous version called clean() first, which collapsed newlines into spaces.
    That turned cells like:
        Physic\nBurning Ambition
    into:
        Physic Burning Ambition
    which is not a real perk name and caused hundreds of unresolved rows.
    """
    raw = normalize_quotes(x)
    if not raw:
        return []

    raw = re.sub(r"\bEnhanced\b", "", raw, flags=re.I)

    out, seen = [], set()
    for p in re.split(r"[\n\r,;]+", raw):
        p = clean(p)
        if not p or norm(p) in {"none", "n/a", "na", "-", "?", "need testing"}:
            continue
        k = norm(p)
        if k not in seen:
            seen.add(k)
            out.append(p)
    return out


def find_header(ws):
    max_row = ws.max_row or 0
    max_col = ws.max_column or 0
    if max_row < 1 or max_col < 1:
        return None

    for r in range(1, min(max_row, 30) + 1):
        h = {}
        for c in range(1, max_col + 1):
            v = clean(ws.cell(r, c).value)
            if v:
                h[norm(v)] = c
        if all(x in h for x in ["name", "perk 1", "perk 2", "tier"]):
            return r, h
    return None


def col(h, *names):
    for n in names:
        if norm(n) in h:
            return h[norm(n)]
    return None


def read_rows(xlsx):
    # IMPORTANT: read_only=False is required for this workbook. In read_only mode,
    # some sheets can report max_row/max_column as None, causing zero detected rows.
    wb = openpyxl.load_workbook(xlsx, data_only=True, read_only=False)
    rows = []

    for ws in wb.worksheets:
        found = find_header(ws)
        if not found:
            continue

        hr, h = found
        c_name, c_tier = col(h, "Name"), col(h, "Tier")
        c_rank, c_p1, c_p2 = col(h, "Rank"), col(h, "Perk 1"), col(h, "Perk 2")
        c_barrel, c_mag = col(h, "Barrel"), col(h, "Mag")
        c_origin, c_notes = col(h, "Origin Trait", "Origin Trai"), col(h, "Notes")

        max_row = ws.max_row or 0
        for r in range(hr + 1, max_row + 1):
            name = clean(ws.cell(r, c_name).value)
            tier = clean(ws.cell(r, c_tier).value).upper()
            if not name or tier not in {"S", "A"}:
                continue

            p1 = split_perks(ws.cell(r, c_p1).value)
            p2 = split_perks(ws.cell(r, c_p2).value)
            if not p1 or not p2:
                continue

            rows.append({
                "sheet": ws.title,
                "row": r,
                "name": name,
                "tier": tier,
                "rank": clean(ws.cell(r, c_rank).value) if c_rank else "",
                "p1": p1,
                "p2": p2,
                "barrel": clean(ws.cell(r, c_barrel).value) if c_barrel else "",
                "mag": clean(ws.cell(r, c_mag).value) if c_mag else "",
                "origin": clean(ws.cell(r, c_origin).value) if c_origin else "",
                "notes": clean(ws.cell(r, c_notes).value) if c_notes else "",
            })

    return rows


def get_json(url, key=None):
    headers = {"User-Agent": "aegis-dim-builder/1.1"}
    if key:
        headers["X-API-Key"] = key
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req) as res:
        return json.loads(res.read().decode("utf-8"))


def download(url, path, key=None):
    headers = {"User-Agent": "aegis-dim-builder/1.1"}
    if key:
        headers["X-API-Key"] = key
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req) as res:
        path.write_bytes(res.read())


def manifest_db(cache):
    key = os.environ.get("BUNGIE_API_KEY")
    cache.mkdir(parents=True, exist_ok=True)
    data = get_json(ROOT + "/Platform/Destiny2/Manifest/", key)
    resp = data.get("Response", data)
    rel = resp["mobileWorldContentPaths"]["en"]
    zpath = cache / "manifest.zip"
    download(ROOT + rel, zpath, key)
    with zipfile.ZipFile(zpath) as z:
        name = next((n for n in z.namelist() if n.endswith((".content", ".sqlite", ".db"))), z.namelist()[0])
        z.extract(name, cache)
    return cache / name


def iter_items(db):
    con = sqlite3.connect(db)
    try:
        tables = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        table = "DestinyInventoryItemDefinition" if "DestinyInventoryItemDefinition" in tables else next(t for t in tables if "InventoryItemDefinition" in t)
        for (txt,) in con.execute(f"SELECT json FROM {table}"):
            try:
                yield json.loads(txt)
            except Exception:
                pass
    finally:
        con.close()


def build_index(db):
    weapons, plugs = defaultdict(list), defaultdict(list)
    for it in iter_items(db):
        name = clean((it.get("displayProperties") or {}).get("name"))
        if not name or "hash" not in it:
            continue
        h = int(it["hash"])
        if int(it.get("itemType", -1)) == 3:
            weapons[norm(name)].append(h)
        if it.get("plug") is not None:
            plugs[norm(name)].append(h)
    return {k: sorted(set(v)) for k, v in weapons.items()}, {k: sorted(set(v)) for k, v in plugs.items()}


def note(row):
    bits = [
        f"{row['name']} - Rank {row['rank']}, Tier {row['tier']} by TheAegisRelic",
        f"Sheet: {row['sheet']}",
        "Column 1: " + ", ".join(row["p1"]),
        "Column 2: " + ", ".join(row["p2"]),
    ]
    if row["barrel"]:
        bits.append("Barrel: " + row["barrel"])
    if row["mag"]:
        bits.append("Mag: " + row["mag"])
    if row["origin"]:
        bits.append("Origin Trait: " + row["origin"])
    if row["notes"]:
        bits.append("Notes: " + row["notes"])
    return "\\n".join(bits).replace("#", "").replace("\r", "")


def generate(rows, weapons, plugs, tiers):
    lines, missing, seen = [], [], set()
    for row in rows:
        if row["tier"] not in tiers:
            continue

        wh = weapons.get(norm(row["name"]), [])
        if not wh:
            missing.append([row["sheet"], row["row"], row["name"], "weapon", row["name"], row["tier"], row["rank"]])
            continue

        p1s, p2s = [], []
        for p in row["p1"]:
            hs = plugs.get(norm(p), [])
            if hs:
                p1s.extend(hs[:20])
            else:
                missing.append([row["sheet"], row["row"], row["name"], "perk_1", p, row["tier"], row["rank"]])
        for p in row["p2"]:
            hs = plugs.get(norm(p), [])
            if hs:
                p2s.extend(hs[:20])
            else:
                missing.append([row["sheet"], row["row"], row["name"], "perk_2", p, row["tier"], row["rank"]])

        if not p1s or not p2s:
            continue

        lines += ["", "//notes:" + note(row)]
        for w, p1, p2 in itertools.product(wh[:20], sorted(set(p1s)), sorted(set(p2s))):
            line = f"dimwishlist:item={w}&perks={p1},{p2}"
            if line not in seen:
                seen.add(line)
                lines.append(line)

    return lines, missing


def write_file(path, title, desc, lines):
    head = [
        "title:" + title,
        "description:" + desc,
        "// Generated from current Aegis Endgame Analysis Excel.",
        "// Strict mode: weapon + one recommended Perk 1 + one recommended Perk 2.",
        "// Barrels, mags, masterworks, and origin traits are intentionally not required.",
        "",
    ]
    Path(path).write_text("\n".join(head + lines).strip() + "\n", encoding="utf-8")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", default="data.xlsx")
    ap.add_argument("--outdir", default=".")
    ap.add_argument("--cache-dir", default=".cache")
    args = ap.parse_args()

    rows = read_rows(Path(args.excel))
    print(f"Found {len(rows)} A/S rows")
    if not rows:
        raise SystemExit("No A/S rows found. Check that the Excel workbook has Name, Perk 1, Perk 2, Tier headers.")

    db = manifest_db(Path(args.cache_dir))
    weapons, plugs = build_index(db)
    print(f"Indexed {len(weapons)} weapon names and {len(plugs)} plug names")

    all_missing = []
    for fn, tiers, title, desc in [
        ("Aegis-Endgame-S.txt", {"S"}, "Aegis Endgame S-tier", "Current Aegis Endgame S-tier weapons."),
        ("Aegis-Endgame-A.txt", {"A"}, "Aegis Endgame A-tier", "Current Aegis Endgame A-tier weapons."),
        ("Aegis-Endgame-A-and-S.txt", {"A", "S"}, "Aegis Endgame A and S-tier", "Current Aegis Endgame A/S-tier weapons."),
    ]:
        lines, missing = generate(rows, weapons, plugs, tiers)
        write_file(Path(args.outdir) / fn, title, desc, lines)
        all_missing += missing
        print(f"Wrote {fn}: {sum(1 for l in lines if l.startswith('dimwishlist:'))} DIM lines; {len(missing)} unresolved")

    with open(Path(args.outdir) / "Aegis-Unresolved-Names.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)
        wr.writerow(["sheet", "row", "weapon", "type", "missing", "tier", "rank"])
        wr.writerows(all_missing)


if __name__ == "__main__":
    main()
