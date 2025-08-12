#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
qvm2vplanx.py — robust XLSX -> vplanx converter (std-lib only).

Key options:
  --sheet N                 1-based worksheet index (default 1)
  --header-row R            1-based header row; use header names (Title/Link/Description/Type)
  --title-col A --link-col C [--desc-col B --type-col D]
                           column letters when there is no header row
  --show-preview            print first 10 rows to help choose row/column numbers
  --no-gzip                 write plain XML instead of .vplanx.gz
"""
import sys, os, re, uuid, time, argparse, logging, gzip, zipfile
import xml.etree.ElementTree as ET

log = logging.getLogger("qvm2vplanx")
logging.basicConfig(level=logging.INFO)
VERSION = "1.2.0"

# ---------- helpers ----------
def col_to_idx(s: str) -> int:
    """Excel column letter -> 0-based index. E.g. A->0, B->1, AA->26"""
    s = s.strip().upper()
    if not s or not re.fullmatch(r"[A-Z]+", s):
        raise ValueError(f"Bad column letter: {s!r}")
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n - 1

def read_shared_strings(z):
    ss = []
    try:
        xml = z.read("xl/sharedStrings.xml")
    except KeyError:
        return ss
    root = ET.fromstring(xml)
    for si in root.findall("{*}si"):
        text = "".join(t.text or "" for t in si.findall(".//{*}t"))
        ss.append(text)
    return ss

def pick_sheet_file(z, sheet_index: int) -> str:
    sheets = sorted([n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n)])
    if not sheets:
        raise RuntimeError("No xl/worksheets/sheetN.xml files in workbook.")
    if sheet_index < 1 or sheet_index > len(sheets):
        raise RuntimeError(f"--sheet {sheet_index} out of range 1..{len(sheets)}")
    return sheets[sheet_index-1]

def read_rows_from_sheet(z, sheet_file: str, shared_strings):
    root = ET.fromstring(z.read(sheet_file))
    rows = []
    for row in root.findall(".//{*}row"):
        vals = []
        for c in row.findall("{*}c"):
            v = c.find("{*}v")
            if v is None:
                vals.append("")
                continue
            if c.get("t") == "s":
                idx = int(v.text)
                vals.append(shared_strings[idx] if idx < len(shared_strings) else "")
            else:
                vals.append(v.text or "")
        rows.append(vals)
    return rows

def preview(rows, n=10):
    print("--- preview: first rows ---")
    for i, r in enumerate(rows[:n], start=1):
        print(f"Row {i:2d}:", [str(x) for x in r])
    print("---------------------------")

# ---------- parse workbook ----------
def parse_xlsx(xlsx_path, sheet_index=1, header_row=None,
               title_col=None, link_col=None, desc_col=None, type_col=None,
               show_preview=False):
    if not zipfile.is_zipfile(xlsx_path):
        raise RuntimeError(f"Not a valid XLSX: {xlsx_path}")
    with zipfile.ZipFile(xlsx_path, "r") as z:
        ss = read_shared_strings(z)
        sheet_file = pick_sheet_file(z, sheet_index)
        rows = read_rows_from_sheet(z, sheet_file, ss)

    if show_preview:
        preview(rows)

    # If explicit column letters are provided -> column-based mode (no headers)
    if title_col and link_col:
        t_idx = col_to_idx(title_col)
        l_idx = col_to_idx(link_col)
        d_idx = col_to_idx(desc_col) if desc_col else None
        ty_idx = col_to_idx(type_col) if type_col else None
        start_row = (header_row or 1)  # if user gave header_row, treat it as data start row
        entries = []
        for r in rows[start_row:]:
            if t_idx >= len(r): continue
            title = r[t_idx].strip()
            if not title: continue
            link = (r[l_idx].strip() if l_idx < len(r) else "")
            desc = (r[d_idx].strip() if d_idx is not None and d_idx < len(r) else "")
            typ  = (r[ty_idx].strip() if ty_idx is not None and ty_idx < len(r) else "")
            entries.append({"title": title, "desc": desc, "link": link, "type": typ})
        if not entries:
            raise RuntimeError("No entries found using the specified column letters.")
        return entries

    # Otherwise, header-based mode: find header row and map names
    if header_row:  # 1-based
        h = header_row - 1
        if h >= len(rows):
            raise RuntimeError(f"--header-row {header_row} beyond last row ({len(rows)})")
        header = [str(x).strip().lower() for x in rows[h]]
    else:
        # heuristic: scan first 30 rows for both 'title' and 'link'
        header, h = None, None
        for i, r in enumerate(rows[:30]):
            low = [str(x).strip().lower() for x in r]
            if any("title" in x for x in low) and any("link" in x for x in low):
                header, h = low, i
                break
        if header is None:
            raise RuntimeError("Header with 'Title' and 'Link' not found. "
                               "Use --header-row or --title-col/--link-col.")

    name_to_idx = {}
    for j, name in enumerate(header):
        if "title" in name and "title" not in name_to_idx: name_to_idx["title"] = j
        elif "link" in name and "link" not in name_to_idx: name_to_idx["link"] = j
        elif "description" in name and "description" not in name_to_idx: name_to_idx["description"] = j
        elif name == "type" and "type" not in name_to_idx: name_to_idx["type"] = j

    if "title" not in name_to_idx or "link" not in name_to_idx:
        raise RuntimeError("Header row exists but missing required 'Title' and 'Link'. "
                           "Use column letters instead (e.g. --title-col A --link-col C).")

    entries = []
    for r in rows[(header_row - 1 + 1) if header_row else (h + 1):]:
        title = r[name_to_idx["title"]].strip() if name_to_idx["title"] < len(r) else ""
        if not title: continue
        link  = r[name_to_idx["link"]].strip()  if name_to_idx["link"]  < len(r) else ""
        desc  = r[name_to_idx["description"]].strip() if "description" in name_to_idx and name_to_idx["description"] < len(r) else ""
        typ   = r[name_to_idx["type"]].strip() if "type" in name_to_idx and name_to_idx["type"] < len(r) else ""
        entries.append({"title": title, "desc": desc, "link": link, "type": typ})

    if not entries:
        raise RuntimeError("No entries found under the detected header row.")
    return entries

# ---------- build vplanx ----------
def build_vplanx(entries, plan_name):
    ET.register_namespace("vplanx", "http://www.cadence.com/vplanx")
    root = ET.Element("vplanx:plan", {"xmlns:vplanx": "http://www.cadence.com/vplanx"})
    md = ET.SubElement(root, "metaData", {"id": str(uuid.uuid1())})
    ET.SubElement(md, "name").text = plan_name
    ET.SubElement(md, "planId").text = md.get("id")
    ET.SubElement(md, "sourceTool").text = "qvm2vplanx"
    ET.SubElement(md, "toolVersion").text = VERSION
    ET.SubElement(md, "schemaVersion").text = "1.0"
    ET.SubElement(md, "buildTime").text = time.strftime("%Y-%m-%d %H:%M:%S")

    relem = ET.SubElement(root, "rootElements")
    for e in entries:
        sect = ET.SubElement(relem, "section", {"id": str(uuid.uuid1())})
        ET.SubElement(sect, "name").text = e["title"]
        attrs = ET.SubElement(sect, "attributes")
        for k, v in [("details", e["desc"]), ("type", e["type"]), ("planned_elements", "1")]:
            a = ET.SubElement(attrs, "attribute")
            ET.SubElement(a, "name").text = k
            ET.SubElement(a, "value").text = v
        mp = ET.SubElement(sect, "metricsPort", {"id": str(uuid.uuid1())})
        ET.SubElement(mp, "name").text = e["title"]
        mps = ET.SubElement(mp, "mappingPatterns")
        mpn = ET.SubElement(mps, "mappingPattern", {"id": str(uuid.uuid1())})
        dms = ET.SubElement(mpn, "domains")
        d = ET.SubElement(dms, "domain"); d.text = "HDL"
        ems = ET.SubElement(mpn, "entityKinds")
        ek = ET.SubElement(ems, "entityKind"); ek.text = "INSTANCE"
        ET.SubElement(mpn, "pattern").text = e["link"]
    return ET.ElementTree(root)

def save_tree(tree, out_path, gzip_out=True):
    with (gzip.open(out_path, "wb") if gzip_out else open(out_path, "wb")) as f:
        tree.write(f, encoding="utf-8", xml_declaration=True)

# ---------- CLI ----------
def main(argv):
    ap = argparse.ArgumentParser(description="Convert XLSX to Cadence vPlanx (std-lib only)")
    ap.add_argument("xlsx", help="Input .xlsx file")
    ap.add_argument("--out", help="Output filename (default: <xlsxname>.vplanx)")
    ap.add_argument("--no-gzip", action="store_true", help="Write plain XML instead of gzipped .vplanx")

    ap.add_argument("--sheet", type=int, default=1, help="Worksheet index (1-based). Default 1")
    ap.add_argument("--header-row", type=int, help="Header row (1-based). If omitted, script searches for Title/Link")
    ap.add_argument("--title-col", help="Column letter for Title when no header row (e.g. A)")
    ap.add_argument("--link-col",  help="Column letter for Link when no header row (e.g. C)")
    ap.add_argument("--desc-col",  help="Column letter for Description (optional)")
    ap.add_argument("--type-col",  help="Column letter for Type (optional)")
    ap.add_argument("--show-preview", action="store_true", help="Print first 10 rows for debugging")

    args = ap.parse_args(argv)

    if not os.path.exists(args.xlsx):
        ap.error(f"file not found: {args.xlsx}")

    try:
        entries = parse_xlsx(
            args.xlsx,
            sheet_index=args.sheet,
            header_row=args.header_row,
            title_col=args.title_col,
            link_col=args.link_col,
            desc_col=args.desc_col,
            type_col=args.type_col,
            show_preview=args.show_preview,
        )
    except Exception as e:
        log.error(str(e))
        sys.exit(2)

    plan_name = os.path.splitext(os.path.basename(args.xlsx))[0]
    tree = build_vplanx(entries, plan_name)
    outname = args.out or (plan_name + ".vplanx")
    save_tree(tree, outname, gzip_out=not args.no_gzip)
    print("Created", outname)

if __name__ == "__main__":
    main(sys.argv[1:])
