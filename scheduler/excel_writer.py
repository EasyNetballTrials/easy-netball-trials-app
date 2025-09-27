# scheduler/excel_writer.py
import io, zipfile, re
from xml.etree import ElementTree as ET
from typing import List, Dict

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r":    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkg":  "http://schemas.openxmlformats.org/package/2006/relationships",
}
for k, v in NS.items():
    ET.register_namespace("" if k == "main" else k, v)

HEADERS = [
    "First Name", "Last Name",
    "PreferredPos1", "PreferredPos2", "PreferredPos3",
    "Active", "Number", "Seed"
]
FIRST_DATA_ROW = 2
LAST_DATA_COL = 8  # H

# ---------- small helpers ----------
def _col_letter(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s

def _col_index_from_ref(ref: str) -> int:
    m = re.match(r"([A-Z]+)", ref or "")
    if not m: return 1
    col = 0
    for ch in m.group(1):
        col = col * 26 + (ord(ch) - 64)
    return col

def _cell_ref(row: int, col: int) -> str:
    return f"{_col_letter(col)}{row}"

def _find_players_sheet_path(zf: zipfile.ZipFile) -> str:
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    r_id = None
    for sh in wb.findall("main:sheets/main:sheet", NS):
        if (sh.get("name") or "").strip().lower() == "players":
            r_id = sh.get(f"{{{NS['r']}}}id"); break
    if not r_id:
        raise ValueError("Couldn't find a sheet named 'Players'.")

    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    tgt = None
    for rel in rels.findall("pkg:Relationship", NS):
        if rel.get("Id") == r_id:
            tgt = rel.get("Target"); break
    if not tgt: raise ValueError("Players sheet rel target not found.")
    if not tgt.startswith("worksheets/"):
        tgt = re.sub(r"^/?xl/", "", tgt)
    return f"xl/{tgt}"

def _get_or_create_row(sheetData: ET.Element, r_index: int) -> ET.Element:
    for row in sheetData.findall("main:row", NS):
        if row.get("r") == str(r_index):
            return row
    row = ET.Element(f"{{{NS['main']}}}row", {"r": str(r_index)})
    sheetData.append(row)
    return row

def _find_cell(row_el: ET.Element, ref: str):
    for c in row_el.findall("main:c", NS):
        if c.get("r") == ref:
            return c
    return None

def _insert_cell_sorted(row_el: ET.Element, ref: str) -> ET.Element:
    # ensure a <c r="ref"> exists and is in ascending column order
    c = _find_cell(row_el, ref)
    if c is not None:
        return c
    new_c = ET.Element(f"{{{NS['main']}}}c", {"r": ref})
    # find insert point
    new_idx = _col_index_from_ref(ref)
    cells = list(row_el.findall("main:c", NS))
    inserted = False
    for i, existing in enumerate(cells):
        if _col_index_from_ref(existing.get("r", "A1")) > new_idx:
            row_el.insert(i, new_c)
            inserted = True
            break
    if not inserted:
        row_el.append(new_c)
    return new_c

def _set_inline_text(cell_el: ET.Element, text: str):
    # keep style attributes (s, etc.), clear only value-related children/attrs
    for child in list(cell_el):
        if child.tag in (f"{{{NS['main']}}}v", f"{{{NS['main']}}}is", f"{{{NS['main']}}}f"):
            cell_el.remove(child)
    if "t" in cell_el.attrib:
        del cell_el.attrib["t"]

    if text is None or text == "":
        return  # leave an empty cell to preserve style/protection

    cell_el.set("t", "inlineStr")
    is_ = ET.SubElement(cell_el, f"{{{NS['main']}}}is")
    t = ET.SubElement(is_, f"{{{NS['main']}}}t")
    # preserve leading/trailing spaces if present
    if text.strip() != text:
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = str(text)

def _scan_prev_max_row(sheetData: ET.Element) -> int:
    max_r = FIRST_DATA_ROW - 1
    valid_cols = {_col_letter(i) for i in range(1, LAST_DATA_COL + 1)}
    for row in sheetData.findall("main:row", NS):
        try:
            r_idx = int(row.get("r", "0"))
        except ValueError:
            continue
        if r_idx < FIRST_DATA_ROW: continue
        any_val = False
        for c in row.findall("main:c", NS):
            ref = c.get("r", "")
            m = re.match(r"([A-Z]+)", ref or "")
            if not m or m.group(1) not in valid_cols:
                continue
            if c.find("main:v", NS) is not None:
                any_val = True; break
            is_ = c.find("main:is", NS)
            if is_ is not None and is_.find("main:t", NS) is not None and (is_.find("main:t", NS).text or "") != "":
                any_val = True; break
        if any_val and r_idx > max_r:
            max_r = r_idx
    return max_r

# ---------- main writer ----------
def inject_players_csv(template_path: str,
                       rows: List[Dict[str, str]],
                       out_stream: io.BytesIO,
                       shell_mode: bool = True) -> None:
    with open(template_path, "rb") as f:
        blob = f.read()

    zin = zipfile.ZipFile(io.BytesIO(blob), "r")
    players_xml_path = _find_players_sheet_path(zin)

    ws = ET.fromstring(zin.read(players_xml_path))
    sheetData = ws.find("main:sheetData", NS)
    if sheetData is None:
        sheetData = ET.SubElement(ws, f"{{{NS['main']}}}sheetData")

    prev_max = _scan_prev_max_row(sheetData)

    # Write rows
    r_idx = FIRST_DATA_ROW
    for rowdict in rows:
        r_el = _get_or_create_row(sheetData, r_idx)
        # fill A..H in-order
        for col_i, header in enumerate(HEADERS, start=1):
            ref = _cell_ref(r_idx, col_i)
            c_el = _insert_cell_sorted(r_el, ref)
            _set_inline_text(c_el, (rowdict.get(header, "") or ""))
        r_idx += 1

    # Clear leftovers in A..H (keep cells/styles)
    for rr in range(r_idx, prev_max + 1):
        r_el = _get_or_create_row(sheetData, rr)
        for col_i in range(1, LAST_DATA_COL + 1):
            ref = _cell_ref(rr, col_i)
            c_el = _insert_cell_sorted(r_el, ref)
            _set_inline_text(c_el, "")

    # Repack: replace only Players sheet xml
    zout = zipfile.ZipFile(out_stream, "w", compression=zipfile.ZIP_DEFLATED)
    for info in zin.infolist():
        data = zin.read(info.filename)
        if info.filename == players_xml_path:
            data = ET.tostring(ws, encoding="utf-8", xml_declaration=True)
        zi = zipfile.ZipInfo(info.filename)
        zi.compress_type = zipfile.ZIP_DEFLATED
        zi.external_attr = info.external_attr
        zout.writestr(zi, data)
    zout.close(); zin.close()
    out_stream.seek(0)
