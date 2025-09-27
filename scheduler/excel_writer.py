# scheduler/excel_writer.py
import io
import zipfile
import re
from xml.etree import ElementTree as ET
from typing import List, Dict, Iterable

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r":    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkg":  "http://schemas.openxmlformats.org/package/2006/relationships"
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

def _col_letter(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s

def _cell_ref(row: int, col: int) -> str:
    return f"{_col_letter(col)}{row}"

def _find_players_sheet_path(zf: zipfile.ZipFile) -> str:
    # 1) find r:id for the sheet named Players
    wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
    r_id = None
    for sh in wb_xml.findall("main:sheets/main:sheet", NS):
        if (sh.get("name") or "").strip().lower() == "players":
            r_id = sh.get(f"{{{NS['r']}}}id")
            break
    if not r_id:
        raise ValueError("Couldn't find a sheet named 'Players' in the template.")

    # 2) resolve the r:id in workbook rels
    rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    target = None
    for rel in rels_xml.findall("pkg:Relationship", NS):
        if rel.get("Id") == r_id:
            target = rel.get("Target")
            break
    if not target:
        raise ValueError("Couldn't resolve Players sheet target from workbook relationships.")

    # normalize to xl/worksheets/sheetN.xml
    if not target.startswith("worksheets/"):
        target = re.sub(r"^/?xl/", "", target)
    return f"xl/{target}"

def _get_or_create_row(sheetData: ET.Element, r_index: int) -> ET.Element:
    # Try to find an existing row element
    for row in sheetData.findall("main:row", NS):
        if row.get("r") == str(r_index):
            return row
    # Create a new row (try to append at the end without disturbing others)
    new_row = ET.Element(f"{{{NS['main']}}}row", {"r": str(r_index)})
    sheetData.append(new_row)
    return new_row

def _find_cell(row_el: ET.Element, ref: str) -> ET.Element:
    for c in row_el.findall("main:c", NS):
        if c.get("r") == ref:
            return c
    return None

def _ensure_cell(row_el: ET.Element, ref: str) -> ET.Element:
    c = _find_cell(row_el, ref)
    if c is None:
        c = ET.Element(f"{{{NS['main']}}}c", {"r": ref})
        row_el.append(c)
    return c

def _set_inline_text(cell_el: ET.Element, text: str):
    # Clear existing content (v, is, f) but keep style attributes if any
    for tag in list(cell_el):
        if tag.tag in {f"{{{NS['main']}}}v", f"{{{NS['main']}}}is", f"{{{NS['main']}}}f"}:
            cell_el.remove(tag)
    if "t" in cell_el.attrib:
        del cell_el.attrib["t"]

    if text is None or text == "":
        # leave the cell empty (preserves style)
        return

    cell_el.set("t", "inlineStr")
    is_ = ET.SubElement(cell_el, f"{{{NS['main']}}}is")
    t = ET.SubElement(is_, f"{{{NS['main']}}}t")
    if text.strip() != text:
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = str(text)

def _scan_prev_data_max(sheetData: ET.Element) -> int:
    """
    Find the highest row index ≥ FIRST_DATA_ROW that has
    any non-empty cell in cols A..H. Used so we can clear leftovers
    without deleting rows.
    """
    max_r = FIRST_DATA_ROW - 1
    cols = {_col_letter(i) for i in range(1, LAST_DATA_COL + 1)}
    for row in sheetData.findall("main:row", NS):
        try:
            r_idx = int(row.get("r", "0"))
        except ValueError:
            continue
        if r_idx < FIRST_DATA_ROW:
            continue
        any_val = False
        for c in row.findall("main:c", NS):
            # Only consider A..H
            ref = c.get("r", "")
            col_letters = re.match(r"[A-Z]+", ref or "")
            if not col_letters:
                continue
            if col_letters.group(0) not in cols:
                continue
            # Has content if it has <v> or <is>/<t>
            if c.find("main:v", NS) is not None:
                any_val = True; break
            is_ = c.find("main:is", NS)
            if is_ is not None and is_.find("main:t", NS) is not None and (is_.find("main:t", NS).text or "") != "":
                any_val = True; break
        if any_val and r_idx > max_r:
            max_r = r_idx
    return max_r

def inject_players_csv(template_path: str,
                       rows: List[Dict[str, str]],
                       out_stream: io.BytesIO,
                       shell_mode: bool = True) -> None:
    """
    In-place update of Players sheet cells A..H, starting row 2.
    - Does NOT delete any rows
    - Does NOT change sheet dimension
    - Preserves styles, drawings, macros, bib cells, buttons, etc.
    """
    with open(template_path, "rb") as f:
        blob = f.read()

    zin = zipfile.ZipFile(io.BytesIO(blob), "r")
    players_xml_path = _find_players_sheet_path(zin)

    # Load sheet xml
    sheet_xml = ET.fromstring(zin.read(players_xml_path))
    sheetData = sheet_xml.find("main:sheetData", NS)
    if sheetData is None:
        sheetData = ET.SubElement(sheet_xml, f"{{{NS['main']}}}sheetData")

    # How far did the template previously have data in A..H?
    prev_max = _scan_prev_data_max(sheetData)

    # How far will we need now?
    new_max = max(prev_max, FIRST_DATA_ROW + len(rows) - 1)

    # Write new data rows
    r_idx = FIRST_DATA_ROW
    for row_dict in rows:
        r_el = _get_or_create_row(sheetData, r_idx)
        for ci, header in enumerate(HEADERS, start=1):
            ref = _cell_ref(r_idx, ci)
            c_el = _ensure_cell(r_el, ref)
            _set_inline_text(c_el, str(row_dict.get(header, "")) if row_dict.get(header) is not None else "")
        r_idx += 1

    # Clear any leftover old data in A..H (but keep the cells & styles)
    for rr in range(r_idx, prev_max + 1):
        r_el = _get_or_create_row(sheetData, rr)
        for ci in range(1, LAST_DATA_COL + 1):
            ref = _cell_ref(rr, ci)
            c_el = _ensure_cell(r_el, ref)
            _set_inline_text(c_el, "")

    # Rebuild the xlsm: copy all parts, replace only Players sheet xml
    out_zip = zipfile.ZipFile(out_stream, "w", compression=zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == players_xml_path:
            data = ET.tostring(sheet_xml, encoding="utf-8", xml_declaration=True)
        zi = zipfile.ZipInfo(item.filename)
        zi.compress_type = zipfile.ZIP_DEFLATED
        zi.external_attr = item.external_attr
        out_zip.writestr(zi, data)
    out_zip.close()
    zin.close()
    out_stream.seek(0)
