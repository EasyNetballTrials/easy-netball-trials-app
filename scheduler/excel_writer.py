# scheduler/excel_writer.py
import io
import zipfile
import re
from xml.etree import ElementTree as ET
from typing import List, Dict, Iterable

# Namespaces used by Excel's XML
NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r":    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
}
for k, v in NS.items():
    ET.register_namespace("" if k == "main" else k, v)  # keep original prefixes in output

def _col_letter(idx: int) -> str:
    """1-based column index -> Excel column letters."""
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s

def _cell_ref(row: int, col: int) -> str:
    return f"{_col_letter(col)}{row}"

def _find_players_sheet_path(zf: zipfile.ZipFile) -> str:
    """
    Map the sheet named 'Players' -> worksheets/sheetN.xml
    by reading xl/workbook.xml and xl/_rels/workbook.xml.rels
    """
    wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
    # 1) find r:id for sheet name="Players"
    r_id = None
    for sh in wb_xml.findall("main:sheets/main:sheet", NS):
        if (sh.get("name") or "").strip().lower() == "players":
            r_id = sh.get(f"{{{NS['r']}}}id")
            break
    if not r_id:
        raise ValueError("Couldn't find a sheet named 'Players' in the template.")

    # 2) resolve r:id in workbook relationships
    rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    target = None
    for rel in rels_xml.findall("main:Relationship", {"main": "http://schemas.openxmlformats.org/package/2006/relationships"}):
        if rel.get("Id") == r_id:
            target = rel.get("Target")
            break
    if not target:
        raise ValueError("Couldn't resolve Players sheet target from workbook relationships.")

    # Normalize (usually 'worksheets/sheetN.xml')
    if not target.startswith("worksheets/"):
        # sometimes starts with '/xl/worksheets/...'
        target = re.sub(r"^/?xl/", "", target)
    return f"xl/{target}"

def _build_inline_string_cell(ref: str, text: str) -> ET.Element:
    """Create <c r='A1' t='inlineStr'><is><t>text</t></is></c>"""
    c = ET.Element(f"{{{NS['main']}}}c", {"r": ref, "t": "inlineStr"})
    is_ = ET.SubElement(c, f"{{{NS['main']}}}is")
    t = ET.SubElement(is_, f"{{{NS['main']}}}t")
    # Excel expects XML-escaped text; ElementTree handles that for us.
    # Preserve leading/trailing spaces by setting xml:space='preserve' if needed:
    if text.strip() != text:
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = "" if text is None else str(text)
    return c

def _clear_sheetdata_from_row(sheet_xml: ET.Element, start_row: int):
    """Remove all <row> elements with r >= start_row (we keep header row 1)."""
    sheetData = sheet_xml.find("main:sheetData", NS)
    if sheetData is None:
        # If missing, create it to keep Excel happy
        sheetData = ET.SubElement(sheet_xml, f"{{{NS['main']}}}sheetData")
        return

    for row in list(sheetData):
        r_attr = row.get("r")
        try:
            r_val = int(r_attr)
        except (TypeError, ValueError):
            continue
        if r_val >= start_row:
            sheetData.remove(row)

def _append_rows(sheet_xml: ET.Element, rows: Iterable[Dict[str, str]], headers: List[str]):
    """
    Append CSV rows starting at row 2.
    All values written as inline strings.
    Columns order: headers left-to-right beginning at A.
    """
    sheetData = sheet_xml.find("main:sheetData", NS)
    if sheetData is None:
        sheetData = ET.SubElement(sheet_xml, f"{{{NS['main']}}}sheetData")

    current_row = 2
    max_col = len(headers)

    for data in rows:
        r_el = ET.SubElement(sheetData, f"{{{NS['main']}}}row", {"r": str(current_row)})
        for ci, header in enumerate(headers, start=1):
            val = data.get(header, "")
            cell = _build_inline_string_cell(_cell_ref(current_row, ci), "" if val is None else str(val))
            r_el.append(cell)
        current_row += 1

    # Update the dimension ref (e.g., A1:Hn)
    dim = sheet_xml.find("main:dimension", NS)
    if dim is None:
        dim = ET.SubElement(sheet_xml, f"{{{NS['main']}}}dimension")
    last_row = max(current_row - 1, 1)
    last_ref = _cell_ref(last_row, max_col if max_col > 0 else 1)
    dim.set("ref", f"A1:{last_ref}")

def inject_players_csv(template_path: str,
                       rows: List[Dict[str, str]],
                       out_stream: io.BytesIO,
                       shell_mode: bool = True) -> None:
    """
    Copy the .xlsm template, replace Players sheet body (A2:Hn) with CSV data.
    'rows' must be a list of dicts keyed by your CSV headers.
    We DO NOT touch macros or any other parts of the file.
    """
    # Decide which headers we output and in what order:
    headers = [
        "First Name", "Last Name",
        "PreferredPos1", "PreferredPos2", "PreferredPos3",
        "Active", "Number", "Seed"
    ]

    # 1) Read template into memory zip
    with open(template_path, "rb") as f:
        blob = f.read()

    src = io.BytesIO(blob)
    zin = zipfile.ZipFile(src, "r")

    # 2) Locate the Players sheet xml path
    players_xml_path = _find_players_sheet_path(zin)

    # 3) Build a new zip in memory while replacing just that sheet
    zout = zipfile.ZipFile(out_stream, "w", compression=zipfile.ZIP_DEFLATED)

    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == players_xml_path:
            # Parse, clear body from row 2, and append our data
            sheet_xml = ET.fromstring(data)
            _clear_sheetdata_from_row(sheet_xml, start_row=2)
            _append_rows(sheet_xml, rows, headers)
            data = ET.tostring(sheet_xml, encoding="utf-8", xml_declaration=True)
        # Write file (unchanged or modified) into destination zip
        zi = zipfile.ZipInfo(item.filename)
        zi.compress_type = zipfile.ZIP_DEFLATED
        zi.external_attr = item.external_attr  # keep permissions
        zout.writestr(zi, data)

    zout.close()
    zin.close()
    out_stream.seek(0)
