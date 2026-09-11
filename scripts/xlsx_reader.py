"""Small dependency-free reader for cell values in ordinary .xlsx worksheets."""

from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": MAIN, "r": REL}


def _column_number(reference):
    letters = re.match(r"[A-Z]+", reference).group()
    number = 0
    for letter in letters:
        number = number * 26 + ord(letter) - 64
    return number


def _number(text):
    if text in (None, ""):
        return ""
    try:
        value = float(text)
        return int(value) if value.is_integer() else value
    except ValueError:
        return text


def read_xlsx(path, sheet_name=None):
    path = Path(path)
    with zipfile.ZipFile(path) as archive:
        shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall(f"{{{MAIN}}}si"):
                shared.append("".join(node.text or "" for node in item.iter(f"{{{MAIN}}}t")))

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rels = {item.attrib["Id"]: item.attrib["Target"] for item in rels_root.findall(f"{{{PKG_REL}}}Relationship")}
        choices = []
        for sheet in workbook.find("m:sheets", NS):
            name = sheet.attrib["name"]
            target = rels[sheet.attrib[f"{{{REL}}}id"]].lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            choices.append((name, target))
        if sheet_name is None:
            selected_name, target = choices[0]
        else:
            selected_name, target = next(item for item in choices if item[0] == sheet_name)

        root = ET.fromstring(archive.read(target))
        rows = []
        for row in root.findall(".//m:row", NS):
            cells = {}
            for cell in row.findall("m:c", NS):
                value = cell.find("m:v", NS)
                if value is None:
                    inline = cell.find("m:is", NS)
                    if inline is None:
                        continue
                    text = "".join(node.text or "" for node in inline.iter(f"{{{MAIN}}}t"))
                else:
                    text = value.text or ""
                    if cell.attrib.get("t") == "s":
                        text = shared[int(text)]
                    elif cell.attrib.get("t") not in ("str", "inlineStr"):
                        text = _number(text)
                cells[_column_number(cell.attrib["r"])] = text
            width = max(cells, default=0)
            rows.append([cells.get(column, "") for column in range(1, width + 1)])
        return selected_name, rows


def sheet_names(path):
    with zipfile.ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        return [sheet.attrib["name"] for sheet in workbook.find("m:sheets", NS)]
