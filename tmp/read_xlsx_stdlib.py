import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def col_number(ref):
    letters = re.match(r"[A-Z]+", ref).group()
    out = 0
    for letter in letters:
        out = out * 26 + ord(letter) - 64
    return out


with zipfile.ZipFile(sys.argv[1]) as archive:
    shared = []
    if "xl/sharedStrings.xml" in archive.namelist():
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        for item in root.findall("m:si", NS):
            shared.append("".join(node.text or "" for node in item.iterfind(".//m:t", NS)))

    sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    for row in sheet.findall(".//m:row", NS):
        row_number = int(row.attrib["r"])
        if row_number > int(sys.argv[2]):
            break
        values = {}
        for cell in row.findall("m:c", NS):
            value = cell.find("m:v", NS)
            if value is None:
                continue
            text = value.text or ""
            if cell.attrib.get("t") == "s":
                text = shared[int(text)]
            values[col_number(cell.attrib["r"])] = text
        if values:
            print(row_number, [values.get(col, "") for col in range(1, max(values) + 1)])
