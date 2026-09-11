import collections
import re
import zipfile
import xml.etree.ElementTree as ET

import xlrd

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def ref2014_group(uoa):
    uoa = int(uoa)
    if uoa <= 4:
        return "Health"
    if uoa <= 15:
        return "STEM"
    if uoa <= 26:
        return "Social Sciences"
    return "Arts and Humanities"


def rae2008_group(uoa):
    uoa = int(uoa)
    if uoa <= 13 or uoa == 15:
        return "Health"
    if uoa == 14 or 16 <= uoa <= 32:
        return "STEM"
    if 34 <= uoa <= 46:
        return "Social Sciences"
    return "Arts and Humanities"


old = collections.defaultdict(float)
sheet = xlrd.open_workbook("data_sources/resdata1415.xls").sheet_by_name("QR1415data")
for row in range(6, sheet.nrows):
    old[rae2008_group(sheet.cell_value(row, 2))] += float(sheet.cell_value(row, 21) or 0)

with zipfile.ZipFile("data_sources/QR-rdp-2019-20.xlsx") as archive:
    shared = []
    strings = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    for item in strings.findall("m:si", NS):
        shared.append("".join(node.text or "" for node in item.iterfind(".//m:t", NS)))
    root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    new = collections.defaultdict(float)
    for row in root.findall(".//m:row", NS):
        if int(row.attrib["r"]) < 9:
            continue
        values = {}
        for cell in row.findall("m:c", NS):
            ref = cell.attrib["r"]
            col = re.match(r"[A-Z]+", ref).group()
            value = cell.find("m:v", NS)
            if value is not None:
                text = value.text
                if cell.attrib.get("t") == "s":
                    text = shared[int(text)]
                values[col] = text
        if values.get("C") and values.get("P"):
            new[ref2014_group(values["C"])] += float(values["P"])

for label, data in [("2014/15", old), ("2019/20", new)]:
    total = sum(data.values())
    print(label, round(total), {key: round(value / total * 100, 3) for key, value in data.items()})
