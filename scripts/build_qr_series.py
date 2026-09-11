#!/usr/bin/env python3
"""Reconstruct QR funding shares by broad discipline, 2010/11 to 2024/25."""

from collections import defaultdict
from pathlib import Path
import csv
import io
import json
import re
import sys
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parent))
from xlsx_reader import read_xlsx

import xlrd


ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data_sources"
OUT = ROOT / "tmp" / "qr_series"
YEARS = [f"{year}-{str(year + 1)[-2:]}" for year in range(2010, 2025)]
GROUPS = ["Arts and Humanities", "Social Sciences", "STEM", "Health"]
FIELDS = ["research_councils", "charities", "central_government", "industry", "overseas",
          "qr_charity_basis", "qr_business_basis"]


def num(value):
    if value in (None, ""):
        return 0.0
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return 0.0


def id_text(value):
    if value in (None, ""):
        return ""
    try:
        return str(int(float(value)))
    except ValueError:
        return re.sub(r"\s+", "", str(value))


def old_cost_centre_group(code):
    code = int(code)
    if code in {1, 2, 4, 5, 6, 7, 8}:
        return "Health"
    if code in {3, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20, 21, 23, 24, 25}:
        return "STEM"
    if code in {26, 27, 28, 29, 34, 38, 41}:
        return "Social Sciences"
    if code in {30, 31, 33, 35, 37}:
        return "Arts and Humanities"
    raise ValueError(f"Unmapped old cost centre {code}")


def new_cost_centre_group(code):
    code = int(code)
    if 101 <= code <= 107:
        return "Health"
    if 109 <= code <= 123:
        return "STEM"
    if code in {108, 124, *range(127, 137)}:
        return "Social Sciences"
    if code in {125, 126, *range(137, 146)}:
        return "Arts and Humanities"
    raise ValueError(f"Unmapped new cost centre {code}")


def rae2008_group(uoa):
    uoa = int(uoa)
    if uoa <= 13 or uoa == 15:
        return "Health"
    if uoa == 14 or 16 <= uoa <= 32:
        return "STEM"
    if 34 <= uoa <= 46:
        return "Social Sciences"
    if uoa == 33 or 47 <= uoa <= 67:
        return "Arts and Humanities"
    raise ValueError(f"Unmapped RAE2008 UOA {uoa}")


def ref2014_group(uoa):
    uoa = int(uoa)
    if 1 <= uoa <= 4:
        return "Health"
    if 5 <= uoa <= 16:
        return "STEM"
    if 17 <= uoa <= 26:
        return "Social Sciences"
    if 27 <= uoa <= 36:
        return "Arts and Humanities"
    raise ValueError(f"Unmapped REF2014 UOA {uoa}")


def ref2021_group(uoa):
    # Provider grant tables retain REF sub-panel suffixes (for example 34A).
    # The broad-discipline classification is determined by the numeric UOA.
    match = re.match(r"\d+", str(uoa).strip())
    if not match:
        raise ValueError(f"Unmapped REF2021 UOA {uoa}")
    uoa = int(match.group())
    if 1 <= uoa <= 4:
        return "Health"
    if 5 <= uoa <= 13:
        return "STEM"
    if uoa in {14, *range(16, 25)}:
        return "Social Sciences"
    if uoa == 15 or 25 <= uoa <= 34:
        return "Arts and Humanities"
    raise ValueError(f"Unmapped REF2021 UOA {uoa}")


class HesaData:
    def __init__(self):
        self.data = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
        self.alias = defaultdict(dict)

    def add_provider(self, year, ukprn, instid, group, values):
        key = id_text(ukprn) or f"inst:{id_text(instid)}"
        if ukprn:
            self.alias[year][id_text(ukprn)] = key
        if instid:
            inst = id_text(instid).zfill(4)[-4:]
            self.alias[year][inst] = key
            self.alias[year][f"H-{inst}"] = key
        for field, value in values.items():
            self.data[year][key][group, field] += value

    def provider_key(self, year, raw):
        text = id_text(raw)
        if text.startswith("H-"):
            text = text[2:].zfill(4)[-4:]
        return self.alias[year].get(text, text)

    def provider_group(self, year, raw, group, field):
        return self.data[year][self.provider_key(year, raw)].get((group, field), 0.0)

    def national(self, year, group, field):
        return sum(values.get((group, field), 0.0) for values in self.data[year].values())


def source_values_old(values, year):
    if year == "2014-15":
        # From 2014/15 HESA inserted UK government R&D tax credits as source 5.
        return {
            "research_councils": values[0],
            "charities": values[1] + values[2],
            "central_government": values[3] + values[4],
            "industry": values[5],
            "overseas": sum(values[6:13]),
            "qr_charity_basis": values[1] + values[7] + values[10],
            "qr_business_basis": values[5] + values[8] + values[11],
        }
    return {
        "research_councils": values[0],
        "charities": values[1] + values[2],
        "central_government": values[3],
        "industry": values[4],
        "overseas": sum(values[5:12]),
        "qr_charity_basis": values[1] + values[6] + values[9],
        "qr_business_basis": values[4] + values[7] + values[10],
    }


def load_hesa_old(hesa, year):
    archive = zipfile.ZipFile(DATA / f"HE_Finance_Plus_{year}.zip")
    member = next(name for name in archive.namelist() if re.search(r"table_5b\.xlsx?$", name, re.I))
    payload = archive.read(member)
    if member.lower().endswith(".xls"):
        book = xlrd.open_workbook(file_contents=payload)
        sheets = [[[sheet.cell_value(row, col) for col in range(sheet.ncols)] for row in range(sheet.nrows)]
                  for sheet in book.sheets()]
    else:
        temp = OUT / f"hesa-{year}.xlsx"
        temp.parent.mkdir(parents=True, exist_ok=True)
        temp.write_bytes(payload)
        sheets = [read_xlsx(temp)[1]]
    for rows in sheets:
        # The cost-centre labels are on row 4 in the published spreadsheet
        # (zero-based row 3).  Old .xls files span continuation sheets; the
        # final continuation contains only non-academic overhead categories.
        if len(rows) <= 3:
            continue
        cc_row = rows[3]
        starts = []
        for index, value in enumerate(cc_row):
            match = re.match(r"^(\d{2,3}) ", str(value).strip())
            if match and index >= 4 and int(match.group(1)) < 200:
                starts.append((index, int(match.group(1))))
        for row in rows:
            if len(row) < 4 or not re.fullmatch(r"\d{4}", id_text(row[0]).zfill(4)[-4:]):
                continue
            region = str(row[2]).strip()
            if region in {"WALE", "SCOT", "NIRE", ""}:
                continue
            for start, code in starts:
                width = 15 if year == "2014-15" else 14
                block = [num(row[start + offset]) if start + offset < len(row) else 0 for offset in range(width)]
                group = old_cost_centre_group(code) if code < 100 else new_cost_centre_group(code)
                hesa.add_provider(year, row[1], row[0], group, source_values_old(block, year))


def load_hesa_csv(hesa, year):
    archive = zipfile.ZipFile(DATA / "table-5 (3).zip")
    member = f"table-5-({year}).csv"
    handle = io.TextIOWrapper(archive.open(member), encoding="utf-8-sig", newline="")
    for _ in range(10):
        next(handle)
    reader = csv.DictReader(handle)
    staged = defaultdict(lambda: defaultdict(float))
    for row in reader:
        # The downloadable CSV also contains regional and national subtotal
        # rows labelled "Total" with a blank UKPRN.  Retain provider records
        # only, otherwise every income stream is counted more than once.
        if (row["Country of HE provider"] != "England"
                or row["HESA cost centre marker"] != "Academic departments"
                or not row["UKPRN"].strip()
                or row.get("Year End Month", "All") != "All"):
            continue
        code_match = re.match(r"(\d{3}) ", row["HESA cost centre"])
        if not code_match:
            continue
        group = new_cost_centre_group(int(code_match.group(1)))
        source = row["Source of income"]
        value = num(row["Value(£000s)"])
        key = (row["UKPRN"], group)
        if "Total Research Councils" in source:
            staged[key]["research_councils"] += value
        elif source.startswith("2 "):
            staged[key]["charities"] += value
            staged[key]["qr_charity_basis"] += value
        elif source.startswith("3 "):
            staged[key]["charities"] += value
        elif source.startswith("4 ") or source.startswith("5 "):
            staged[key]["central_government"] += value
        elif source.startswith("6 "):
            staged[key]["industry"] += value
            staged[key]["qr_business_basis"] += value
        elif re.match(r"(?:8|9|10|11|12|13|14) ", source):
            staged[key]["overseas"] += value
            if source.startswith("9 ") or source.startswith("12 "):
                staged[key]["qr_charity_basis"] += value
            if source.startswith("10 ") or source.startswith("13 "):
                staged[key]["qr_business_basis"] += value
    for (ukprn, group), values in staged.items():
        hesa.add_provider(year, ukprn, "", group, values)


def load_hesa():
    hesa = HesaData()
    for year in YEARS[:5]:
        load_hesa_old(hesa, year)
    for year in YEARS[5:]:
        load_hesa_csv(hesa, year)
    return hesa


def split_provider_grants(hesa, year, grants, basis):
    result = defaultdict(float)
    unresolved = 0.0
    for provider, amount in grants:
        weights = {group: hesa.provider_group(year, provider, group, basis) for group in GROUPS}
        total = sum(weights.values())
        if total <= 0:
            unresolved += amount
            continue
        for group, weight in weights.items():
            result[group] += amount * weight / total
    # A small number of recipients have no qualifying income in the same-year
    # HESA table (especially after mergers or name/identifier changes).  Keep
    # the published QR control total by distributing this residual in the
    # national subject pattern for the relevant qualifying-income stream.
    if unresolved:
        national = {group: hesa.national(year, group, basis) for group in GROUPS}
        national_total = sum(national.values())
        if national_total:
            for group, weight in national.items():
                result[group] += unresolved * weight / national_total
    return result, unresolved


def old_qr(year, hesa):
    suffix = year[:4][-2:] + year[-2:]
    path = DATA / f"resdata{suffix}.xls"
    sheet = xlrd.open_workbook(path).sheet_by_index(0)
    headers = [str(sheet.cell_value(5, col)) for col in range(sheet.ncols)]
    main_col = next(i for i, h in enumerate(headers) if "Mainstream QR allocation" in h)
    london_col = next(i for i, h in enumerate(headers) if "Allocation for London" in h)
    rdp_col = next(i for i, h in enumerate(headers) if "RDP supervision allocation (£)" in h)
    charity_col = next((i for i, h in enumerate(headers) if "Charity support funding (£)" in h), None)
    result = defaultdict(lambda: defaultdict(float))
    for row in range(6, sheet.nrows):
        group = rae2008_group(sheet.cell_value(row, 2))
        result[group]["mainstream"] += num(sheet.cell_value(row, main_col))
        result[group]["london"] += num(sheet.cell_value(row, london_col))
        result[group]["rdp"] += num(sheet.cell_value(row, rdp_col))
        if charity_col is not None:
            result[group]["charity_qr"] += num(sheet.cell_value(row, charity_col))

    if year in {"2010-11", "2011-12"}:
        grants_path = DATA / f"business{suffix}.xls"
        grant_sheet = xlrd.open_workbook(grants_path).sheet_by_index(0)
        header_row = next(r for r in range(10) if any("Business research element" in str(grant_sheet.cell_value(r, c)) for c in range(grant_sheet.ncols)))
        grants = [(grant_sheet.cell_value(r, 0), num(grant_sheet.cell_value(r, grant_sheet.ncols - 1)))
                  for r in range(header_row + 1, grant_sheet.nrows)]
        business, unresolved = split_provider_grants(hesa, year, grants,
                                                     "industry" if year == "2010-11" else "qr_business_basis")
        for group in GROUPS:
            result[group]["business_qr"] = business[group]
        result["_meta"]["unresolved_business"] = unresolved
    else:
        grants_path = DATA / f"businesscharity{suffix}.xls"
        grant_sheet = xlrd.open_workbook(grants_path).sheet_by_index(0)
        headers = [str(grant_sheet.cell_value(5, col)) for col in range(grant_sheet.ncols)]
        business_col = next(i for i, h in enumerate(headers) if "Business research element (£)" in h)
        charity_col = next(i for i, h in enumerate(headers) if "Charity support fund (£)" in h)
        business_grants = [(grant_sheet.cell_value(r, 0), num(grant_sheet.cell_value(r, business_col))) for r in range(6, grant_sheet.nrows)]
        charity_grants = [(grant_sheet.cell_value(r, 0), num(grant_sheet.cell_value(r, charity_col))) for r in range(6, grant_sheet.nrows)]
        business, ub = split_provider_grants(hesa, year, business_grants, "qr_business_basis")
        charity, uc = split_provider_grants(hesa, year, charity_grants, "qr_charity_basis")
        for group in GROUPS:
            result[group]["business_qr"] = business[group]
            result[group]["charity_qr"] = charity[group]
        result["_meta"]["unresolved_business"] = ub
        result["_meta"]["unresolved_charity"] = uc
    return result


def xlsx_detailed(path, ref_version, component):
    _, rows = read_xlsx(path)
    phrase = "Mainstream QR allocation" if component == "mainstream" else "QR RDP supervision allocation"
    header_index = next(index for index, row in enumerate(rows) if any(phrase in str(value) for value in row))
    header = [str(value) for value in rows[header_index]]
    uoa_col = next(i for i, h in enumerate(header) if "Unit of assessment" in h)
    provider_col = next(i for i, h in enumerate(header) if "Provider Reference" in h or "Provider Reference Number" in h)
    if component == "mainstream":
        value_col = next(i for i, h in enumerate(header) if "Mainstream QR allocation" in h)
        london_col = next(i for i, h in enumerate(header) if "Allocation for London" in h)
    else:
        value_col = next(i for i, h in enumerate(header)
                         if "QR RDP supervision allocation" in h and "£" in h)
        london_col = None
    grouper = ref2014_group if ref_version == 2014 else ref2021_group
    result = defaultdict(lambda: defaultdict(float))
    provider_result = defaultdict(lambda: defaultdict(float))
    uoa_result = defaultdict(lambda: defaultdict(float))
    for row in rows[header_index + 1:]:
        if len(row) <= max(provider_col, uoa_col, value_col) or not id_text(row[provider_col]).isdigit():
            continue
        uoa = int(num(row[uoa_col]))
        group = grouper(uoa)
        provider = id_text(row[provider_col])
        value = num(row[value_col])
        result[group][component] += value
        provider_result[provider][group] += value
        uoa_result[uoa][component] += value
        if london_col is not None:
            london = num(row[london_col])
            result[group]["london"] += london
            provider_result[provider][group] += london
            uoa_result[uoa]["london"] += london
    return result, provider_result, uoa_result


def allocation_grants_xlsx(path, year, sheet_name=None):
    _, rows = read_xlsx(path, sheet_name)
    # Most allocation workbooks put UKPRN and the grant labels on one row.  The
    # 2022/23 annex uses a two-line header: descriptive labels first and field
    # names (including UKPRN) underneath.
    header_index = next(index for index, row in enumerate(rows)
                        if any("qr charity support" in str(value).lower() for value in row))
    header = [str(value) for value in rows[header_index]]
    provider_candidates = [i for i, h in enumerate(header)
                           if h.strip() == "UKPRN" or "Provider Reference" in h]
    data_start = header_index + 1
    if provider_candidates:
        provider_col = provider_candidates[0]
    else:
        field_header = [str(value) for value in rows[header_index + 1]]
        provider_col = next(i for i, h in enumerate(field_header) if h.strip() == "UKPRN")
        data_start = header_index + 2
    charity_col = next(i for i, h in enumerate(header)
                       if h.lower().startswith("qr charity support"))
    business_col = next(i for i, h in enumerate(header)
                        if h.lower().startswith("qr business research"))
    charity = []
    business = []
    for row in rows[data_start:]:
        if len(row) <= max(provider_col, charity_col, business_col) or not id_text(row[provider_col]).isdigit():
            continue
        charity.append((row[provider_col], num(row[charity_col])))
        business.append((row[provider_col], num(row[business_col])))
    return charity, business


def provider_zip_data(year):
    base = DATA / "re_provider_grant_tables" / year
    result = defaultdict(lambda: defaultdict(float))
    rdp_provider = defaultdict(lambda: defaultdict(float))
    charity = []
    business = []
    for path in sorted(base.glob("*.zip")):
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            table_a = next(name for name in names if name.endswith("_TableA.csv"))
            rows = list(csv.DictReader(io.TextIOWrapper(archive.open(table_a), encoding="utf-8-sig")))
            provider = rows[0]["UKPRN"] if rows else ""
            for row in rows:
                if row["Sub-Fund"] == "QR charity support fund":
                    charity.append((provider, num(row["Allocation"])))
                elif row["Sub-Fund"] == "QR business research element":
                    business.append((provider, num(row["Allocation"])))
            table_b = next(name for name in names if name.endswith("_TableB.csv"))
            for row in csv.DictReader(io.TextIOWrapper(archive.open(table_b), encoding="utf-8-sig")):
                group = ref2021_group(row["UOA"])
                result[group]["mainstream"] += num(row["Mainstream QR funds (£)"])
                result[group]["london"] += num(row["London weighting on mainstream QR (£)"])
            if year == "2023-24":
                table_d = next(name for name in names if name.endswith("_TableD.csv"))
                for row in csv.DictReader(io.TextIOWrapper(archive.open(table_d), encoding="utf-8-sig")):
                    if not row.get("UOA"):
                        continue
                    group = ref2021_group(row["UOA"])
                    value_key = next(key for key in row if "RDP" in key and "fund" in key.lower() and "£" in key)
                    value = num(row[value_key])
                    result[group]["rdp"] += value
                    rdp_provider[provider][group] += value
    return result, rdp_provider, charity, business


def london_weights_from_2019():
    _, rows = read_xlsx(DATA / "QR-charity-business-2019-20.xlsx")
    header = [str(value) for value in rows[4]]
    weights = {}
    for row in rows[5:]:
        if len(row) < 8 or not id_text(row[0]).isdigit():
            continue
        average = num(row[6])
        weighted = num(row[7])
        weights[id_text(row[0])] = weighted / average if average else 1.0
    return weights


def missing_year_qr(year, hesa, base_main_uoa, base_rdp_uoa, london_weights):
    pots = {
        "2015-16": (1_050_000_000, 240_000_000, 198_000_000, 64_000_000),
        "2016-17": (1_050_000_000, 240_000_000, 198_000_000, 64_000_000),
        "2017-18": (1_050_000_000, 240_000_000, 198_000_000, 64_000_000),
        "2018-19": (1_050_000_000, 260_000_000, 204_000_000, 64_000_000),
    }
    main_pot, rdp_pot, charity_pot, business_pot = pots[year]
    result = defaultdict(lambda: defaultdict(float))
    for component, base, pot in [("mainstream_london", base_main_uoa, main_pot), ("rdp", base_rdp_uoa, rdp_pot)]:
        adjusted = {}
        for uoa, values in base.items():
            value = values.get("mainstream", 0) + values.get("london", 0) if component == "mainstream_london" else values.get("rdp", 0)
            if uoa == 4:
                value *= 1.42 / 1.6
            adjusted[uoa] = value
        scale = pot / sum(adjusted.values())
        for uoa, value in adjusted.items():
            group = ref2014_group(uoa)
            if component == "mainstream_london":
                cost_adjustment = 1.42 / 1.6 if uoa == 4 else 1.0
                result[group]["mainstream"] += base[uoa].get("mainstream", 0) * cost_adjustment * scale
                result[group]["london"] += base[uoa].get("london", 0) * cost_adjustment * scale
            else:
                result[group]["rdp"] += value * scale

    char_weights = defaultdict(float)
    bus_weights = defaultdict(float)
    for provider, values in hesa.data[year].items():
        london = london_weights.get(provider, 1.0)
        for group in GROUPS:
            char_weights[group] += values.get((group, "qr_charity_basis"), 0) * london
            bus_weights[group] += values.get((group, "qr_business_basis"), 0)
    for group in GROUPS:
        result[group]["charity_qr"] = charity_pot * char_weights[group] / sum(char_weights.values())
        result[group]["business_qr"] = business_pot * bus_weights[group] / sum(bus_weights.values())
    return result


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    hesa = load_hesa()
    qr = {}
    for year in YEARS[:5]:
        qr[year] = old_qr(year, hesa)

    detailed = {}
    for year in ["2019-20", "2020-21", "2021-22", "2022-23"]:
        ref = 2021 if year == "2022-23" else 2014
        main_result, _, main_uoa = xlsx_detailed(DATA / ("QR-mainstream-london-gcrf-2019-20.xlsx" if year == "2019-20" else f"QR-mainstream-london-{year}.xlsx"), ref, "mainstream")
        rdp_result, _, rdp_uoa = xlsx_detailed(DATA / f"QR-rdp-{year}.xlsx", ref, "rdp")
        combined = defaultdict(lambda: defaultdict(float))
        for group in GROUPS:
            combined[group].update(main_result[group])
            combined[group].update(rdp_result[group])
        detailed[year] = (combined, main_uoa, rdp_uoa)

    weights = london_weights_from_2019()
    base_main_uoa = detailed["2019-20"][1]
    base_rdp_uoa = detailed["2019-20"][2]
    for year in YEARS[5:9]:
        qr[year] = missing_year_qr(year, hesa, base_main_uoa, base_rdp_uoa, weights)

    allocation_files = {
        "2019-20": (DATA / "QR-charity-business-2019-20.xlsx", None),
        "2020-21": (DATA / "RE-04102021-ResearchKEF-AnnexA-InstitutionBreakdown-2020-21.xlsx", "Table_1"),
        "2021-22": (DATA / "RE-280322-ResearchKnowledgeExchangeFundingAnnexA-2021-22.xlsx", "Table_1"),
        "2022-23": (DATA / "RE-130323-GrantAllocations2022To2023AnnexA.xlsx", "Table_1"),
    }
    for year, (path, sheet) in allocation_files.items():
        combined = detailed[year][0]
        charity_grants, business_grants = allocation_grants_xlsx(path, year, sheet)
        charity, uc = split_provider_grants(hesa, year, charity_grants, "qr_charity_basis")
        business, ub = split_provider_grants(hesa, year, business_grants, "qr_business_basis")
        for group in GROUPS:
            combined[group]["charity_qr"] = charity[group]
            combined[group]["business_qr"] = business[group]
        combined["_meta"]["unresolved_charity"] = uc
        combined["_meta"]["unresolved_business"] = ub
        qr[year] = combined

    result_2023, rdp_2023, charity_2023, business_2023 = provider_zip_data("2023-24")
    char, uc = split_provider_grants(hesa, "2023-24", charity_2023, "qr_charity_basis")
    bus, ub = split_provider_grants(hesa, "2023-24", business_2023, "qr_business_basis")
    for group in GROUPS:
        result_2023[group]["charity_qr"] = char[group]
        result_2023[group]["business_qr"] = bus[group]
    result_2023["_meta"]["unresolved_charity"] = uc
    result_2023["_meta"]["unresolved_business"] = ub
    qr["2023-24"] = result_2023

    result_2024, _, charity_2024, business_2024 = provider_zip_data("2024-25")
    total_rdp_2023 = {provider: sum(values.values()) for provider, values in rdp_2023.items()}
    table_a_rdp_2024 = {}
    for path in (DATA / "re_provider_grant_tables" / "2024-25").glob("*.zip"):
        with zipfile.ZipFile(path) as archive:
            table_a = next(name for name in archive.namelist() if name.endswith("_TableA.csv"))
            rows = list(csv.DictReader(io.TextIOWrapper(archive.open(table_a), encoding="utf-8-sig")))
            if rows:
                provider = rows[0]["UKPRN"]
                for row in rows:
                    if row["Sub-Fund"] == "QR RDP supervision funds":
                        table_a_rdp_2024[provider] = num(row["Allocation"])
    national_fallback = defaultdict(float)
    for values in rdp_2023.values():
        for group, value in values.items():
            national_fallback[group] += value
    national_total = sum(national_fallback.values())
    for provider, allocation in table_a_rdp_2024.items():
        previous = total_rdp_2023.get(provider, 0)
        for group in GROUPS:
            share = rdp_2023[provider][group] / previous if previous else national_fallback[group] / national_total
            result_2024[group]["rdp"] += allocation * share
    char, uc = split_provider_grants(hesa, "2024-25", charity_2024, "qr_charity_basis")
    bus, ub = split_provider_grants(hesa, "2024-25", business_2024, "qr_business_basis")
    for group in GROUPS:
        result_2024[group]["charity_qr"] = char[group]
        result_2024[group]["business_qr"] = bus[group]
    result_2024["_meta"]["unresolved_charity"] = uc
    result_2024["_meta"]["unresolved_business"] = ub
    qr["2024-25"] = result_2024

    rows = []
    components = []
    rdp_2014_total = sum(qr["2014-15"][group].get("rdp", 0) for group in GROUPS)
    missing_rdp_pots = {"2015-16": 240_000_000, "2016-17": 240_000_000,
                        "2017-18": 240_000_000, "2018-19": 260_000_000}
    for year in YEARS:
        for group in GROUPS:
            external = {field: hesa.national(year, group, field) * 1000 for field in FIELDS[:5]}
            q = qr[year][group]
            mainstream = q.get("mainstream", 0)
            london = q.get("london", 0)
            other_qr = london + q.get("rdp", 0) + q.get("charity_qr", 0) + q.get("business_qr", 0)
            total_qr = mainstream + other_qr
            identifiable = total_qr + sum(external.values())
            if year in missing_rdp_pots:
                alternative_rdp = (missing_rdp_pots[year]
                                   * qr["2014-15"][group].get("rdp", 0) / rdp_2014_total)
                alternative_total_qr = total_qr - q.get("rdp", 0) + alternative_rdp
                alternative_share = alternative_total_qr / (alternative_total_qr + sum(external.values()))
                central_share = total_qr / identifiable
                sensitivity_low = min(central_share, alternative_share)
                sensitivity_high = max(central_share, alternative_share)
            else:
                sensitivity_low = sensitivity_high = total_qr / identifiable if identifiable else ""
            if year in YEARS[5:9]:
                status = "Reconstructed (missing detailed QR workbook)"
            elif year == "2024-25":
                status = "Published QR; RDP subject mix carried forward from 2023-24"
            else:
                status = "Published QR; provider support funds allocated by HESA income"
            rows.append({
                "academic_year": year,
                "discipline": group,
                "qr_share": total_qr / identifiable if identifiable else "",
                "mainstream_qr_share": mainstream / identifiable if identifiable else "",
                "other_qr_share": other_qr / identifiable if identifiable else "",
                "qr_share_sensitivity_low": sensitivity_low,
                "qr_share_sensitivity_high": sensitivity_high,
                "qr_data_status": status,
            })
            component_row = {"academic_year": year, "discipline": group,
                             "mainstream_qr": mainstream, "london_qr": london,
                             "rdp_qr": q.get("rdp", 0), "charity_qr": q.get("charity_qr", 0),
                             "business_qr": q.get("business_qr", 0), "total_qr": total_qr,
                             **external, "identifiable_research_funding": identifiable}
            components.append(component_row)

    with (OUT / "qr_share_series.csv").open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader(); writer.writerows(rows)
    with (OUT / "qr_components.csv").open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(components[0]))
        writer.writeheader(); writer.writerows(components)
    diagnostics = {year: dict(qr[year].get("_meta", {})) for year in YEARS}
    (OUT / "diagnostics.json").write_text(json.dumps(diagnostics, indent=2), encoding="utf-8")
    print(json.dumps({"rows": len(rows), "components": len(components), "2010-11": [row for row in rows if row["academic_year"] == "2010-11"], "diagnostics": diagnostics}, indent=2))


if __name__ == "__main__":
    main()
