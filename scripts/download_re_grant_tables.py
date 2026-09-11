#!/usr/bin/env python3
"""Download Research England provider grant-table ZIP files from UKRI indexes."""

from concurrent.futures import ThreadPoolExecutor, as_completed
from html.parser import HTMLParser
from pathlib import Path
import csv
import urllib.request
import zipfile


ROOT = Path(__file__).resolve().parents[1]
SOURCES = ROOT / "data_sources" / "re_provider_grant_tables"
INDEXES = {
    "2023-24": ROOT / "tmp" / "ukri-2023-24.html",
    "2024-25": ROOT / "tmp" / "ukri-2024-25.html",
}


class LinkParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.links = []

    def handle_starttag(self, tag, attrs):
        if tag == "a":
            href = dict(attrs).get("href", "")
            if "funding.re.ukri.org/files/html/grant-data-tables-" in href:
                self.links.append(href.rstrip("/"))


def download(item):
    year, page_url = item
    slug = page_url.rsplit("/", 1)[-1]
    download_url = page_url.replace("/files/html/", "/files/download/") + "/grant-data-tables"
    target_dir = SOURCES / year
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / f"{slug}.zip"
    if not target.exists():
        with urllib.request.urlopen(download_url, timeout=60) as response:
            target.write_bytes(response.read())
    if not zipfile.is_zipfile(target):
        raise ValueError(f"Invalid ZIP: {target}")
    return year, slug, download_url, target.stat().st_size


def main():
    items = []
    for year, path in INDEXES.items():
        parser = LinkParser()
        parser.feed(path.read_text(encoding="utf-8"))
        items.extend((year, link) for link in sorted(set(parser.links)))

    records = []
    with ThreadPoolExecutor(max_workers=8) as pool:
        futures = [pool.submit(download, item) for item in items]
        for future in as_completed(futures):
            records.append(future.result())

    SOURCES.mkdir(parents=True, exist_ok=True)
    with (SOURCES / "manifest.csv").open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["academic_year", "provider_slug", "download_url", "bytes"])
        writer.writerows(sorted(records))

    for year in INDEXES:
        count = sum(record[0] == year for record in records)
        size = sum(record[3] for record in records if record[0] == year)
        print(f"{year}: {count} ZIP files, {size:,} bytes")


if __name__ == "__main__":
    main()
