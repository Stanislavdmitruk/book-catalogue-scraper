"""Data quality checks for the scraper output.

Compares the result with what the website itself reports and checks the
files for duplicates, empty fields, wrong types and broken characters.

Usage:
    python validate.py            # checks output/books.json and output/books.xlsx
    python validate.py --out results
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from pathlib import Path

from openpyxl import load_workbook

from scraper import START_URL, get_soup, make_session

REQUIRED = ["upc", "title", "category", "url", "image_url"]


def site_total() -> int:
    soup = get_soup(make_session(), START_URL)
    return int(re.search(r"(\d+)\s+results", soup.select_one("form.form-horizontal").get_text()).group(1))


def run_checks(out_dir: Path) -> list[tuple[str, bool, str]]:
    books = json.loads((out_dir / "books.json").read_text(encoding="utf-8"))
    checks: list[tuple[str, bool, str]] = []

    def check(name: str, ok: bool, detail: str) -> None:
        checks.append((name, ok, detail))

    expected = site_total()
    check("Row count matches the site", len(books) == expected, f"{len(books)} scraped, site reports {expected}")

    upcs = [b["upc"] for b in books]
    check("No duplicate UPCs", len(upcs) == len(set(upcs)), f"{len(upcs) - len(set(upcs))} duplicates")

    empty = [(b["upc"], k) for b in books for k in REQUIRED if not str(b.get(k, "")).strip()]
    check("Required fields filled", not empty, f"{len(empty)} empty values")

    bad_price = [b["upc"] for b in books if not isinstance(b["price_gbp"], (int, float)) or b["price_gbp"] <= 0]
    check("Prices are positive numbers", not bad_price, f"{len(bad_price)} bad prices")

    bad_rating = [b["upc"] for b in books if b["rating"] not in range(1, 6)]
    check("Ratings are 1-5", not bad_rating, f"{len(bad_rating)} out of range")

    broken = ("Â", "â€", "�", "﻿", "\xad")
    mojibake = [b["upc"] for b in books if any(s in str(v) for v in b.values() for s in broken)]
    check("No broken or invisible characters", not mojibake, f"{len(mojibake)} records affected")

    def has_teaser(d: str) -> bool:
        # a duplicated teaser means the text's first 80 characters appear again later
        return len(d) >= 80 and d.find(d[:80], 1) > 0

    teasers = [b["upc"] for b in books if has_teaser(b["description"]) or b["description"].endswith("...more")]
    check("Descriptions cleaned (no duplicated teaser)", not teasers, f"{len(teasers)} records affected")

    wb = load_workbook(out_dir / "books.xlsx", read_only=True)
    xlsx_rows = wb["Books"].max_row - 1
    check("Excel sheets present", wb.sheetnames == ["Summary", "Books", "By Category"], ", ".join(wb.sheetnames))
    check("Excel row count matches JSON", xlsx_rows == len(books), f"{xlsx_rows} rows in Excel")

    with (out_dir / "books.csv").open(encoding="utf-8-sig", newline="") as fh:
        csv_rows = sum(1 for _ in csv.DictReader(fh))
    check("CSV row count matches JSON", csv_rows == len(books), f"{csv_rows} rows in CSV")

    categories = {b["category"] for b in books}
    check("Categories found", len(categories) > 1, f"{len(categories)} categories")
    return checks


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Validate scraper output.")
    parser.add_argument("--out", default="output", help="folder with books.json / books.xlsx")
    args = parser.parse_args()

    checks = run_checks(Path(args.out))
    for name, ok, detail in checks:
        print(f"[{'PASS' if ok else 'FAIL'}] {name}: {detail}")
    failed = sum(not ok for _, ok, _ in checks)
    print(f"\n{len(checks) - failed}/{len(checks)} checks passed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
