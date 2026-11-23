"""Book catalogue scraper for books.toscrape.com.

books.toscrape.com is a public sandbox website built for practising web
scraping. This script walks the whole paginated catalogue, opens every
product page and exports the collected data to Excel, CSV and JSON.

Usage:
    python scraper.py                  # scrape the full catalogue
    python scraper.py --limit 40       # quick test run on the first 40 books
    python scraper.py --workers 4 --out results
"""
from __future__ import annotations

import argparse
import logging
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from exporters import export_all

BASE_URL = "https://books.toscrape.com/"
START_URL = urljoin(BASE_URL, "catalogue/page-1.html")
RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

log = logging.getLogger("scraper")


@dataclass
class Book:
    upc: str
    title: str
    category: str
    price_gbp: float
    price_excl_tax_gbp: float
    tax_gbp: float
    rating: int
    in_stock: bool
    stock_count: int
    reviews: int
    url: str
    image_url: str
    description: str


def make_session(retries: int = 3) -> requests.Session:
    """HTTP session with connection pooling and automatic retries on server errors."""
    session = requests.Session()
    retry = Retry(
        total=retries,
        backoff_factor=0.5,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET",),
    )
    adapter = HTTPAdapter(max_retries=retry, pool_maxsize=16)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers["User-Agent"] = "Mozilla/5.0 (compatible; portfolio-demo-scraper)"
    return session


def get_soup(session: requests.Session, url: str) -> BeautifulSoup:
    resp = session.get(url, timeout=20)
    resp.raise_for_status()
    # The server does not send a charset header, so requests would guess
    # ISO-8859-1 and turn "£" into "Â£". Decode the raw bytes as UTF-8 instead.
    return BeautifulSoup(resp.content, "lxml", from_encoding="utf-8")


def parse_money(text: str) -> float:
    """'£51.77' -> 51.77"""
    return float(re.sub(r"[^\d.]", "", text))


# Broken encoding found in the source text (UTF-8 punctuation read as cp1252)
MOJIBAKE = {"â€™": "’", "â€˜": "‘", "â€œ": "“", "â€\x9d": "”", "â€”": "—", "â€“": "–", "â€¦": "…", 'â€"': "—"}
INVISIBLE = dict.fromkeys(map(ord, "﻿​\xad"))  # BOM, zero-width space, soft hyphen


def clean_text(text: str) -> str:
    """Remove invisible characters, repair known mojibake, collapse whitespace."""
    text = text.translate(INVISIBLE)
    for broken, fixed in MOJIBAKE.items():
        text = text.replace(broken, fixed)
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()


def clean_description(text: str) -> str:
    """Fix quirks in the site's own HTML.

    1. Descriptions start with a cut-off teaser followed by the full text:
       "It's hard to ... love th It's hard to ... love that Silverstein..."
       When the text restarts from its own beginning, the teaser is dropped.
       The teaser and the full text differ slightly in spacing, so the
       restart is confirmed on the first 100 non-space characters.
    2. Some descriptions end with a "...more" link label.
    """
    text = clean_text(text)
    head = text[:40]
    repeat_at = text.find(head, 1) if len(head) == 40 else -1
    if repeat_at > 0:
        teaser = re.sub(r"\s", "", text[:repeat_at])[:100]
        rest = re.sub(r"\s", "", text[repeat_at:])[:100]
        if len(rest) == 100 and rest == teaser:
            text = text[repeat_at:]
    return re.sub(r"\s*\.\.\.more$", "", text)


def collect_book_urls(session: requests.Session, limit: int | None = None) -> tuple[list[str], int | None]:
    """Walk the paginated catalogue.

    Returns the product page URLs and the total number of books the site
    says it has (used later to check that nothing was missed).
    """
    urls: list[str] = []
    expected_total: int | None = None
    page_url: str | None = START_URL

    page_no = 0
    while page_url:
        page_no += 1
        soup = get_soup(session, page_url)
        if expected_total is None:
            counter = soup.select_one("form.form-horizontal")
            match = re.search(r"(\d+)\s+results", counter.get_text()) if counter else None
            expected_total = int(match.group(1)) if match else None

        for link in soup.select("article.product_pod h3 a"):
            urls.append(urljoin(page_url, link["href"]))

        if limit and len(urls) >= limit:
            return urls[:limit], expected_total

        next_link = soup.select_one("li.next a")
        page_url = urljoin(page_url, next_link["href"]) if next_link else None
        if page_no % 10 == 0 or page_url is None:
            log.info("Listing pages read: %d, book links found: %d", page_no, len(urls))

    return urls, expected_total


def parse_book(session: requests.Session, url: str) -> Book:
    soup = get_soup(session, url)

    info = {
        row.th.get_text(strip=True): row.td.get_text(strip=True)
        for row in soup.select("table.table-striped tr")
    }
    availability = info.get("Availability", "")
    stock_match = re.search(r"\((\d+) available\)", availability)

    # Breadcrumb links: Home > Books > <Category>
    category = soup.select("ul.breadcrumb li a")[-1].get_text(strip=True)

    rating_tag = soup.select_one("p.star-rating")
    rating = next((RATING_WORDS[c] for c in rating_tag.get("class", []) if c in RATING_WORDS), 0)

    description_tag = soup.select_one("#product_description + p")
    image_tag = soup.select_one("#product_gallery img")

    return Book(
        upc=info["UPC"],
        title=clean_text(soup.select_one("div.product_main h1").get_text()),
        category=category,
        price_gbp=parse_money(info["Price (incl. tax)"]),
        price_excl_tax_gbp=parse_money(info["Price (excl. tax)"]),
        tax_gbp=parse_money(info["Tax"]),
        rating=rating,
        in_stock=availability.startswith("In stock"),
        stock_count=int(stock_match.group(1)) if stock_match else 0,
        reviews=int(info.get("Number of reviews", "0") or 0),
        url=url,
        image_url=urljoin(url, image_tag["src"]) if image_tag else "",
        description=clean_description(description_tag.get_text()) if description_tag else "",
    )


def scrape(limit: int | None, workers: int) -> tuple[list[Book], int | None, list[str]]:
    session = make_session()
    urls, expected_total = collect_book_urls(session, limit)
    log.info("Found %d book pages, downloading details with %d workers...", len(urls), workers)

    results: dict[int, Book] = {}
    failed: list[str] = []
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(parse_book, session, url): i for i, url in enumerate(urls)}
        for done, future in enumerate(as_completed(futures), start=1):
            index = futures[future]
            try:
                results[index] = future.result()
            except Exception as exc:  # keep going, report failures at the end
                failed.append(urls[index])
                log.warning("Failed %s: %s", urls[index], exc)
            if done % 100 == 0 or done == len(urls):
                log.info("Book pages processed: %d/%d", done, len(urls))

    # Keep the same order as the site's catalogue
    books = [results[i] for i in sorted(results)]
    return books, expected_total, failed


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(encoding="utf-8")  # Windows consoles default to a legacy code page

    parser = argparse.ArgumentParser(description="Scrape books.toscrape.com into Excel, CSV and JSON.")
    parser.add_argument("--limit", type=int, default=None, help="only scrape the first N books (for testing)")
    parser.add_argument("--workers", type=int, default=8, help="parallel downloads (default: 8)")
    parser.add_argument("--out", default="output", help="output folder (default: output)")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)-7s %(message)s", datefmt="%H:%M:%S")

    started = time.perf_counter()
    books, expected_total, failed = scrape(args.limit, args.workers)
    files = export_all(books, Path(args.out), source_url=BASE_URL)
    elapsed = time.perf_counter() - started

    log.info("Done in %.1f s: %d books scraped (site reports %s in total), %d failed",
             elapsed, len(books), expected_total, len(failed))
    for path in files:
        log.info("Saved %s", path)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
