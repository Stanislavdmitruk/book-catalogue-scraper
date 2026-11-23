"""Export scraped books to CSV, JSON and a formatted Excel workbook."""
from __future__ import annotations

import csv
import json
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import TYPE_CHECKING

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

if TYPE_CHECKING:
    from scraper import Book

# column key -> (Excel header, column width, number format)
COLUMNS = {
    "upc": ("UPC", 18, None),
    "title": ("Title", 45, None),
    "category": ("Category", 18, None),
    "price_gbp": ("Price (£)", 11, "£#,##0.00"),
    "price_excl_tax_gbp": ("Price excl. tax (£)", 18, "£#,##0.00"),
    "tax_gbp": ("Tax (£)", 9, "£#,##0.00"),
    "rating": ("Rating (1-5)", 13, "0"),
    "in_stock": ("In stock", 10, None),
    "stock_count": ("Stock qty", 11, "0"),
    "reviews": ("Reviews", 10, "0"),
    "url": ("Product URL", 40, None),
    "image_url": ("Image URL", 40, None),
    "description": ("Description", 60, None),
}

HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14, color="1F3A5F")


def export_all(books: list["Book"], out_dir: Path, source_url: str) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    rows = [asdict(b) for b in books]
    scraped_at = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    return [
        to_csv(rows, out_dir / "books.csv"),
        to_json(rows, out_dir / "books.json"),
        to_xlsx(rows, out_dir / "books.xlsx", source_url, scraped_at),
    ]


def to_csv(rows: list[dict], path: Path) -> Path:
    keys = list(COLUMNS)
    # utf-8-sig adds a BOM so Excel opens "£" and accented titles correctly
    with path.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=keys)
        writer.writeheader()
        writer.writerows(rows)
    return path


def to_json(rows: list[dict], path: Path) -> Path:
    path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def to_xlsx(rows: list[dict], path: Path, source_url: str, scraped_at: str) -> Path:
    df = pd.DataFrame(rows, columns=list(COLUMNS))
    wb = Workbook()

    _summary_sheet(wb.active, df, source_url, scraped_at)
    _books_sheet(wb.create_sheet("Books"), rows)
    _category_sheet(wb.create_sheet("By Category"), df)

    wb.save(path)
    return path


def _style_header(ws, row: int, n_cols: int) -> None:
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _summary_sheet(ws, df: pd.DataFrame, source_url: str, scraped_at: str) -> None:
    ws.title = "Summary"
    ws["A1"] = "Book catalogue: scrape summary"
    ws["A1"].font = TITLE_FONT

    stock_value = float((df["price_gbp"] * df["stock_count"]).sum()) if len(df) else 0.0
    stats = [
        ("Source", source_url),
        ("Scraped at", scraped_at),
        ("Books collected", len(df)),
        ("Categories", df["category"].nunique()),
        ("Average price (£)", round(float(df["price_gbp"].mean()), 2) if len(df) else 0),
        ("Cheapest book (£)", float(df["price_gbp"].min()) if len(df) else 0),
        ("Most expensive book (£)", float(df["price_gbp"].max()) if len(df) else 0),
        ("Units in stock", int(df["stock_count"].sum())),
        ("Stock value (£)", round(stock_value, 2)),
        ("5-star books", int((df["rating"] == 5).sum())),
    ]
    for i, (label, value) in enumerate(stats, start=3):
        ws.cell(row=i, column=1, value=label).font = Font(bold=True)
        cell = ws.cell(row=i, column=2, value=value)
        if "(£)" in label:
            cell.number_format = "£#,##0.00"
        elif isinstance(value, int):
            cell.number_format = "#,##0"
        cell.alignment = Alignment(horizontal="left")
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 34


def _books_sheet(ws, rows: list[dict]) -> None:
    headers = [spec[0] for spec in COLUMNS.values()]
    ws.append(headers)
    for row in rows:
        ws.append([("Yes" if row[k] else "No") if k == "in_stock" else row[k] for k in COLUMNS])

    n_rows, n_cols = len(rows) + 1, len(headers)
    _style_header(ws, 1, n_cols)
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "C2"  # keep header row and UPC/Title visible while scrolling

    for idx, (key, (_, width, fmt)) in enumerate(COLUMNS.items(), start=1):
        letter = get_column_letter(idx)
        ws.column_dimensions[letter].width = width
        if fmt:
            for (cell,) in ws.iter_rows(min_row=2, max_row=n_rows, min_col=idx, max_col=idx):
                cell.number_format = fmt
        if key in ("url", "image_url"):
            for (cell,) in ws.iter_rows(min_row=2, max_row=n_rows, min_col=idx, max_col=idx):
                if cell.value:
                    cell.hyperlink = cell.value
                    cell.font = Font(color="0563C1", underline="single")

    if n_rows > 1:
        ref = f"A1:{get_column_letter(n_cols)}{n_rows}"
        table = Table(displayName="Books", ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(table)

        price_col = get_column_letter(list(COLUMNS).index("price_gbp") + 1)
        stock_col = get_column_letter(list(COLUMNS).index("stock_count") + 1)
        ws.conditional_formatting.add(
            f"{price_col}2:{price_col}{n_rows}",
            ColorScaleRule(start_type="min", start_color="E2EFDA", end_type="max", end_color="F8CBAD"),
        )
        ws.conditional_formatting.add(
            f"{stock_col}2:{stock_col}{n_rows}",
            DataBarRule(start_type="min", end_type="max", color="5B9BD5"),
        )


def _category_sheet(ws, df: pd.DataFrame) -> None:
    summary = (
        df.groupby("category")
        .agg(
            books=("upc", "count"),
            avg_price=("price_gbp", "mean"),
            min_price=("price_gbp", "min"),
            max_price=("price_gbp", "max"),
            units_in_stock=("stock_count", "sum"),
            avg_rating=("rating", "mean"),
        )
        .sort_values(["books", "avg_price"], ascending=[False, False])
        .reset_index()
    )

    headers = ["Category", "Books", "Avg price (£)", "Min price (£)", "Max price (£)", "Units in stock", "Avg rating"]
    formats = [None, "0", "£#,##0.00", "£#,##0.00", "£#,##0.00", "#,##0", "0.00"]
    widths = [22, 9, 14, 14, 14, 15, 11]
    ws.append(headers)
    for record in summary.to_dict("records"):  # native Python types, not numpy
        ws.append(list(record.values()))
    _style_header(ws, 1, len(headers))
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    n_rows = len(summary) + 1
    for idx, (fmt, width) in enumerate(zip(formats, widths), start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
        if fmt:
            for (cell,) in ws.iter_rows(min_row=2, max_row=n_rows, min_col=idx, max_col=idx):
                cell.number_format = fmt

    if n_rows > 1:
        top = min(10, n_rows - 1)
        chart = BarChart()
        chart.type = "bar"
        chart.title = f"Top {top} categories by number of books"
        chart.y_axis.title = "Books"
        chart.legend = None
        chart.height, chart.width = 9, 18
        data = Reference(ws, min_col=2, min_row=1, max_row=top + 1)
        labels = Reference(ws, min_col=1, min_row=2, max_row=top + 1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(labels)
        chart.x_axis.scaling.orientation = "maxMin"  # biggest category on top
        chart.x_axis.delete = False  # show both axes explicitly
        chart.y_axis.delete = False
        ws.add_chart(chart, "I2")

