"""Render portfolio preview images from the real scraper output.

Reads output/books.json and writes PNG images (1600x1200) to assets/.
Every number and row in the images comes from the scraped data.
"""
from __future__ import annotations

import json
import textwrap
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch

OUT = Path("assets")
NAVY, BLUE, LIGHT, INK, MUTED = "#1F3A5F", "#3B82C4", "#EEF3F9", "#1E293B", "#64748B"

available = {f.name for f in font_manager.fontManager.ttflist}
# Font fallback list: Segoe UI for text, DejaVu Sans for glyphs like ★ that Segoe UI lacks
plt.rcParams["font.family"] = [f for f in ("Segoe UI", "DejaVu Sans") if f in available] or ["sans-serif"]


def load() -> pd.DataFrame:
    return pd.DataFrame(json.loads(Path("output/books.json").read_text(encoding="utf-8")))


def new_canvas():
    fig = plt.figure(figsize=(16, 12), dpi=100)
    fig.patch.set_facecolor("white")
    return fig


def header(fig, title: str, subtitle: str) -> None:
    band = fig.add_axes([0, 0.86, 1, 0.14])
    band.set_axis_off()
    band.add_patch(plt.Rectangle((0, 0), 1, 1, color=NAVY, transform=band.transAxes))
    band.text(0.04, 0.60, title, color="white", fontsize=40, fontweight="bold", va="center")
    band.text(0.04, 0.24, subtitle, color="#CBD5E1", fontsize=19, va="center")


def kpi_row(fig, items: list[tuple[str, str]], y: float) -> None:
    width, gap, left = 0.205, 0.02, 0.04
    for i, (value, label) in enumerate(items):
        ax = fig.add_axes([left + i * (width + gap), y, width, 0.11])
        ax.set_axis_off()
        ax.add_patch(FancyBboxPatch((0.02, 0.04), 0.96, 0.92, boxstyle="round,pad=0,rounding_size=0.08",
                                    facecolor=LIGHT, edgecolor="none", transform=ax.transAxes))
        ax.text(0.5, 0.62, value, ha="center", va="center", fontsize=30, fontweight="bold", color=NAVY)
        ax.text(0.5, 0.25, label, ha="center", va="center", fontsize=14, color=MUTED)


def table(fig, rect, df: pd.DataFrame, col_widths: list[float], font_size: int = 13, row_scale: float = 2.0) -> None:
    ax = fig.add_axes(rect)
    ax.set_axis_off()
    tbl = ax.table(cellText=df.values, colLabels=list(df.columns), colWidths=col_widths,
                   cellLoc="left", colLoc="left", loc="upper center")
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(font_size)
    tbl.scale(1, row_scale)
    for (row, col), cell in tbl.get_celld().items():
        cell.PAD = min(0.1, 0.012 / col_widths[col])  # same small left padding in every column
        cell.set_edgecolor("#D9E2EC")
        if row == 0:
            cell.set_facecolor(NAVY)
            cell.set_text_props(color="white", fontweight="bold")
        else:
            cell.set_facecolor(LIGHT if row % 2 == 0 else "white")
            cell.set_text_props(color=INK)


def short(text: str, width: int) -> str:
    return textwrap.shorten(text, width=width, placeholder="…")


def preview_rows(df: pd.DataFrame, n: int, title_width: int) -> pd.DataFrame:
    sample = df.head(n)
    return pd.DataFrame({
        "Title": sample["title"].map(lambda t: short(t, title_width)),
        "Category": sample["category"],
        "Price": sample["price_gbp"].map(lambda p: f"£{p:,.2f}"),
        "Rating": sample["rating"].map(lambda r: "★" * r + "☆" * (5 - r)),
        "In stock": sample["stock_count"].map(lambda s: f"{s} pcs"),
    })


def cover(df: pd.DataFrame, checks_passed: str) -> Path:
    fig = new_canvas()
    header(fig, "Book Catalogue Scraper", "Python web scraping  →  Excel · CSV · JSON")
    kpi_row(fig, [
        (f"{len(df):,}", "books collected"),
        (f"{df['category'].nunique()}", "categories"),
        (f"{len(df.columns)}", "fields per book"),
        (checks_passed, "data checks passed"),
    ], y=0.70)
    fig.text(0.04, 0.655, "Sample of the collected data", fontsize=18, fontweight="bold", color=INK)
    table(fig, [0.04, 0.08, 0.92, 0.56], preview_rows(df, 14, 46), [0.42, 0.18, 0.1, 0.16, 0.14], 15, 2.45)
    fig.text(0.04, 0.035, "Demo project · source: books.toscrape.com (a public sandbox site for scraping practice)",
             fontsize=13, color=MUTED)
    return save(fig, "1-cover.png")


def data_sample(df: pd.DataFrame) -> Path:
    fig = new_canvas()
    header(fig, "Output: Books sheet", "Every product page parsed into clean, typed columns")
    sample = df.iloc[100:120]
    rows = pd.DataFrame({
        "UPC": sample["upc"],
        "Title": sample["title"].map(lambda t: short(t, 34)),
        "Category": sample["category"].map(lambda c: short(c, 16)),
        "Price": sample["price_gbp"].map(lambda p: f"£{p:,.2f}"),
        "Rating": sample["rating"],
        "Stock": sample["stock_count"],
        "Reviews": sample["reviews"],
    })
    table(fig, [0.04, 0.10, 0.92, 0.72], rows, [0.17, 0.33, 0.16, 0.09, 0.08, 0.08, 0.09], 14, 2.3)
    fig.text(0.04, 0.05, "Also exported: product URL, image URL, price excl. tax, tax, availability, description",
             fontsize=14, color=MUTED)
    return save(fig, "2-data-sample.png")


def insights(df: pd.DataFrame) -> Path:
    fig = new_canvas()
    header(fig, "What the data shows", "Quick analysis of the scraped catalogue · per-category stats are also in the Excel file")

    by_cat = (df.groupby("category").agg(books=("upc", "count"), avg_price=("price_gbp", "mean"))
              .sort_values("books", ascending=False).head(12).iloc[::-1])
    ax1 = fig.add_axes([0.17, 0.10, 0.33, 0.68])
    ax1.barh(by_cat.index, by_cat["books"], color=BLUE)
    ax1.set_title("Books per category (top 12)", fontsize=17, fontweight="bold", color=INK, loc="left", pad=14)
    for y, v in enumerate(by_cat["books"]):
        ax1.text(v + 1.5, y, str(v), va="center", fontsize=12, color=INK)

    ratings = df["rating"].value_counts().sort_index()
    ax2 = fig.add_axes([0.60, 0.50, 0.36, 0.28])
    ax2.bar([f"{r}★" for r in ratings.index], ratings.values, color=NAVY, width=0.6)
    ax2.set_title("Books by rating", fontsize=17, fontweight="bold", color=INK, loc="left", pad=14)
    for x, v in enumerate(ratings.values):
        ax2.text(x, v + 4, str(v), ha="center", fontsize=12, color=INK)

    ax3 = fig.add_axes([0.60, 0.10, 0.36, 0.28])
    ax3.hist(df["price_gbp"], bins=20, color=BLUE, edgecolor="white")
    ax3.set_title("Price distribution (£)", fontsize=17, fontweight="bold", color=INK, loc="left", pad=14)

    for ax in (ax1, ax2, ax3):
        ax.spines[["top", "right"]].set_visible(False)
        ax.spines[["left", "bottom"]].set_color("#CBD5E1")
        ax.tick_params(colors=INK, labelsize=12)
    fig.text(0.04, 0.03, "“Default” and “Add a comment” are category names used by the sandbox site itself.",
             fontsize=13, color=MUTED)
    return save(fig, "3-insights.png")


def save(fig, name: str) -> Path:
    OUT.mkdir(exist_ok=True)
    path = OUT / name
    fig.savefig(path, dpi=100, facecolor="white")
    plt.close(fig)
    return path


if __name__ == "__main__":
    from validate import run_checks  # re-runs the data checks so the cover shows the real result

    results = run_checks(Path("output"))
    passed = f"{sum(ok for _, ok, _ in results)} / {len(results)}"
    data = load()
    for p in (cover(data, passed), data_sample(data), insights(data)):
        print("Saved", p)
