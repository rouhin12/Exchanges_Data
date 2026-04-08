#!/usr/bin/env python3
"""Download FII/DII cash-segment data (NSE-style buy/sell/net) and save to Excel.

Sources:
  - Moneycontrol cash market page embeds ~30 trading days in __NEXT_DATA__ (JSON),
    including FII/DII gross purchase, gross sales, and net (Rs crore).
  - NSE India `fiidiiTradeReact` JSON for the latest session (official provisional figures).

Older rows can be kept from the existing workbook, or bulk-loaded from optional
`fii_dii_backfill.csv` (same columns as the Excel sheet) — e.g. an export from
NSE historical FII/FPI & DII reports — then trimmed to the last five years.

The old Moneycontrol HTML table selector (`div.fidi_tbescrol`) no longer matches
the live page (Next.js); scraping that div returns no usable rows.
"""

from __future__ import annotations

import argparse
import json
import re
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import requests

OUTPUT_FILE = Path(__file__).parent / "fii_dii_data.xlsx"
BACKFILL_CSV = Path(__file__).parent / "fii_dii_backfill.csv"

MC_CASH_URL = "https://www.moneycontrol.com/markets/fii-dii-data/cash/"
NSE_FIIDII_URL = "https://www.nseindia.com/api/fiidiiTradeReact"

# Values are stored in Rs crore (server.py divides by 100 for Rs bn in the UI).
YEARS_HISTORY = 5

HEADERS_BROWSER = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-IN,en;q=0.9",
}

HEADERS_NSE_JSON = {
    "User-Agent": HEADERS_BROWSER["User-Agent"],
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://www.nseindia.com/",
    "Origin": "https://www.nseindia.com",
}


def parse_inr(value: str) -> float:
    """Parse numbers like '-11,163.06' or '21,952.95' to float."""
    if value is None:
        return float("nan")
    s = str(value).strip().replace(",", "")
    s = s.replace("−", "-").replace("\u2212", "-")
    if not s or s == "-":
        return float("nan")
    return float(s)


def extract_next_data_json(html: str) -> dict | None:
    m = re.search(
        r'<script id="__NEXT_DATA__"[^>]*type="application/json"[^>]*>([^<]+)</script>',
        html,
    )
    if not m:
        m = re.search(r'<script id="__NEXT_DATA__"[^>]*>([^<]+)</script>', html)
    if not m:
        return None
    try:
        return json.loads(m.group(1))
    except json.JSONDecodeError:
        return None


def fetch_moneycontrol_cash() -> pd.DataFrame | None:
    """Parse FII/DII cash table from Moneycontrol embedded JSON."""
    try:
        r = requests.get(MC_CASH_URL, headers=HEADERS_BROWSER, timeout=30)
        r.raise_for_status()
    except requests.RequestException as e:
        print(f"ERROR: Moneycontrol request failed: {e}")
        return None

    payload = extract_next_data_json(r.text)
    if not payload:
        print("ERROR: __NEXT_DATA__ not found on Moneycontrol cash page")
        return None

    rows_raw = (
        payload.get("props", {})
        .get("pageProps", {})
        .get("FiiDiiData", {})
        .get("fiiDiiData")
    )
    if not rows_raw:
        print("ERROR: No FiiDiiData.fiiDiiData in page JSON")
        return None

    out: list[dict[str, float | object]] = []
    for row in rows_raw:
        if "fiiPurchase" not in row:
            continue
        d = pd.to_datetime(row["date"], errors="coerce")
        if pd.isna(d):
            continue
        out.append(
            {
                "Date": d,
                "FII_Gross_Purchase": parse_inr(row.get("fiiPurchase")),
                "FII_Gross_Sales": parse_inr(row.get("fiiSales")),
                "FII_Net": parse_inr(row.get("fiiNet")),
                "DII_Gross_Purchase": parse_inr(row.get("diiPurchase")),
                "DII_Gross_Sales": parse_inr(row.get("diiSale")),
                "DII_Net": parse_inr(row.get("diiNet")),
            }
        )

    if not out:
        print("ERROR: No cash-segment rows parsed from Moneycontrol JSON")
        return None

    df = pd.DataFrame(out)
    print(f"OK: Moneycontrol cash segment - {len(df)} rows")
    return df


def nse_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS_NSE_JSON)
    s.get("https://www.nseindia.com", timeout=20)
    s.get("https://www.nseindia.com/option-chain", timeout=20)
    return s


def fetch_nse_fiidii() -> pd.DataFrame | None:
    """Latest trading day from NSE (FII/FPI + DII, Rs crore)."""
    try:
        s = nse_session()
        r = s.get(NSE_FIIDII_URL, timeout=25)
        r.raise_for_status()
        data = r.json()
    except (requests.RequestException, ValueError) as e:
        print(f"ERROR: NSE fiidiiTradeReact failed: {e}")
        return None

    if not isinstance(data, list) or len(data) < 2:
        print("ERROR: Unexpected NSE fiidiiTradeReact JSON shape")
        return None

    by_cat = {item.get("category"): item for item in data}
    fii = by_cat.get("FII/FPI") or by_cat.get("FII")
    dii = by_cat.get("DII")
    if not fii or not dii:
        print("ERROR: NSE response missing FII/FPI or DII category")
        return None

    d_str = fii.get("date") or dii.get("date")
    dt = pd.to_datetime(d_str, format="%d-%b-%Y", errors="coerce")
    if pd.isna(dt):
        dt = pd.to_datetime(d_str, errors="coerce")
    if pd.isna(dt):
        print("ERROR: Could not parse NSE trade date")
        return None

    row = {
        "Date": dt,
        "FII_Gross_Purchase": float(fii.get("buyValue", 0) or 0),
        "FII_Gross_Sales": float(fii.get("sellValue", 0) or 0),
        "FII_Net": float(fii.get("netValue", 0) or 0),
        "DII_Gross_Purchase": float(dii.get("buyValue", 0) or 0),
        "DII_Gross_Sales": float(dii.get("sellValue", 0) or 0),
        "DII_Net": float(dii.get("netValue", 0) or 0),
    }
    print(f"OK: NSE fiidiiTradeReact - date {dt.date()}")
    return pd.DataFrame([row])


def load_backfill_csv(path: Path) -> pd.DataFrame | None:
    if not path.exists():
        return None
    try:
        df = pd.read_csv(path)
    except Exception as e:
        print(f"WARNING: Could not read backfill CSV {path}: {e}")
        return None
    required = {
        "Date",
        "FII_Gross_Purchase",
        "FII_Gross_Sales",
        "FII_Net",
        "DII_Gross_Purchase",
        "DII_Gross_Sales",
        "DII_Net",
    }
    if not required.issubset(set(df.columns)):
        print(f"WARNING: Backfill CSV missing columns {required - set(df.columns)}")
        return None
    df = df.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    print(f"OK: Backfill CSV - {len(df)} rows from {path.name}")
    return df


def normalize_frame(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Date"] = pd.to_datetime(df["Date"]).dt.normalize()
    num_cols = [
        "FII_Gross_Purchase",
        "FII_Gross_Sales",
        "FII_Net",
        "DII_Gross_Purchase",
        "DII_Gross_Sales",
        "DII_Net",
    ]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def merge_sources(
    df_mc: pd.DataFrame | None,
    df_nse: pd.DataFrame | None,
    df_backfill: pd.DataFrame | None,
) -> pd.DataFrame | None:
    parts: list[pd.DataFrame] = []
    if df_backfill is not None and not df_backfill.empty:
        parts.append(normalize_frame(df_backfill))
    if df_mc is not None and not df_mc.empty:
        parts.append(normalize_frame(df_mc))
    if df_nse is not None and not df_nse.empty:
        parts.append(normalize_frame(df_nse))
    if not parts:
        return None
    df = pd.concat(parts, ignore_index=True)
    df = df.drop_duplicates(subset=["Date"], keep="last")
    df = df.sort_values("Date")
    return df


def load_existing_excel(path: Path) -> pd.DataFrame | None:
    if not path.exists():
        return None
    try:
        df = pd.read_excel(path)
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df = df.dropna(subset=["Date"])
        return normalize_frame(df)
    except Exception as e:
        print(f"WARNING: Could not read existing Excel: {e}")
        return None


def trim_to_years(df: pd.DataFrame, years: int) -> pd.DataFrame:
    end = pd.Timestamp.now().normalize()
    start = end - timedelta(days=365 * years)
    return df[(df["Date"] >= start) & (df["Date"] <= end)]


def save_to_excel(df: pd.DataFrame, path: Path) -> bool:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="FII_DII", index=False)
        print(f"OK: Saved {path}")
        print(f"   Records: {len(df)}")
        if not df.empty:
            print(f"   Date range: {df['Date'].min().date()} to {df['Date'].max().date()}")
        return True
    except Exception as e:
        print(f"ERROR: Save failed: {e}")
        return False


def run_download(
    *,
    backfill_path: Path | None = None,
    merge_existing: bool = True,
    years: int = YEARS_HISTORY,
) -> pd.DataFrame | None:
    df_mc = fetch_moneycontrol_cash()
    df_nse = fetch_nse_fiidii()
    df_back = load_backfill_csv(backfill_path) if backfill_path else None

    df = merge_sources(df_mc, df_nse, df_back)
    if df is None or df.empty:
        print("ERROR: No data from Moneycontrol, NSE, or backfill CSV")
        return None

    if merge_existing:
        existing = load_existing_excel(OUTPUT_FILE)
        if existing is not None and not existing.empty:
            df = merge_sources(df, None, existing)

    df = trim_to_years(df, years)
    df = df.drop_duplicates(subset=["Date"], keep="last").sort_values("Date")

    n_unique = df["Date"].nunique()
    approx_trading_days = years * 252
    if len(df) < approx_trading_days * 0.5:
        print(
            f"NOTE: Only {len(df)} rows (~{n_unique} days) in the last {years} years. "
            "Moneycontrol embeds ~30 recent days; NSE JSON is latest day only. "
            f"Add bulk history via {BACKFILL_CSV.name} (columns matching the sheet) "
            "from an NSE historical CSV export, then re-run."
        )

    return df


def main() -> None:
    parser = argparse.ArgumentParser(description="Download FII/DII cash data to Excel.")
    parser.add_argument(
        "--no-merge-existing",
        action="store_true",
        help="Do not merge rows from an existing fii_dii_data.xlsx",
    )
    parser.add_argument(
        "--backfill",
        type=Path,
        default=None,
        help=f"Optional CSV path (default: {BACKFILL_CSV} if present)",
    )
    parser.add_argument(
        "--years",
        type=int,
        default=YEARS_HISTORY,
        help=f"Keep only the last N years (default: {YEARS_HISTORY})",
    )
    args = parser.parse_args()

    back_path = args.backfill
    if back_path is None and BACKFILL_CSV.exists():
        back_path = BACKFILL_CSV

    df = run_download(
        backfill_path=back_path,
        merge_existing=not args.no_merge_existing,
        years=args.years,
    )
    if df is not None:
        save_to_excel(df, OUTPUT_FILE)


if __name__ == "__main__":
    main()
