#!/usr/bin/env python3
"""Serve the Exchange Data Viewer web page with Excel-backed API."""

from __future__ import annotations

import json
from datetime import datetime, timedelta
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import pandas as pd


BASE_DIR = Path(__file__).resolve().parents[1]
WEB_DIR = BASE_DIR / "web"

NSE_CASH_FILE = BASE_DIR / "NSE" / "Cash segment" / "nse_daily.xlsx"
BSE_CASH_FILE = BASE_DIR / "BSE" / "bse_daily.xlsx"
NSE_FNO_FILE = BASE_DIR / "NSE" / "FnO" / "nse_fno_consolidated.xlsx"
BSE_FNO_FILE = BASE_DIR / "BSE" / "FnO" / "bse_fno_consolidated.xlsx"
FII_DII_FILE = BASE_DIR / "FII_DII" / "fii_dii_data.xlsx"

# In-memory cache for Excel data
_data_cache = {
    "nse_cash": None,
    "bse_cash": None,
    "nse_fno": None,
    "bse_fno": None,
    "fii_dii": None,
}


def to_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.replace(" ", "", regex=False),
        errors="coerce",
    )


def get_numeric_series(df: pd.DataFrame, column: str) -> pd.Series:
    if column in df.columns:
        return to_numeric(df[column])
    return pd.Series(0, index=df.index, dtype=float)


def load_excel_file(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    xls = pd.ExcelFile(path)
    frames: list[pd.DataFrame] = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df["Source_File"] = path.name
        df["Source_Sheet"] = sheet
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def find_date_column(df: pd.DataFrame) -> str | None:
    for col in df.columns:
        col_l = str(col).lower()
        if "date" in col_l or "timestamp" in col_l:
            return col
    return None


def select_column(df: pd.DataFrame, keywords: list[str]) -> str | None:
    keywords = [k.lower() for k in keywords]
    for col in df.columns:
        col_l = str(col).lower()
        if all(k in col_l for k in keywords):
            return col
    return None


def filter_dates(df: pd.DataFrame, date_col: str, start: str | None, end: str | None) -> pd.DataFrame:
    df = df.copy()
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=[date_col])
    if start:
        df = df[df[date_col] >= pd.to_datetime(start)]
    if end:
        df = df[df[date_col] <= pd.to_datetime(end)]
    return df


def to_fy_label(dt: pd.Timestamp) -> str:
    year_end = dt.year + 1 if dt.month >= 4 else dt.year
    return f"FY{year_end % 100:02d}"


def period_label(agg: str, dt_series: pd.Series) -> pd.Series:
    if agg == "daily":
        return dt_series.dt.normalize()
    if agg == "weekly":
        start = dt_series - pd.to_timedelta(dt_series.dt.weekday, unit="D")
        return start.dt.normalize()
    if agg == "monthly":
        return dt_series.dt.to_period("M").dt.to_timestamp()
    if agg == "quarterly":
        # Financial quarters: Q1=Apr-Jun, Q2=Jul-Sep, Q3=Oct-Dec, Q4=Jan-Mar
        months = dt_series.dt.month
        years = dt_series.dt.year
        
        # Determine quarter start month:
        # Jan-Mar (1-3): month 1 → Timestamp(year, 1, 1) → Q4 of previous FY
        # Apr-Jun (4-6): month 4 → Timestamp(year, 4, 1) → Q1 of current FY  
        # Jul-Sep (7-9): month 7 → Timestamp(year, 7, 1) → Q2 of current FY
        # Oct-Dec (10-12): month 10 → Timestamp(year, 10, 1) → Q3 of current FY
        q_start_month = pd.Series(
            [1 if m < 4 else 4 if m < 7 else 7 if m < 10 else 10 for m in months],
            index=dt_series.index
        )
        
        return pd.to_datetime(
            pd.DataFrame({
                'year': years,
                'month': q_start_month,
                'day': 1
            })
        )
    
    # Yearly (default)
    year_start = dt_series.dt.year - (dt_series.dt.month < 4).astype(int)
    return pd.to_datetime(dict(year=year_start, month=4, day=1))


def format_period_label(agg: str, period_start: pd.Timestamp) -> str:
    if agg == "daily":
        return period_start.strftime("%d/%m/%Y")
    if agg == "quarterly":
        # Financial quarters: Q1=Apr-Jun, Q2=Jul-Sep, Q3=Oct-Dec, Q4=Jan-Mar
        month = period_start.month
        year = period_start.year
        if month < 4:
            fy_year = year
            quarter = 4
        else:
            fy_year = year + 1
            quarter = (month - 4) // 3 + 1
        return f"Q{quarter} FY{fy_year % 100}"
    if agg == "weekly":
        period_end = period_start + pd.Timedelta(days=6)
    elif agg == "monthly":
        period_end = period_start + pd.offsets.MonthEnd(0)
    else:
        period_end = period_start + pd.DateOffset(years=1) - pd.Timedelta(days=1)

    start_text = period_start.strftime("%d/%m/%Y")
    end_text = period_end.strftime("%d/%m/%Y")
    return f"{start_text} to {end_text}"


def aggregate_exchange(
    exchange: str,
    agg: str,
    start: str | None,
    end: str | None,
    segment: str,
    nse_cash: pd.DataFrame,
    bse_cash: pd.DataFrame,
    nse_fno: pd.DataFrame,
    bse_fno: pd.DataFrame,
) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    use_cash = segment in ("cash", "both")
    use_fno = segment in ("fno", "both")

    if exchange == "NSE":
        cash_df = nse_cash
        fno_df = nse_fno
        cash_date_col = find_date_column(cash_df) or "Date"
        cash_value_col = select_column(cash_df, ["traded", "value"]) or "Traded Value (₹ Crores)"
        cash_volume_col = select_column(cash_df, ["traded", "quantity"]) or "Traded quantity (in Lakhs)"
        fut_turnover_cols = [
            "Index Futures - Turnover (₹ Crores)",
            "Stock Futures - Turnover (₹ Crores)",
            "Vol Futures - Turnover (₹ Crores)",
        ]
        io_notional_cols = ["Index Options - Notional Turnover (₹ Crores)"]
        io_premium_cols = ["Index Options - Premium Turnover (₹ Crores)"]
        eo_notional_cols = ["Stock Options - Notional Turnover (₹ Crores)"]
        idx_fut_volume_col = "Index Futures - No. Of Contracts"
        eq_fut_volume_col = "Stock Futures - No. Of Contracts"
        eq_opt_volume_col = "Stock Options - No. Of Contracts"
    else:
        cash_df = bse_cash
        fno_df = bse_fno
        cash_date_col = find_date_column(cash_df) or "Date"
        cash_value_col = select_column(cash_df, ["net", "turnover"]) or "Net Turnover"
        cash_volume_col = select_column(cash_df, ["shares"]) or "No. of Shares (Cr)"
        fut_turnover_cols = [
            "Index Futures - Turnover (₹ Cr)",
            "Stock Futures - Turnover (₹ Cr)",
        ]
        io_notional_cols = ["Index Options - Notional Turnover (₹ Cr)"]
        io_premium_cols = ["Index Options - Premium Turnover (₹ Cr)"]
        eo_notional_cols = ["Stock Options - Notional Turnover (₹ Cr)"]
        idx_fut_volume_col = "Index Futures - Volume (Shares/Contracts)"
        eq_fut_volume_col = "Stock Futures - Volume (Shares/Contracts)"
        eq_opt_volume_col = "Stock Options - Volume (Shares/Contracts)"

    cash_grouped = pd.Series(dtype=float)
    cash_volume_grouped = pd.Series(dtype=float)
    cash_days = pd.Series(dtype=int)
    if use_cash and not cash_df.empty:
        cash_df = filter_dates(cash_df, cash_date_col, start, end)
        if not cash_df.empty:
            cash_df["Date"] = pd.to_datetime(cash_df[cash_date_col], errors="coerce")
            cash_df["__cash_value"] = get_numeric_series(cash_df, cash_value_col)
            cash_df["__cash_volume"] = get_numeric_series(cash_df, cash_volume_col)
            cash_df["Period"] = period_label(agg, cash_df["Date"])
            cash_grouped = cash_df.groupby("Period")["__cash_value"].sum().rename("cash_turnover")
            cash_volume_grouped = cash_df.groupby("Period")["__cash_volume"].sum().rename("cash_volume")
            cash_days = cash_df.groupby("Period")["Date"].nunique().rename("cash_days")

    fno_grouped = pd.DataFrame()
    fno_days = pd.Series(dtype=int)
    if use_fno and not fno_df.empty:
        fno_date_col = find_date_column(fno_df) or "Date"
        fno_df = filter_dates(fno_df, fno_date_col, start, end)
        if not fno_df.empty:
            fno_df["Date"] = pd.to_datetime(fno_df[fno_date_col], errors="coerce")
            for col in fut_turnover_cols + io_notional_cols + io_premium_cols + eo_notional_cols:
                if col in fno_df.columns:
                    fno_df[col] = to_numeric(fno_df[col])
            fno_df["Period"] = period_label(agg, fno_df["Date"])

            fut_turnover_cols_existing = [c for c in fut_turnover_cols if c in fno_df.columns]
            io_notional_cols_existing = [c for c in io_notional_cols if c in fno_df.columns]
            io_premium_cols_existing = [c for c in io_premium_cols if c in fno_df.columns]
            eo_notional_cols_existing = [c for c in eo_notional_cols if c in fno_df.columns]

            futures_turnover = fno_df[fut_turnover_cols_existing].sum(axis=1) if fut_turnover_cols_existing else pd.Series(0, index=fno_df.index)
            io_notional = fno_df[io_notional_cols_existing].sum(axis=1) if io_notional_cols_existing else pd.Series(0, index=fno_df.index)
            io_premium = fno_df[io_premium_cols_existing].sum(axis=1) if io_premium_cols_existing else pd.Series(0, index=fno_df.index)
            eo_notional = fno_df[eo_notional_cols_existing].sum(axis=1) if eo_notional_cols_existing else pd.Series(0, index=fno_df.index)

            idx_fut_volume = get_numeric_series(fno_df, idx_fut_volume_col)
            eq_fut_volume = get_numeric_series(fno_df, eq_fut_volume_col)
            eq_opt_volume = get_numeric_series(fno_df, eq_opt_volume_col)
            total_contracts = idx_fut_volume + eq_fut_volume + eq_opt_volume

            metrics_df = pd.DataFrame({
                "Period": fno_df["Period"],
                "futures_turnover": futures_turnover,
                "index_options_notional": io_notional,
                "index_options_premium": io_premium,
                "equity_options_notional": eo_notional,
                "index_futures_volume": idx_fut_volume,
                "equity_futures_volume": eq_fut_volume,
                "equity_options_volume": eq_opt_volume,
                "total_contracts": total_contracts,
            })
            fno_grouped = metrics_df.groupby("Period").sum()
            fno_days = fno_df.groupby("Period")["Date"].nunique().rename("fno_days")

    if cash_grouped.empty and fno_grouped.empty:
        return rows

    periods = sorted(set(cash_grouped.index).union(set(fno_grouped.index)))
    for period in periods:
        period_start = pd.Timestamp(period)
        period_label_text = format_period_label(agg, period_start)

        if segment == "cash":
            days = int(cash_days.get(period, 0))
        elif segment == "fno":
            days = int(fno_days.get(period, 0))
        else:
            days = int(max(cash_days.get(period, 0), fno_days.get(period, 0)))

        cash_val = float(cash_grouped.get(period, 0)) / 100.0
        cash_vol_val = float(cash_volume_grouped.get(period, 0))

        if period in fno_grouped.index:
            row = fno_grouped.loc[period]
            futures_val = float(row.get("futures_turnover", 0)) / 100.0
            io_notional_val = float(row.get("index_options_notional", 0)) / 100.0
            io_premium_val = float(row.get("index_options_premium", 0)) / 100.0
            eo_notional_val = float(row.get("equity_options_notional", 0)) / 100.0
            idx_fut_vol_val = float(row.get("index_futures_volume", 0))
            eq_fut_vol_val = float(row.get("equity_futures_volume", 0))
            eq_opt_vol_val = float(row.get("equity_options_volume", 0))
            total_contracts_val = float(row.get("total_contracts", 0))
        else:
            futures_val = 0.0
            io_notional_val = 0.0
            io_premium_val = 0.0
            eo_notional_val = 0.0
            idx_fut_vol_val = 0.0
            eq_fut_vol_val = 0.0
            eq_opt_vol_val = 0.0
            total_contracts_val = 0.0

        avg_cash_turnover = cash_val / days if days else 0.0
        avg_futures_turnover = futures_val / days if days else 0.0
        avg_io_notional = io_notional_val / days if days else 0.0
        avg_io_premium = io_premium_val / days if days else 0.0
        avg_eo_notional = eo_notional_val / days if days else 0.0
        avg_cash_volume = cash_vol_val / days if days else 0.0
        avg_contracts = total_contracts_val / days if days else 0.0

        rows.append({
            "period": period_label_text,
            "period_sort": period_start.strftime("%Y-%m-%d"),
            "exchange": exchange,
            "trading_days": days,
            "cash_turnover_bn": round(cash_val, 2),
            "cash_volume": round(cash_vol_val, 2),
            "futures_turnover_bn": round(futures_val, 2),
            "index_options_notional_bn": round(io_notional_val, 2),
            "index_options_premium_bn": round(io_premium_val, 2),
            "equity_options_notional_bn": round(eo_notional_val, 2),
            "index_futures_volume": round(idx_fut_vol_val, 2),
            "equity_futures_volume": round(eq_fut_vol_val, 2),
            "equity_options_volume": round(eq_opt_vol_val, 2),
            "total_contracts": round(total_contracts_val, 2),
            "avg_cash_turnover_bn": round(avg_cash_turnover, 2),
            "avg_futures_turnover_bn": round(avg_futures_turnover, 2),
            "avg_index_options_notional_bn": round(avg_io_notional, 2),
            "avg_index_options_premium_bn": round(avg_io_premium, 2),
            "avg_equity_options_notional_bn": round(avg_eo_notional, 2),
            "avg_cash_volume": round(avg_cash_volume, 2),
            "avg_contracts": round(avg_contracts, 2),
        })

    return rows


def build_summary(
    agg: str,
    start: str | None,
    end: str | None,
    exchange: str,
    segment: str,
) -> dict[str, object]:
    # Load from cache, or load and cache if not present
    global _data_cache
    
    if _data_cache["nse_cash"] is None:
        _data_cache["nse_cash"] = load_excel_file(NSE_CASH_FILE)
        _data_cache["bse_cash"] = load_excel_file(BSE_CASH_FILE)
        _data_cache["nse_fno"] = load_excel_file(NSE_FNO_FILE)
        _data_cache["bse_fno"] = load_excel_file(BSE_FNO_FILE)
    
    nse_cash = _data_cache["nse_cash"]
    bse_cash = _data_cache["bse_cash"]
    nse_fno = _data_cache["nse_fno"]
    bse_fno = _data_cache["bse_fno"]

    rows: list[dict[str, object]] = []
    if exchange in ("all", "NSE"):
        rows.extend(aggregate_exchange("NSE", agg, start, end, segment, nse_cash, bse_cash, nse_fno, bse_fno))
    if exchange in ("all", "BSE"):
        rows.extend(aggregate_exchange("BSE", agg, start, end, segment, nse_cash, bse_cash, nse_fno, bse_fno))

    rows = sorted(rows, key=lambda r: (r.get("period_sort", ""), r["exchange"]))
    return {"rows": rows}


def get_fii_dii_data(start: str | None, end: str | None) -> dict[str, object]:
    """Get FII/DII data from Excel file."""
    global _data_cache
    
    if _data_cache["fii_dii"] is None:
        if not FII_DII_FILE.exists():
            return {"rows": []}
        
        df = pd.read_excel(FII_DII_FILE)
        df["Date"] = pd.to_datetime(df["Date"])
        _data_cache["fii_dii"] = df
    else:
        df = _data_cache["fii_dii"]
    
    # Filter by date range
    if start:
        df = df[df["Date"] >= pd.to_datetime(start)]
    if end:
        df = df[df["Date"] <= pd.to_datetime(end)]
    
    # Convert to JSON-serializable format
    rows = []
    for _, row in df.iterrows():
        rows.append({
            "date": row["Date"].strftime("%d-%b-%Y"),
            "date_sort": row["Date"].strftime("%Y-%m-%d"),
            "fii_gross_purchase": round(float(row["FII_Gross_Purchase"]), 2),
            "fii_gross_sales": round(float(row["FII_Gross_Sales"]), 2),
            "fii_net": round(float(row["FII_Net"]), 2),
            "dii_gross_purchase": round(float(row["DII_Gross_Purchase"]), 2),
            "dii_gross_sales": round(float(row["DII_Gross_Sales"]), 2),
            "dii_net": round(float(row["DII_Net"]), 2),
        })
    
    rows = sorted(rows, key=lambda r: r["date_sort"], reverse=True)
    return {"rows": rows}


class Handler(SimpleHTTPRequestHandler):
    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/api/summary":
            query = parse_qs(parsed.query)
            agg = query.get("agg", ["weekly"])[0]
            start = query.get("from", [None])[0]
            end = query.get("to", [None])[0]
            exchange = query.get("exchange", ["all"])[0]
            segment = query.get("segment", ["both"])[0]

            try:
                payload = build_summary(agg, start, end, exchange, segment)
                body = json.dumps(payload).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as exc:
                body = json.dumps({"error": str(exc)}).encode("utf-8")
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            return

        if parsed.path == "/api/fii_dii":
            query = parse_qs(parsed.query)
            start = query.get("from", [None])[0]
            end = query.get("to", [None])[0]

            try:
                payload = get_fii_dii_data(start, end)
                body = json.dumps(payload).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as exc:
                body = json.dumps({"error": str(exc)}).encode("utf-8")
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            return

        super().do_GET()


def main() -> None:
    server = ThreadingHTTPServer(("", 8000), Handler)
    server.RequestHandlerClass.directory = str(WEB_DIR)
    print("Server running at http://localhost:8000")
    server.serve_forever()


if __name__ == "__main__":
    main()
