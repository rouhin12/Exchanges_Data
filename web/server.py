from __future__ import annotations

"""Data aggregation backend for the Exchange Data Viewer.

This module exposes:
  - build_summary(agg, start, end, exchange, segment)
  - get_fii_dii_data(start, end)

Both are used by the Streamlit app in app.py. There is no HTTP server here.
"""

from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import math
import pandas as pd


BASE_DIR = Path(__file__).resolve().parents[1]

NSE_CASH_FILE = BASE_DIR / "NSE" / "Cash segment" / "nse_daily.xlsx"
BSE_CASH_FILE = BASE_DIR / "BSE" / "bse_daily.xlsx"
NSE_FNO_FILE = BASE_DIR / "NSE" / "FnO" / "nse_fno_consolidated.xlsx"
BSE_FNO_FILE = BASE_DIR / "BSE" / "FnO" / "bse_fno_consolidated.xlsx"
FII_DII_FILE = BASE_DIR / "FII_DII" / "fii_dii_data.xlsx"

_data_cache: dict[str, Any] = {
    "nse_cash": None,
    "bse_cash": None,
    "nse_fno": None,
    "bse_fno": None,
    "fii_dii": None,
}


def to_numeric(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.replace(" ", "", regex=False),
        errors="coerce",
    )
    numeric = numeric.fillna(0).clip(lower=0)
    return numeric


def get_numeric_series(df: pd.DataFrame, column: str) -> pd.Series:
    if column in df.columns:
        return to_numeric(df[column])
    return pd.Series(0, index=df.index, dtype=float)


def sanitize_row(row: dict) -> dict:
    sanitized: dict[str, Any] = {}
    for key, value in row.items():
        if isinstance(value, (int, float)):
            val = float(value)
            if pd.isna(val) or math.isinf(val) or val < 0:
                sanitized[key] = 0.0
            else:
                sanitized[key] = round(val, 2) if isinstance(value, float) else val
        else:
            sanitized[key] = value
    return sanitized


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


def period_label(agg: str, dt_series: pd.Series) -> pd.Series:
    if agg == "daily":
        return dt_series.dt.normalize()
    if agg == "weekly":
        start = dt_series - pd.to_timedelta(dt_series.dt.weekday, unit="D")
        return start.dt.normalize()
    if agg == "monthly":
        return dt_series.dt.to_period("M").dt.to_timestamp()
    if agg == "quarterly":
        months = dt_series.dt.month
        years = dt_series.dt.year
        q_start_month = pd.Series(
            [1 if m < 4 else 4 if m < 7 else 7 if m < 10 else 10 for m in months],
            index=dt_series.index,
        )
        return pd.to_datetime(
            pd.DataFrame({"year": years, "month": q_start_month, "day": 1})
        )
    # Yearly (financial year starting April)
    year_start = dt_series.dt.year - (dt_series.dt.month < 4).astype(int)
    return pd.to_datetime(dict(year=year_start, month=4, day=1))


def format_period_label(agg: str, period_start: pd.Timestamp) -> str:
    if agg == "daily":
        return period_start.strftime("%d/%m/%Y")
    if agg == "quarterly":
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

    # Never show dates beyond "today" in labels.
    today = datetime.today().date()
    today_ts = pd.Timestamp(today)
    if period_end > today_ts:
        period_end = today_ts

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
        idx_fut_turnover_cols = ["Index Futures - Turnover (₹ Crores)"]
        eq_fut_turnover_cols = ["Stock Futures - Turnover (₹ Crores)"]
        fut_turnover_cols = idx_fut_turnover_cols + eq_fut_turnover_cols + [
            "Vol Futures - Turnover (₹ Crores)",
        ]
        io_notional_cols = ["Index Options - Notional Turnover (₹ Crores)"]
        io_premium_cols = ["Index Options - Premium Turnover (₹ Crores)"]
        eo_notional_cols = ["Stock Options - Notional Turnover (₹ Crores)"]
        eo_premium_cols = ["Stock Options - Premium Turnover (₹ Crores)"]
        idx_fut_volume_col = "Index Futures - No. Of Contracts"
        eq_fut_volume_col = "Stock Futures - No. Of Contracts"
        idx_opt_volume_col = "Index Options - No. Of Contracts"
        eq_opt_volume_col = "Stock Options - No. Of Contracts"
    else:
        cash_df = bse_cash
        fno_df = bse_fno
        cash_date_col = find_date_column(cash_df) or "Date"
        cash_value_col = select_column(cash_df, ["net", "turnover"]) or "Net Turnover"
        cash_volume_col = select_column(cash_df, ["shares"]) or "No. of Shares (Cr)"
        idx_fut_turnover_cols = ["Index Futures - Turnover (₹ Cr)"]
        eq_fut_turnover_cols = ["Stock Futures - Turnover (₹ Cr)"]
        fut_turnover_cols = idx_fut_turnover_cols + eq_fut_turnover_cols
        io_notional_cols = ["Index Options - Notional Turnover (₹ Cr)"]
        io_premium_cols = ["Index Options - Premium Turnover (₹ Cr)"]
        eo_notional_cols = ["Stock Options - Notional Turnover (₹ Cr)"]
        eo_premium_cols = ["Stock Options - Premium Turnover (₹ Cr)"]
        idx_fut_volume_col = "Index Futures - Volume (Shares/Contracts)"
        eq_fut_volume_col = "Stock Futures - Volume (Shares/Contracts)"
        idx_opt_volume_col = "Index Options - Volume (Shares/Contracts)"
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

            # Normalise traded volume to billions of shares:
            # - NSE source is in lakhs (1e5)  -> divide by 1e4 to get billions.
            # - BSE source is raw shares     -> divide by 1e9 to get billions.
            if exchange == "NSE":
                cash_volume_grouped = cash_volume_grouped / 1e4
            else:
                cash_volume_grouped = cash_volume_grouped / 1e9

            cash_days = cash_df.groupby("Period")["Date"].nunique().rename("cash_days")

    fno_grouped = pd.DataFrame()
    fno_days = pd.Series(dtype=int)
    if use_fno and not fno_df.empty:
        fno_date_col = find_date_column(fno_df) or "Date"
        fno_df = filter_dates(fno_df, fno_date_col, start, end)
        if not fno_df.empty:
            fno_df["Date"] = pd.to_datetime(fno_df[fno_date_col], errors="coerce")
            for col in fut_turnover_cols + io_notional_cols + io_premium_cols + eo_notional_cols + eo_premium_cols:
                if col in fno_df.columns:
                    fno_df[col] = to_numeric(fno_df[col])
            fno_df["Period"] = period_label(agg, fno_df["Date"])

            fut_turnover_cols_existing = [c for c in fut_turnover_cols if c in fno_df.columns]
            idx_fut_turnover_cols_existing = [c for c in idx_fut_turnover_cols if c in fno_df.columns]
            eq_fut_turnover_cols_existing = [c for c in eq_fut_turnover_cols if c in fno_df.columns]
            io_notional_cols_existing = [c for c in io_notional_cols if c in fno_df.columns]
            io_premium_cols_existing = [c for c in io_premium_cols if c in fno_df.columns]
            eo_notional_cols_existing = [c for c in eo_notional_cols if c in fno_df.columns]
            eo_premium_cols_existing = [c for c in eo_premium_cols if c in fno_df.columns]

            # Total futures turnover (index + stock [+ vol])
            futures_turnover = (
                fno_df[fut_turnover_cols_existing].sum(axis=1)
                if fut_turnover_cols_existing
                else pd.Series(0, index=fno_df.index)
            )
            # Split futures turnover into index vs stock
            index_futures_turnover = (
                fno_df[idx_fut_turnover_cols_existing].sum(axis=1)
                if idx_fut_turnover_cols_existing
                else pd.Series(0, index=fno_df.index)
            )
            equity_futures_turnover = (
                fno_df[eq_fut_turnover_cols_existing].sum(axis=1)
                if eq_fut_turnover_cols_existing
                else pd.Series(0, index=fno_df.index)
            )
            io_notional = (
                fno_df[io_notional_cols_existing].sum(axis=1)
                if io_notional_cols_existing
                else pd.Series(0, index=fno_df.index)
            )
            io_premium = (
                fno_df[io_premium_cols_existing].sum(axis=1)
                if io_premium_cols_existing
                else pd.Series(0, index=fno_df.index)
            )
            eo_notional = (
                fno_df[eo_notional_cols_existing].sum(axis=1)
                if eo_notional_cols_existing
                else pd.Series(0, index=fno_df.index)
            )
            eo_premium = (
                fno_df[eo_premium_cols_existing].sum(axis=1)
                if eo_premium_cols_existing
                else pd.Series(0, index=fno_df.index)
            )

            idx_fut_volume = get_numeric_series(fno_df, idx_fut_volume_col).fillna(0)
            eq_fut_volume = get_numeric_series(fno_df, eq_fut_volume_col).fillna(0)
            idx_opt_volume = get_numeric_series(fno_df, idx_opt_volume_col).fillna(0)
            eq_opt_volume = get_numeric_series(fno_df, eq_opt_volume_col).fillna(0)
            total_contracts = idx_fut_volume + eq_fut_volume + idx_opt_volume + eq_opt_volume

            metrics_df = pd.DataFrame(
                {
                    "Period": fno_df["Period"],
                    "futures_turnover": futures_turnover,
                    "index_futures_turnover": index_futures_turnover,
                    "equity_futures_turnover": equity_futures_turnover,
                    "index_options_notional": io_notional,
                    "index_options_premium": io_premium,
                    "equity_options_notional": eo_notional,
                    "equity_options_premium": eo_premium,
                    "index_futures_volume": idx_fut_volume,
                    "equity_futures_volume": eq_fut_volume,
                    "index_options_volume": idx_opt_volume,
                    "equity_options_volume": eq_opt_volume,
                    "total_contracts": total_contracts,
                }
            )
            fno_grouped = metrics_df.groupby("Period").sum()
            fno_days = fno_df.groupby("Period")["Date"].nunique().rename("fno_days")

    if cash_grouped.empty and fno_grouped.empty:
        return rows

    periods = sorted(set(cash_grouped.index).union(set(fno_grouped.index)))
    for period in periods:
        period_start = pd.Timestamp(period)
        period_label_text = format_period_label(agg, period_start)

        if segment == "cash":
            days = int(cash_days.get(period, 0)) or 1
        elif segment == "fno":
            days = int(fno_days.get(period, 0)) or 1
        else:
            cash_day_count = int(cash_days.get(period, 0)) or 0
            fno_day_count = int(fno_days.get(period, 0)) or 0
            days = max(cash_day_count, fno_day_count) or 1

        cash_val = float(cash_grouped.get(period, 0) or 0) / 100.0
        cash_val = max(0, cash_val)
        cash_vol_val = float(cash_volume_grouped.get(period, 0) or 0)
        cash_vol_val = max(0, cash_vol_val)

        if period in fno_grouped.index:
            row = fno_grouped.loc[period]
            futures_val = float(row.get("futures_turnover", 0) or 0) / 100.0
            # Split futures turnover into index vs stock for separate tables
            idx_fut_turnover_val = float(row.get("index_futures_turnover", 0) or 0) / 100.0
            eq_fut_turnover_val = float(row.get("equity_futures_turnover", 0) or 0) / 100.0
            io_notional_val = float(row.get("index_options_notional", 0) or 0) / 100.0
            io_premium_val = float(row.get("index_options_premium", 0) or 0) / 100.0
            eo_notional_val = float(row.get("equity_options_notional", 0) or 0) / 100.0
            eo_premium_val = float(row.get("equity_options_premium", 0) or 0) / 100.0
            idx_fut_vol_val = float(row.get("index_futures_volume", 0) or 0)
            eq_fut_vol_val = float(row.get("equity_futures_volume", 0) or 0)
            idx_opt_vol_val = float(row.get("index_options_volume", 0) or 0)
            eq_opt_vol_val = float(row.get("equity_options_volume", 0) or 0)
            total_contracts_val = float(row.get("total_contracts", 0) or 0)
        else:
            futures_val = 0.0
            idx_fut_turnover_val = 0.0
            eq_fut_turnover_val = 0.0
            io_notional_val = 0.0
            io_premium_val = 0.0
            eo_notional_val = 0.0
            eo_premium_val = 0.0
            idx_fut_vol_val = 0.0
            eq_fut_vol_val = 0.0
            idx_opt_vol_val = 0.0
            eq_opt_vol_val = 0.0
            total_contracts_val = 0.0

        future_vals = [
            futures_val,
            idx_fut_turnover_val,
            eq_fut_turnover_val,
            io_notional_val,
            io_premium_val,
            eo_notional_val,
            eo_premium_val,
            idx_fut_vol_val,
            eq_fut_vol_val,
            idx_opt_vol_val,
            eq_opt_vol_val,
            total_contracts_val,
        ]
        future_vals = [max(0, v) for v in future_vals]
        (
            futures_val,
            idx_fut_turnover_val,
            eq_fut_turnover_val,
            io_notional_val,
            io_premium_val,
            eo_notional_val,
            eo_premium_val,
            idx_fut_vol_val,
            eq_fut_vol_val,
            idx_opt_vol_val,
            eq_opt_vol_val,
            total_contracts_val,
        ) = future_vals

        avg_cash_turnover = cash_val / days if days > 0 else 0.0
        avg_futures_turnover = futures_val / days if days > 0 else 0.0
        avg_index_futures_turnover = idx_fut_turnover_val / days if days > 0 else 0.0
        avg_equity_futures_turnover = eq_fut_turnover_val / days if days > 0 else 0.0
        avg_io_notional = io_notional_val / days if days > 0 else 0.0
        avg_io_premium = io_premium_val / days if days > 0 else 0.0
        avg_eo_notional = eo_notional_val / days if days > 0 else 0.0
        avg_eo_premium = eo_premium_val / days if days > 0 else 0.0
        avg_cash_volume = cash_vol_val / days if days > 0 else 0.0
        avg_contracts = total_contracts_val / days if days > 0 else 0.0

        row_dict = {
            "period": period_label_text,
            "period_sort": period_start.strftime("%Y-%m-%d"),
            "exchange": exchange,
            "trading_days": days,
            "cash_turnover_bn": round(cash_val, 2),
            "cash_volume": round(cash_vol_val, 2),
            "futures_turnover_bn": round(futures_val, 2),
            "index_futures_turnover_bn": round(idx_fut_turnover_val, 2),
            "equity_futures_turnover_bn": round(eq_fut_turnover_val, 2),
            "index_options_notional_bn": round(io_notional_val, 2),
            "index_options_premium_bn": round(io_premium_val, 2),
            "equity_options_notional_bn": round(eo_notional_val, 2),
            "equity_options_premium_bn": round(eo_premium_val, 2),
            "index_futures_volume": round(idx_fut_vol_val, 2),
            "equity_futures_volume": round(eq_fut_vol_val, 2),
            "index_options_volume": round(idx_opt_vol_val, 2),
            "equity_options_volume": round(eq_opt_vol_val, 2),
            "total_contracts": round(total_contracts_val, 2),
            "avg_cash_turnover_bn": round(avg_cash_turnover, 2),
            "avg_futures_turnover_bn": round(avg_futures_turnover, 2),
            "avg_index_futures_turnover_bn": round(avg_index_futures_turnover, 2),
            "avg_equity_futures_turnover_bn": round(avg_equity_futures_turnover, 2),
            "avg_index_options_notional_bn": round(avg_io_notional, 2),
            "avg_index_options_premium_bn": round(avg_io_premium, 2),
            "avg_equity_options_notional_bn": round(avg_eo_notional, 2),
            "avg_equity_options_premium_bn": round(avg_eo_premium, 2),
            "avg_cash_volume": round(avg_cash_volume, 2),
            "avg_contracts": round(avg_contracts, 2),
        }

        row_dict = sanitize_row(row_dict)
        rows.append(row_dict)

    return rows


def build_summary(
    agg: str,
    start: str | None,
    end: str | None,
    exchange: str,
    segment: str,
) -> dict[str, object]:
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


def get_fii_dii_data(start: str | None, end: str | None, agg: str = "daily") -> dict[str, object]:
    """Get FII/DII data from Excel file, aggregated like the main summary.

    Returns values in Rs bn. `agg` can be "daily", "weekly", "monthly",
    "quarterly", "yearly", or "entire" (entire period).
    """
    global _data_cache

    if _data_cache["fii_dii"] is None:
        if not FII_DII_FILE.exists():
            return {"rows": []}

        df = pd.read_excel(FII_DII_FILE)
        df["Date"] = pd.to_datetime(df["Date"])
        _data_cache["fii_dii"] = df
    else:
        df = _data_cache["fii_dii"]

    if start:
        df = df[df["Date"] >= pd.to_datetime(start)]
    if end:
        df = df[df["Date"] <= pd.to_datetime(end)]

    if df.empty:
        return {"rows": []}

    df = df.copy()
    df["Date"] = pd.to_datetime(df["Date"])

    if agg == "entire":
        # Single bucket for the entire period
        period_start = df["Date"].min().normalize()
        df["Period"] = period_start
    elif agg == "daily":
        df["Period"] = df["Date"].dt.normalize()
    else:
        df["Period"] = period_label(agg, df["Date"])

    grouped = df.groupby("Period").agg(
        {
            "FII_Gross_Purchase": "sum",
            "FII_Gross_Sales": "sum",
            "FII_Net": "sum",
            "DII_Gross_Purchase": "sum",
            "DII_Gross_Sales": "sum",
            "DII_Net": "sum",
        }
    )

    rows: list[dict[str, object]] = []
    for period, row in grouped.iterrows():
        period_start = pd.Timestamp(period)
        if agg == "daily":
            label = period_start.strftime("%d/%m/%Y")
        else:
            label = format_period_label(agg if agg != "entire" else "yearly", period_start)

        rows.append(
            {
                "date": label,
                "date_sort": period_start.strftime("%Y-%m-%d"),
                # Crores -> Rs bn
                "fii_gross_purchase": round(float(row["FII_Gross_Purchase"]) / 100.0, 2),
                "fii_gross_sales": round(float(row["FII_Gross_Sales"]) / 100.0, 2),
                "fii_net": round(float(row["FII_Net"]) / 100.0, 2),
                "dii_gross_purchase": round(float(row["DII_Gross_Purchase"]) / 100.0, 2),
                "dii_gross_sales": round(float(row["DII_Gross_Sales"]) / 100.0, 2),
                "dii_net": round(float(row["DII_Net"]) / 100.0, 2),
            }
        )

    rows = sorted(rows, key=lambda r: r["date_sort"], reverse=True)
    return {"rows": rows}

