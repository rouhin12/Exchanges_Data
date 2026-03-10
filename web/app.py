from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Literal

import pandas as pd
import streamlit as st
import altair as alt

from server import build_summary, get_fii_dii_data


Exchange = Literal["all", "NSE", "BSE"]
Segment = Literal["both", "cash", "fno"]
Agg = Literal["daily", "weekly", "monthly", "quarterly", "yearly", "entire"]


@dataclass(frozen=True)
class Range:
    start: date
    end: date


def _to_iso(d: date | None) -> str | None:
    return d.isoformat() if d else None


def _df_from_rows(rows: list[dict]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    if "period_sort" in df.columns:
        df["period_sort"] = pd.to_datetime(df["period_sort"], errors="coerce")
    return df


@st.cache_data(show_spinner=False)
def _summary_df(agg: Agg, start: date | None, end: date | None, exchange: Exchange, segment: Segment) -> pd.DataFrame:
    payload = build_summary(agg, _to_iso(start), _to_iso(end), exchange, segment)
    return _df_from_rows(payload.get("rows", []))


def _entire_period_summary(start: date | None, end: date | None, exchange: Exchange, segment: Segment) -> pd.DataFrame:
    df = _summary_df("daily", start, end, exchange, segment)
    if df.empty:
        return df

    numeric_cols = [
        c
        for c in df.columns
        if c
        not in {
            "period",
            "period_sort",
            "exchange",
        }
        and pd.api.types.is_numeric_dtype(df[c])
    ]
    grouped = df.groupby("exchange", dropna=False)[numeric_cols].sum().reset_index()
    grouped.insert(0, "period_sort", pd.to_datetime(_to_iso(start) or df["period_sort"].min()).normalize())
    label = f"{_to_iso(start) or ''} to {_to_iso(end) or ''}".strip()
    grouped.insert(0, "period", label if label != "to" else "Entire period")

    # Recompute averages using aggregated trading_days (sum across daily rows).
    if "trading_days" in grouped.columns:
        days = grouped["trading_days"].replace(0, 1)
        avg_map = {
            "avg_cash_turnover_bn": ("cash_turnover_bn", days),
            "avg_futures_turnover_bn": ("futures_turnover_bn", days),
            "avg_index_options_notional_bn": ("index_options_notional_bn", days),
            "avg_index_options_premium_bn": ("index_options_premium_bn", days),
            "avg_equity_options_notional_bn": ("equity_options_notional_bn", days),
            "avg_equity_options_premium_bn": ("equity_options_premium_bn", days),
            "avg_cash_volume": ("cash_volume", days),
            "avg_contracts": ("total_contracts", days),
        }
        for out_col, (src_col, denom) in avg_map.items():
            if src_col in grouped.columns:
                grouped[out_col] = (pd.to_numeric(grouped[src_col], errors="coerce").fillna(0) / denom).round(2)

    return grouped


def _format_num(x: object, digits: int = 2) -> str:
    if x is None:
        return "-"
    try:
        v = float(x)
    except Exception:
        return "-"
    if pd.isna(v):
        return "-"
    return f"{v:,.{digits}f}"


def _format_int(x: object) -> str:
    if x is None:
        return "-"
    try:
        v = float(x)
    except Exception:
        return "-"
    if pd.isna(v):
        return "-"
    return f"{int(round(v)):,}"


def _preset_last_4_weeks(anchor: date) -> Range:
    return Range(start=anchor - timedelta(days=28), end=anchor)


def _preset_ytd(anchor: date) -> Range:
    return Range(start=date(anchor.year, 1, 1), end=anchor)


def _preset_last_fy(anchor: date) -> Range:
    # FY: Apr 1 → Mar 31
    year = anchor.year - 1 if anchor.month >= 4 else anchor.year - 2
    return Range(start=date(year, 4, 1), end=date(year + 1, 3, 31))


def _set_preset_range(kind: str, anchor: date) -> None:
    """Callback for quick range buttons to safely update date inputs."""
    if kind == "last4w":
        r = _preset_last_4_weeks(anchor)
    elif kind == "ytd":
        r = _preset_ytd(anchor)
    elif kind == "lastfy":
        r = _preset_last_fy(anchor)
    elif kind == "fyytd":
        # From start of current financial year to latest info date
        # Current FY: starts on 1 Apr of the year if month >= 4, else 1 Apr previous year
        fy_start_year = anchor.year if anchor.month >= 4 else anchor.year - 1
        r = Range(start=date(fy_start_year, 4, 1), end=anchor)
    else:
        return
    st.session_state["from_date"] = r.start
    st.session_state["to_date"] = r.end


@st.cache_data(show_spinner=False)
def _last_market_day(today: date) -> date:
    # Determine the latest info date available in the full daily series.
    df = _summary_df("daily", None, None, "all", "both")
    if df.empty or "period_sort" not in df.columns:
        return today - timedelta(days=1)
    last = df["period_sort"].dropna().max()
    if pd.isna(last):
        return today - timedelta(days=1)
    return last.date()


def _segment_tables(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    if df.empty:
        return {}

    def by_exchange(total_col: str, days_col: str = "trading_days", extra: dict[str, str] | None = None) -> pd.DataFrame:
        extra = extra or {}
        needed = ["exchange", total_col, days_col, *extra.values()]
        for c in needed:
            if c not in df.columns:
                return pd.DataFrame()
        g = (
            df.groupby("exchange", dropna=False)[[total_col, days_col, *extra.values()]]
            .sum(numeric_only=True)
            .reset_index()
        )
        g["avg_daily"] = (g[total_col] / g[days_col].replace(0, 1)).round(2)
        total = float(g[total_col].sum() or 0)
        if total > 0:
            g["market_share_pct"] = (g[total_col] / total * 100).round(2)
        else:
            g["market_share_pct"] = 0.0
        out = g.rename(
            columns={
                "exchange": "Exchange",
                days_col: "Trading Days",
                total_col: "Total Turnover (Rs bn)",
                "avg_daily": "Avg Daily Turnover (Rs bn)",
                "market_share_pct": "Market Share (%)",
                **{k: k for k in extra.keys()},
            }
        )
        for label, col in extra.items():
            out = out.rename(columns={col: label})
        cols = ["Exchange", "Trading Days", "Total Turnover (Rs bn)", "Market Share (%)", "Avg Daily Turnover (Rs bn)"]
        cols += [c for c in extra.keys() if c in out.columns]
        return out[cols]

    tables: dict[str, pd.DataFrame] = {}

    cash = by_exchange("cash_turnover_bn", extra={"Total Volume": "cash_volume"})
    if not cash.empty:
        tables["Cash Segment"] = cash

    idx_fut = by_exchange("futures_turnover_bn", extra={"Index Futures Contracts": "index_futures_volume"})
    if not idx_fut.empty:
        tables["Index Futures"] = idx_fut

    stk_fut = by_exchange("futures_turnover_bn", extra={"Stock Futures Contracts": "equity_futures_volume"})
    if not stk_fut.empty:
        tables["Stock Futures"] = stk_fut

    idx_opt = by_exchange("index_options_notional_bn", extra={"Index Options Contracts": "index_options_volume"})
    if not idx_opt.empty:
        tables["Index Options"] = idx_opt

    stk_opt = by_exchange("equity_options_notional_bn", extra={"Stock Options Contracts": "equity_options_volume"})
    if not stk_opt.empty:
        tables["Stock Options"] = stk_opt

    return tables


def _comparison_payload(
    start: date,
    end: date,
    exchange: Exchange,
    segment: Segment,
    today: date,
) -> tuple[pd.DataFrame, pd.DataFrame, Range, Range]:
    # Cap selected range to today
    if end > today:
        end = today
    selected = Range(start=start, end=end)

    prev_start = date(start.year - 1, start.month, start.day)
    prev_end = date(end.year - 1, end.month, end.day)
    today_last_year = date(today.year - 1, today.month, today.day)
    if prev_end > today_last_year:
        prev_end = today_last_year
    previous = Range(start=prev_start, end=prev_end)

    curr_df = _summary_df("daily", selected.start, selected.end, exchange, segment)
    prev_df = _summary_df("daily", previous.start, previous.end, exchange, segment)
    return curr_df, prev_df, selected, previous


def _sum_by_exchange(df: pd.DataFrame, keys: list[str]) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["exchange", *keys])
    have = [k for k in keys if k in df.columns]
    g = df.groupby("exchange", dropna=False)[have].sum(numeric_only=True).reset_index()
    for k in keys:
        if k not in g.columns:
            g[k] = 0.0
    return g[["exchange", *keys]]


def _pct_change(prev: float, curr: float) -> str:
    if abs(prev) > 1e-8:
        return f"{((curr - prev) / abs(prev) * 100):.2f}%"
    if abs(curr) > 1e-8:
        return "∞"
    return "-"


def main() -> None:
    st.set_page_config(page_title="Exchange Data Viewer", layout="wide")
    st.title("Exchange Data Viewer")
    st.caption("Excel-backed dashboard for NSE/BSE cash + derivatives and FII/DII flows.")

    today = date.today()
    latest_info_date = _last_market_day(today)

    if "from_date" not in st.session_state or "to_date" not in st.session_state:
        st.session_state["from_date"] = latest_info_date
        st.session_state["to_date"] = latest_info_date

    with st.sidebar:
        st.subheader("Filters")
        exchange: Exchange = st.selectbox("Exchange", ["all", "NSE", "BSE"], index=0)
        segment: Segment = st.selectbox("Segment", ["both", "cash", "fno"], index=0)
        agg: Agg = st.selectbox("Aggregation", ["daily", "weekly", "monthly", "quarterly", "yearly", "entire"], index=1)

        c1, c2 = st.columns(2)
        with c1:
            st.date_input("From", key="from_date")
        with c2:
            st.date_input("To", key="to_date")

        # Quick range buttons: two per row
        row1_col1, row1_col2 = st.columns(2)
        with row1_col1:
            st.button(
                "Last 4w",
                use_container_width=True,
                on_click=_set_preset_range,
                kwargs={"kind": "last4w", "anchor": latest_info_date},
            )
        with row1_col2:
            st.button(
                "YTD",
                use_container_width=True,
                on_click=_set_preset_range,
                kwargs={"kind": "ytd", "anchor": latest_info_date},
            )

        row2_col1, row2_col2 = st.columns(2)
        with row2_col1:
            st.button(
                "Last FY",
                use_container_width=True,
                on_click=_set_preset_range,
                kwargs={"kind": "lastfy", "anchor": latest_info_date},
            )
        with row2_col2:
            st.button(
                "FY YTD",
                use_container_width=True,
                on_click=_set_preset_range,
                kwargs={"kind": "fyytd", "anchor": latest_info_date},
            )

        st.divider()
        chart_segment = st.selectbox("Chart segment", ["cash", "futures", "options", "futures-options"], index=3)

    from_d: date = st.session_state["from_date"]
    to_d: date = st.session_state["to_date"]

    if agg == "entire":
        df = _entire_period_summary(from_d, to_d, exchange, segment)
    else:
        df = _summary_df(agg, from_d, to_d, exchange, segment)

    tab_summary, tab_table, tab_fii_dii, tab_comparison = st.tabs(["Summary", "Detailed Table", "FII/DII", "Comparison"])

    with tab_summary:
        if df.empty:
            st.info("No data found for the selected filters.")
        else:
            st.subheader("Segment summaries")
            tables = _segment_tables(df)
            if not tables:
                st.info("No segment totals available for the selected filters.")
            else:
                for title, tdf in tables.items():
                    st.markdown(f"**{title}**")
                    st.dataframe(tdf, width="stretch", hide_index=True)

            st.subheader("Trend (last 8 quarters)")
            qdf = _summary_df("quarterly", None, None, exchange, segment)
            if qdf.empty:
                st.info("No quarterly data available.")
            else:
                qdf = qdf.sort_values(["exchange", "period_sort"])
                qdf = qdf.groupby("exchange", dropna=False).tail(8)

                def avg_daily_turnover(d: pd.DataFrame) -> pd.Series:
                    days = pd.to_numeric(d.get("trading_days", 1), errors="coerce").fillna(1).replace(0, 1)
                    if chart_segment == "cash":
                        total = pd.to_numeric(d.get("cash_turnover_bn", 0), errors="coerce").fillna(0)
                    elif chart_segment == "futures":
                        total = pd.to_numeric(d.get("futures_turnover_bn", 0), errors="coerce").fillna(0)
                    elif chart_segment == "options":
                        total = pd.to_numeric(d.get("index_options_notional_bn", 0), errors="coerce").fillna(0) + pd.to_numeric(
                            d.get("equity_options_notional_bn", 0), errors="coerce"
                        ).fillna(0)
                    else:
                        total = (
                            pd.to_numeric(d.get("futures_turnover_bn", 0), errors="coerce").fillna(0)
                            + pd.to_numeric(d.get("index_options_notional_bn", 0), errors="coerce").fillna(0)
                            + pd.to_numeric(d.get("equity_options_notional_bn", 0), errors="coerce").fillna(0)
                        )
                    return (total / days).round(2)

                c1, c2 = st.columns(2)
                for ex, col in [("NSE", c1), ("BSE", c2)]:
                    with col:
                        exdf = qdf[qdf["exchange"] == ex].copy()
                        if exdf.empty:
                            st.info(f"No {ex} data.")
                            continue
                        exdf["Avg Daily Turnover (Rs bn)"] = avg_daily_turnover(exdf)
                        chart_df = exdf[["period_sort", "Avg Daily Turnover (Rs bn)"]].dropna(subset=["period_sort"]).copy()
                        if chart_df.empty:
                            st.info("No quarterly points to chart.")
                            continue

                        # Build quarter label as "Q# FY##" (financial year quarters),
                        # while keeping chronological ordering via period_sort.
                        m = chart_df["period_sort"].dt.month
                        y = chart_df["period_sort"].dt.year

                        # period_sort is quarter start date: Jan/Apr/Jul/Oct
                        qn = pd.Series(4, index=chart_df.index, dtype=int)  # Jan -> Q4
                        qn = qn.where(m.ne(4), 1)  # Apr -> Q1
                        qn = qn.where(m.ne(7), 2)  # Jul -> Q2
                        qn = qn.where(m.ne(10), 3)  # Oct -> Q3

                        fy_end_year = y.where(m.lt(4), y + 1)  # Jan-Mar belong to FY ending same year; else +1
                        chart_df["q_label"] = "Q" + qn.astype(str) + " FY" + (fy_end_year % 100).astype(str).str.zfill(2)

                        q_axis = alt.Axis(
                            title="Quarter",
                            labelAngle=90,
                            labelAlign="center",
                            labelBaseline="top",   # keep labels below the axis line
                            labelPadding=24,       # push labels further down
                            tickSize=4,
                        )

                        chart = (
                            alt.Chart(chart_df)
                            .mark_line(point=True)
                            .encode(
                                x=alt.X(
                                    "q_label:N",
                                    sort=alt.SortField(field="period_sort", order="ascending"),
                                    axis=q_axis,
                                ),
                                y=alt.Y("Avg Daily Turnover (Rs bn):Q", title="Avg Daily Turnover (Rs bn)"),
                                tooltip=[
                                    alt.Tooltip("q_label:N", title="Quarter"),
                                    alt.Tooltip("period_sort:T", title="Quarter start"),
                                    alt.Tooltip("Avg Daily Turnover (Rs bn):Q", format=",.2f"),
                                ],
                            )
                            .properties(height=320, title=f"{ex} - Avg Daily Turnover (Last 8 Quarters)")
                        )
                        st.altair_chart(chart, width="stretch")

    with tab_table:
        if df.empty:
            st.info("No rows to display.")
        else:
            display_names = {
                "period": "Dates",
                "exchange": "Exchange",
                "trading_days": "Trading Days",
                "cash_turnover_bn": "Cash Turnover (Rs bn)",
                "avg_cash_turnover_bn": "Cash Avg Daily Turnover (Rs bn)",
                "index_options_premium_bn": "Index Options Premium Turnover (Rs bn)",
                "avg_index_options_premium_bn": "Index Options Premium Avg Daily (Rs bn)",
                "equity_options_premium_bn": "Stock Options Premium Turnover (Rs bn)",
                "avg_equity_options_premium_bn": "Stock Options Premium Avg Daily (Rs bn)",
                "index_options_notional_bn": "Index Options Notional Turnover (Rs bn)",
                "avg_index_options_notional_bn": "Index Options Notional Avg Daily (Rs bn)",
                "equity_options_notional_bn": "Stock Options Notional Turnover (Rs bn)",
                "avg_equity_options_notional_bn": "Stock Options Notional Avg Daily (Rs bn)",
                "futures_turnover_bn": "Futures Turnover (Stock + Index) (Rs bn)",
                "avg_futures_turnover_bn": "Futures Avg Daily (Stock + Index) (Rs bn)",
                "index_options_volume": "Index Options Contracts",
                "equity_options_volume": "Stock Options Contracts",
                "index_futures_volume": "Index Futures Contracts",
                "equity_futures_volume": "Stock Futures Contracts",
                "total_contracts": "Total Contracts (F&O)",
            }

            show_cols = [
                "period",
                "exchange",
                "trading_days",
                "cash_turnover_bn",
                "avg_cash_turnover_bn",
                "index_options_premium_bn",
                "avg_index_options_premium_bn",
                "equity_options_premium_bn",
                "avg_equity_options_premium_bn",
                "index_options_notional_bn",
                "avg_index_options_notional_bn",
                "equity_options_notional_bn",
                "avg_equity_options_notional_bn",
                "futures_turnover_bn",
                "avg_futures_turnover_bn",
                "index_options_volume",
                "equity_options_volume",
                "index_futures_volume",
                "equity_futures_volume",
                "total_contracts",
            ]
            cols = [c for c in show_cols if c in df.columns]
            sort_cols = ["exchange"]
            if "period_sort" in df.columns:
                sort_cols = ["period_sort", "exchange"]
            elif "period" in df.columns:
                sort_cols = ["period", "exchange"]

            st.dataframe(
                df.sort_values(sort_cols, na_position="last")[cols].rename(columns={c: display_names.get(c, c.replace("_", " ").title()) for c in cols}),
                width="stretch",
                hide_index=True,
            )

    with tab_fii_dii:
        payload = get_fii_dii_data(_to_iso(from_d), _to_iso(to_d))
        rows = payload.get("rows", [])
        if not rows:
            st.info("No FII/DII data for the selected range.")
        else:
            fdf = pd.DataFrame(rows)
            if "date_sort" in fdf.columns:
                fdf["date_sort"] = pd.to_datetime(fdf["date_sort"], errors="coerce")
                fdf = fdf.sort_values("date_sort", ascending=False)
            cols = [
                "date",
                "fii_gross_purchase",
                "fii_gross_sales",
                "fii_net",
                "dii_gross_purchase",
                "dii_gross_sales",
                "dii_net",
            ]
            cols = [c for c in cols if c in fdf.columns]
            st.dataframe(fdf[cols], width="stretch", hide_index=True)

    with tab_comparison:
        st.subheader("Selected range vs previous year (daily totals)")
        if from_d is None or to_d is None:
            st.info("Select a date range first.")
        else:
            curr_df, prev_df, selected, previous = _comparison_payload(from_d, to_d, exchange, segment, today)
            if curr_df.empty or prev_df.empty:
                st.info("Not enough data to compute comparison for both ranges.")
            else:
                keys = [
                    "trading_days",
                    "cash_turnover_bn",
                    "index_options_premium_bn",
                    "equity_options_premium_bn",
                    "index_options_notional_bn",
                    "equity_options_notional_bn",
                    "futures_turnover_bn",
                    "index_options_volume",
                    "equity_options_volume",
                    "index_futures_volume",
                    "equity_futures_volume",
                    "total_contracts",
                ]
                curr = _sum_by_exchange(curr_df, keys).set_index("exchange")
                prev = _sum_by_exchange(prev_df, keys).set_index("exchange")

                metric_labels = {
                    "trading_days": "Trading Days",
                    "cash_turnover_bn": "Cash Turnover (Rs bn)",
                    "index_options_premium_bn": "Index Options Premium Turnover (Rs bn)",
                    "equity_options_premium_bn": "Stock Options Premium Turnover (Rs bn)",
                    "index_options_notional_bn": "Index Options Notional Turnover (Rs bn)",
                    "equity_options_notional_bn": "Stock Options Notional Turnover (Rs bn)",
                    "futures_turnover_bn": "Futures Turnover (Stock + Index) (Rs bn)",
                    "index_options_volume": "Index Options Total Contracts",
                    "equity_options_volume": "Stock Options Total Contracts",
                    "index_futures_volume": "Index Futures Total Contracts",
                    "equity_futures_volume": "Stock Futures Total Contracts",
                    "total_contracts": "Total Contracts (F&O)",
                }

                def build_table(ex: str) -> pd.DataFrame:
                    rows_out = []
                    for k in keys:
                        pv = float(prev.loc[ex, k]) if ex in prev.index else 0.0
                        cv = float(curr.loc[ex, k]) if ex in curr.index else 0.0
                        rows_out.append(
                            {
                                "Metric": metric_labels.get(k, k.replace("_", " ").title()),
                                f"{previous.start} → {previous.end}": pv,
                                f"{selected.start} → {selected.end}": cv,
                                "% Change": _pct_change(pv, cv),
                            }
                        )
                    out = pd.DataFrame(rows_out)
                    return out

                st.markdown("**BSE**")
                st.dataframe(build_table("BSE"), width="stretch", hide_index=True)
                st.divider()
                st.markdown("**NSE**")
                st.dataframe(build_table("NSE"), width="stretch", hide_index=True)


if __name__ == "__main__":
    main()

