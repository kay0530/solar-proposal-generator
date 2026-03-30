"""
demand_calc.py - Demand cut calculation engine

Calculates peak demand reduction from iPals hourly data and produces
chart-ready 2-week window data for PP9/EP5 slides.
"""

from __future__ import annotations


def calc_demand_cut(
    hourly_rows: list[dict],
    basic_rate_kw: float = 0.0,
    power_factor_pct: int = 85,
) -> dict:
    """Calculate demand cut metrics from iPals hourly data.

    Args:
        hourly_rows: list of dicts with keys:
            month, day, hour, demand_kw, gen_kw, self_consumption_kw, surplus_kw
        basic_rate_kw: basic charge unit price (yen/kW/month) from electricity master
        power_factor_pct: power factor percentage (default 85)

    Returns:
        dict with peak values, savings, and chart data
    """
    if not hourly_rows:
        return {}

    # --- Peak detection ---
    peak_before_kw = 0.0
    peak_before_idx = 0
    peak_after_kw = 0.0
    peak_after_idx = 0

    for i, row in enumerate(hourly_rows):
        demand = float(row.get("demand_kw", 0) or 0)
        self_c = float(row.get("self_consumption_kw", 0) or 0)
        net_demand = demand - self_c

        if demand > peak_before_kw:
            peak_before_kw = demand
            peak_before_idx = i

        if net_demand > peak_after_kw:
            peak_after_kw = net_demand
            peak_after_idx = i

    demand_cut_kw = peak_before_kw - peak_after_kw

    peak_before_month = hourly_rows[peak_before_idx].get("month", 0)
    peak_after_month = hourly_rows[peak_after_idx].get("month", 0)

    # --- Basic fee calculation ---
    # Formula: basic_rate * peak_kw * (185 - PF%) / 100
    # At PF=85%: factor=1.00, PF=100%: factor=0.85
    pf_factor = (185 - power_factor_pct) / 100
    monthly_basic_before = basic_rate_kw * peak_before_kw * pf_factor
    monthly_basic_after = basic_rate_kw * peak_after_kw * pf_factor
    monthly_basic_saving = monthly_basic_before - monthly_basic_after
    annual_basic_saving = monthly_basic_saving * 12

    # --- 2-week chart window (centered on peak_before day) ---
    peak_week_before, peak_week_after = _select_peak_weeks(
        hourly_rows, peak_before_idx
    )

    return {
        "peak_before_kw": round(peak_before_kw, 1),
        "peak_after_kw": round(peak_after_kw, 1),
        "demand_cut_kw": round(demand_cut_kw, 1),
        "peak_before_month": peak_before_month,
        "peak_after_month": peak_after_month,
        "basic_rate_kw": basic_rate_kw,
        "power_factor_pct": power_factor_pct,
        "pf_factor": round(pf_factor, 4),
        "monthly_basic_before": round(monthly_basic_before),
        "monthly_basic_after": round(monthly_basic_after),
        "monthly_basic_saving": round(monthly_basic_saving),
        "annual_basic_saving": round(annual_basic_saving),
        "peak_week_before": peak_week_before,
        "peak_week_after": peak_week_after,
    }


def _select_peak_weeks(
    hourly_rows: list[dict], peak_idx: int
) -> tuple[list[dict], list[dict]]:
    """Select 14-day (336-hour) window centered on the peak demand hour.

    Returns (before_window, after_window) where:
      - before_window: raw demand values for chart
      - after_window: net demand (demand - self_consumption) for chart
    """
    n = len(hourly_rows)
    half_window = 7 * 24  # 7 days in hours

    # Find the start of the peak day (hour 1)
    peak_row = hourly_rows[peak_idx]
    peak_day_start = peak_idx - (int(peak_row.get("hour", 1)) - 1)
    peak_day_start = max(0, peak_day_start)

    # Center window: 7 days before peak day, 7 days after (inclusive of peak day)
    window_start = peak_day_start - half_window
    window_end = peak_day_start + half_window + 24  # +24 for peak day itself

    # Clamp to data bounds
    if window_start < 0:
        window_start = 0
        window_end = min(n, 14 * 24)
    if window_end > n:
        window_end = n
        window_start = max(0, n - 14 * 24)

    window = hourly_rows[window_start:window_end]

    before_window = []
    after_window = []
    for row in window:
        demand = float(row.get("demand_kw", 0) or 0)
        self_c = float(row.get("self_consumption_kw", 0) or 0)
        label = f"{int(row.get('month', 0))}/{int(row.get('day', 0))} {int(row.get('hour', 0))}:00"
        before_window.append({"label": label, "value": demand})
        after_window.append({"label": label, "value": max(0, demand - self_c)})

    return before_window, after_window
