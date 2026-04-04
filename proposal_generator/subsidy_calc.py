"""
subsidy_calc.py - Subsidy auto-calculation engine

Calculates applicable subsidy amounts based on system specs and pricing.
Based on the Excel "補助金" sheet structure.
"""
from __future__ import annotations

import math


def calc_pv_output(panel_kw: float, pcs_kw: float) -> float:
    """発電出力 = MIN(パネルkW, PCS kW), rounded down to integer."""
    if panel_kw <= 0 or pcs_kw <= 0:
        return 0.0
    return math.floor(min(panel_kw, pcs_kw))


def calc_all_subsidies(
    panel_kw: float,
    pcs_kw: float,
    selling_price: float,
    battery_price: float,
    battery_kwh: float,
    company_size: str,
    proposal_type: str,
    cost_ratio: float = 0.663,
) -> list[dict]:
    """Calculate all applicable subsidies.

    Args:
        panel_kw: Total panel output (kW)
        pcs_kw: Total PCS output (kW)
        selling_price: Equipment selling price (yen)
        battery_price: Battery selling price (yen)
        battery_kwh: Total battery capacity (kWh)
        company_size: "大企業" or "中小企業" or ""
        proposal_type: "ppa" or "epc"
        cost_ratio: Cost ratio for Tokyo profit exclusion (default ~66.3%)

    Returns:
        List of dicts with name, amount, notes, applicable
    """
    pv_output = calc_pv_output(panel_kw, pcs_kw)
    has_battery = battery_kwh > 0
    results: list[dict] = []

    # --- 1. 環境省 - ストレージパリティ (蓄電池必須) ---
    if has_battery:
        # PV subsidy: PPA=50,000/kW, EPC=40,000/kW, capped at 20M
        pv_unit = 40_000 if proposal_type == "epc" else 50_000
        pv_kw = min(panel_kw, pcs_kw) if panel_kw > 0 and pcs_kw > 0 else 0
        pv_amount = min(pv_kw * pv_unit, 20_000_000)

        # Battery subsidy: 45,000/kWh (industrial >20kWh: 40,000/kWh)
        bat_unit = 40_000 if battery_kwh > 20 else 45_000
        bat_calc = battery_kwh * bat_unit
        # Cap: min(calculated, battery_price * 1/3), floor to 1000 yen
        bat_cap = battery_price / 3 if battery_price > 0 else float("inf")
        bat_amount = math.floor(min(bat_calc, bat_cap) / 1000) * 1000
        # Battery subsidy capped at 10M
        bat_amount = min(bat_amount, 10_000_000)

        # Total capped at 30M (PV 20M + ESS 10M)
        total = min(pv_amount + bat_amount, 30_000_000)
        results.append({
            "name": "環境省（ストレージパリティ）",
            "amount": int(total),
            "notes": "蓄電池必須。費用効率性40,000円/t-CO2以下の制約あり",
            "applicable": True,
        })
    else:
        results.append({
            "name": "環境省（ストレージパリティ）",
            "amount": 0,
            "notes": "蓄電池が必要です",
            "applicable": False,
        })

    # --- 2. 東京都 (利益排除方式) ---
    if company_size:
        deemed_pv = (selling_price - battery_price) * cost_ratio if selling_price > 0 else 0
        deemed_ess = battery_price * cost_ratio if battery_price > 0 else 0

        if company_size == "中小企業":
            pv_a = min(pv_output * 200_000, deemed_pv * 2 / 3) if deemed_pv > 0 else pv_output * 200_000
            bat_a = min(battery_kwh * 150_000, deemed_ess * 3 / 4) if has_battery and deemed_ess > 0 else (battery_kwh * 150_000 if has_battery else 0)
        else:  # 大企業
            pv_a = min(pv_output * 150_000, deemed_pv / 2) if deemed_pv > 0 else pv_output * 150_000
            bat_a = min(battery_kwh * 130_000, deemed_ess * 2 / 3) if has_battery and deemed_ess > 0 else (battery_kwh * 130_000 if has_battery else 0)

        tokyo_total = min(int(pv_a + bat_a), 200_000_000)
        results.append({
            "name": "東京都",
            "amount": tokyo_total,
            "notes": f"利益排除方式（原価率{cost_ratio*100:.1f}%）。{company_size}向け",
            "applicable": True,
        })
    else:
        results.append({
            "name": "東京都",
            "amount": 0,
            "notes": "企業規模を選択してください",
            "applicable": False,
        })

    # --- 3. 神奈川県 ---
    pv_a = pv_output * 80_000
    bat_a = battery_kwh * 50_000 if has_battery else 0
    kanagawa = int(pv_a + bat_a)
    if company_size == "大企業":
        kanagawa = min(kanagawa, 30_000_000)
    results.append({
        "name": "神奈川県",
        "amount": kanagawa,
        "notes": "大企業は上限3,000万円" if company_size == "大企業" else "",
        "applicable": True,
    })

    # --- 4. 埼玉県① (EPC専用, 中小企業) ---
    if proposal_type == "epc" and company_size == "中小企業":
        saitama1 = int(min(selling_price / 2, 5_000_000)) if selling_price > 0 else 0
        results.append({
            "name": "埼玉県①（CO2排出削減）",
            "amount": saitama1,
            "notes": "EPC専用。補助率1/2以内、上限500万円",
            "applicable": True,
        })

    # --- 5. 埼玉県② ---
    pv_a = pv_output * 50_000
    bat_a = battery_price / 3 if battery_price > 0 and has_battery else 0
    saitama2 = int(min(pv_a + bat_a, 15_000_000))
    results.append({
        "name": "埼玉県②",
        "amount": saitama2,
        "notes": "上限1,500万円",
        "applicable": True,
    })

    # --- 6. 群馬県 (中小企業のみ) ---
    if company_size == "中小企業":
        pv_a = math.floor(min(panel_kw, pcs_kw) * 10) / 10 * 50_000
        bat_a = battery_price / 3 if battery_price > 0 and has_battery else 0
        cap = 15_000_000 if has_battery else 5_000_000
        gunma = int(min(pv_a + bat_a, cap))
        results.append({
            "name": "群馬県",
            "amount": gunma,
            "notes": f"中小企業のみ。上限{cap//10000:,}万円",
            "applicable": True,
        })

    # --- 7. 静岡県 (中小企業のみ) ---
    if company_size == "中小企業":
        pv_a = pv_output * 40_000
        bat_c1 = battery_kwh * 53_000
        bat_c2 = battery_price / 3 if battery_price > 0 else float("inf")
        bat_a = min(bat_c1, bat_c2) if has_battery else 0
        shizuoka = int(pv_a + bat_a)
        results.append({
            "name": "静岡県",
            "amount": shizuoka,
            "notes": "中小企業のみ",
            "applicable": True,
        })

    return results
