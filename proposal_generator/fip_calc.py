"""
fip_calc.py - FIP (Feed-in Premium) revenue calculation engine

Calculates FIP revenue from surplus electricity sold under the
Feed-in Premium scheme, including balancing cost estimation.

FIP revenue = (Market price + Premium) * Surplus kWh - Balancing cost

Key concepts:
  - FIP Premium: Fixed premium added to market electricity price (yen/kWh)
  - Market price: JEPX spot market average (yen/kWh)
  - Balancing cost: Penalty for imbalance between forecast and actual generation
"""

from __future__ import annotations


# ---------------------------------------------------------------------------
# Default parameters
# ---------------------------------------------------------------------------

DEFAULT_MARKET_PRICE = 12.0        # JEPX average (yen/kWh), approx 2024-2025
DEFAULT_BALANCING_RATE = 1.0       # yen/kWh, typical balancing cost estimate


# ---------------------------------------------------------------------------
# FIP revenue calculation
# ---------------------------------------------------------------------------

def calc_fip_revenue(
    surplus_kwh: float,
    premium_yen_per_kwh: float,
    market_price_yen_per_kwh: float = DEFAULT_MARKET_PRICE,
) -> float:
    """Calculate annual FIP gross revenue (before balancing cost).

    FIP revenue = surplus_kwh * (market_price + premium)

    Args:
        surplus_kwh:              Annual surplus electricity (kWh)
        premium_yen_per_kwh:      FIP premium unit price (yen/kWh)
        market_price_yen_per_kwh: Expected JEPX market price (yen/kWh)

    Returns:
        Annual FIP gross revenue (yen)
    """
    if surplus_kwh <= 0 or premium_yen_per_kwh < 0:
        return 0.0
    return surplus_kwh * (market_price_yen_per_kwh + premium_yen_per_kwh)


def calc_balancing_cost(
    surplus_kwh: float,
    rate: float = DEFAULT_BALANCING_RATE,
) -> float:
    """Estimate annual balancing cost (imbalance penalty).

    A simplified estimate: flat rate per kWh of surplus generation.
    In practice, balancing cost depends on forecast accuracy and
    imbalance settlement rules.

    Args:
        surplus_kwh: Annual surplus electricity (kWh)
        rate:        Balancing cost rate (yen/kWh), default 1.0

    Returns:
        Estimated annual balancing cost (yen)
    """
    if surplus_kwh <= 0 or rate < 0:
        return 0.0
    return surplus_kwh * rate


def calc_fip_net_revenue(
    surplus_kwh: float,
    premium_yen_per_kwh: float,
    market_price_yen_per_kwh: float = DEFAULT_MARKET_PRICE,
    balancing_rate: float = DEFAULT_BALANCING_RATE,
) -> dict:
    """Calculate FIP net revenue with breakdown.

    Args:
        surplus_kwh:              Annual surplus electricity (kWh)
        premium_yen_per_kwh:      FIP premium unit price (yen/kWh)
        market_price_yen_per_kwh: Expected JEPX market price (yen/kWh)
        balancing_rate:           Balancing cost rate (yen/kWh)

    Returns:
        Dict with:
          - gross_revenue: FIP gross revenue (yen)
          - balancing_cost: Estimated balancing cost (yen)
          - net_revenue: FIP net revenue (yen)
          - effective_rate: Effective sell rate after balancing (yen/kWh)
          - surplus_kwh: Input surplus (kWh)
    """
    gross = calc_fip_revenue(surplus_kwh, premium_yen_per_kwh, market_price_yen_per_kwh)
    balancing = calc_balancing_cost(surplus_kwh, balancing_rate)
    net = gross - balancing
    effective = net / surplus_kwh if surplus_kwh > 0 else 0.0

    return {
        "gross_revenue": round(gross),
        "balancing_cost": round(balancing),
        "net_revenue": round(net),
        "effective_rate": round(effective, 2),
        "surplus_kwh": round(surplus_kwh),
    }


def calc_self_consumption_vs_fip(
    self_consumption_kwh: float,
    surplus_kwh: float,
    electricity_unit_price: float,
    premium_yen_per_kwh: float,
    market_price_yen_per_kwh: float = DEFAULT_MARKET_PRICE,
    balancing_rate: float = DEFAULT_BALANCING_RATE,
) -> dict:
    """Compare self-consumption savings vs FIP revenue.

    Args:
        self_consumption_kwh:     Annual self-consumption (kWh)
        surplus_kwh:              Annual surplus electricity (kWh)
        electricity_unit_price:   Current electricity purchase price (yen/kWh)
        premium_yen_per_kwh:      FIP premium (yen/kWh)
        market_price_yen_per_kwh: Expected market price (yen/kWh)
        balancing_rate:           Balancing cost rate (yen/kWh)

    Returns:
        Dict with comparison data for the FIP slide.
    """
    self_consumption_saving = self_consumption_kwh * electricity_unit_price

    fip = calc_fip_net_revenue(
        surplus_kwh, premium_yen_per_kwh,
        market_price_yen_per_kwh, balancing_rate,
    )

    total_benefit = self_consumption_saving + fip["net_revenue"]

    return {
        "self_consumption_kwh": round(self_consumption_kwh),
        "self_consumption_saving": round(self_consumption_saving),
        "fip_net_revenue": fip["net_revenue"],
        "fip_gross_revenue": fip["gross_revenue"],
        "fip_balancing_cost": fip["balancing_cost"],
        "fip_effective_rate": fip["effective_rate"],
        "total_annual_benefit": round(total_benefit),
        "surplus_kwh": round(surplus_kwh),
    }
