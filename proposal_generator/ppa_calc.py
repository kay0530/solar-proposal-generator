"""
ppa_calc.py - PPA unit price auto-calculation engine

Calculates the minimum PPA unit price such that DSCR >= 1.30
for all years in the financing period.

Supports two financing types:
  - Lease (リース): fixed annual payments via PMT annuity formula
  - Bank loan (銀行融資): equal principal repayment (元金均等返済)
    with declining interest, plus fire insurance and depreciation tax

Key assumptions:
  - Generation & self-consumption degrade 0.5%/year
  - Surplus revenue: only if enabled (normally 0 due to RPR)
  - DSCR = Annual PPA Revenue / Annual Total Cost
  - For lease: worst DSCR is in the final year (fixed cost, declining generation)
  - For bank loan: worst DSCR could be any year (costs decline but so does generation)

Finance rate defaults:
  - シーエナジー (CE): IRR = 3.10% (their minimum)
  - みずほリース:      5.50% fixed
  - 群馬銀行:          1.80% bank loan
"""

from __future__ import annotations

import math


# ---------------------------------------------------------------------------
# Finance type constants
# ---------------------------------------------------------------------------

FINANCE_TYPE_LEASE = "lease"
FINANCE_TYPE_LOAN = "loan"

# Map company names to finance types
FINANCE_TYPE_MAP: dict[str, str] = {
    "シーエナジー": FINANCE_TYPE_LEASE,
    "みずほリース": FINANCE_TYPE_LEASE,
    "群馬銀行": FINANCE_TYPE_LOAN,
}

# Default rates per company (replaces old LEASE_RATE_MAP)
DEFAULT_RATE_MAP: dict[str, float] = {
    "シーエナジー": 0.0310,   # CE target IRR = 3.10%
    "みずほリース": 0.0550,   # Fixed 5.5%
    "群馬銀行": 0.0180,       # Bank loan 1.8%
}

# Keep backward compatibility alias
LEASE_RATE_MAP = DEFAULT_RATE_MAP

# Revenue share (シーエナジー: post-lease revenue split)
CE_REVENUE_SHARE_RATE = 0.0  # Currently 0% per PPAリース U14 default

# Re-lease (みずほリース: 1/10 of original lease for years after lease term)
MIZUHO_RELEASE_RATIO = 0.1  # 再リース = original_lease * 10%

DEGRADATION_RATE = 0.005  # 0.5% per year

# Default O&M costs (from PPAリース sheet)
DEFAULT_MAINTENANCE_YEN_PER_KW = 1_200   # 保守メンテナンス費 (円/kW/年)
DEFAULT_INSURANCE_YEN_FIXED = 120_000    # 保安管理業務委託費 (円/年・自社負担時)

# Depreciation tax constants (for solar, 17-year useful life)
DEPRECIATION_RATE_R = 0.127        # 減価率 r (for fixed asset tax assessment)
DEPRECIATION_TAX_RATE = 0.014      # 償却資産税率 1.4%

# Fire insurance
FIRE_INSURANCE_PER_MILLION = 3107  # 3,107 yen per million yen of equipment


# ---------------------------------------------------------------------------
# Financial helpers
# ---------------------------------------------------------------------------

def pmt(rate: float, nper: int, pv: float) -> float:
    """Annual lease payment (annuity formula).

    Args:
        rate: Annual interest rate (e.g. 0.031 for 3.1%)
        nper: Number of periods (years)
        pv:   Present value = principal (positive number)

    Returns:
        Annual payment (positive = outflow from lessee perspective)
    """
    if rate == 0:
        return pv / nper
    return pv * rate / (1 - (1 + rate) ** (-nper))


def npv(rate: float, cashflows: list[float]) -> float:
    """Net Present Value of a cash flow series.

    cashflows[0] is at t=1, cashflows[1] at t=2, etc.
    """
    return sum(cf / (1 + rate) ** (i + 1) for i, cf in enumerate(cashflows))


def irr(cashflows: list[float], guess: float = 0.05, tol: float = 1e-7, max_iter: int = 200) -> float:
    """Internal Rate of Return via Newton-Raphson.

    cashflows[0] = initial investment (negative), cashflows[1..] = annual inflows.
    """
    r = guess
    for _ in range(max_iter):
        f = sum(cf / (1 + r) ** i for i, cf in enumerate(cashflows))
        df = sum(-i * cf / (1 + r) ** (i + 1) for i, cf in enumerate(cashflows))
        if df == 0:
            break
        r_new = r - f / df
        if abs(r_new - r) < tol:
            return r_new
        r = r_new
    return r


# ---------------------------------------------------------------------------
# Bank loan specific helpers
# ---------------------------------------------------------------------------

def calc_fire_insurance_annual(selling_price: float) -> int:
    """Calculate annual fire insurance for bank loan (群馬銀行).

    Formula: ceil(selling_price / 1,000,000 * 3,107 / 1000) * 1000
    Rounds UP to nearest 1,000 yen.

    Args:
        selling_price: Equipment selling price (yen)

    Returns:
        Annual fire insurance amount (yen), rounded up to nearest 1,000
    """
    return math.ceil(selling_price / 1_000_000 * FIRE_INSURANCE_PER_MILLION / 1000) * 1000


def calc_depreciation_tax_schedule(selling_price: float, term_years: int) -> list[int]:
    """Calculate yearly depreciation tax (償却資産税) schedule for bank loan.

    Year 1 assessed value = selling_price * (1 - r/2) = price * 0.9365
    Year 2+ assessed value = prior * (1 - r) = prior * 0.873
    Assessed value is truncated to nearest 1,000 yen.
    Annual tax = truncated_assessed * 1.4%

    Args:
        selling_price: Equipment selling price (yen)
        term_years:    Financing term (years)

    Returns:
        List of annual depreciation tax amounts (length = term_years)
    """
    schedule: list[int] = []
    assessed = selling_price * (1 - DEPRECIATION_RATE_R / 2)

    for year in range(1, term_years + 1):
        if year >= 2:
            assessed = assessed * (1 - DEPRECIATION_RATE_R)

        # Truncate to nearest 1,000 yen
        truncated = int(assessed // 1000) * 1000
        tax = truncated * DEPRECIATION_TAX_RATE
        schedule.append(round(tax))

    return schedule


def calc_bank_loan_annual_payments(
    principal: float,
    annual_rate: float,
    term_years: int,
) -> list[dict]:
    """Calculate bank loan annual payment schedule (元金均等返済).

    Equal principal repayment each month, with declining interest.

    Args:
        principal:    Loan principal (yen)
        annual_rate:  Annual interest rate (e.g. 0.018 for 1.8%)
        term_years:   Loan term (years)

    Returns:
        List of dicts: [{"year": y, "principal": p, "interest": i, "total": p+i}, ...]
    """
    total_months = term_years * 12
    monthly_principal = principal / total_months
    monthly_rate = annual_rate / 12

    schedule: list[dict] = []
    for year in range(1, term_years + 1):
        year_principal = 0.0
        year_interest = 0.0

        for month_in_year in range(12):
            # Global month index (0-based)
            m = (year - 1) * 12 + month_in_year
            remaining = principal - monthly_principal * m
            interest = remaining * monthly_rate

            year_principal += monthly_principal
            year_interest += interest

        schedule.append({
            "year": year,
            "principal": round(year_principal),
            "interest": round(year_interest),
            "total": round(year_principal + year_interest),
        })

    return schedule


# ---------------------------------------------------------------------------
# Finance type detection
# ---------------------------------------------------------------------------

def get_finance_type(company: str) -> str:
    """Determine finance type from company name.

    Args:
        company: Finance company name

    Returns:
        FINANCE_TYPE_LEASE or FINANCE_TYPE_LOAN
    """
    return FINANCE_TYPE_MAP.get(company, FINANCE_TYPE_LEASE)


# ---------------------------------------------------------------------------
# Main calculation
# ---------------------------------------------------------------------------

def calc_lease_payment(
    principal: float,
    lease_company: str,
    lease_rate_pct: float,
    lease_years: int,
) -> tuple[float, float]:
    """Calculate annual lease payment and effective rate.

    For lease companies: uses PMT annuity formula.
    For bank loan companies: returns year-1 total payment as the representative
    annual payment (note: actual payments vary by year for bank loans).

    Args:
        principal:      System cost net of subsidy (yen)
        lease_company:  Finance company name
        lease_rate_pct: User-specified rate (%) -- used only for unknown companies
        lease_years:    Finance term (years)

    Returns:
        (annual_payment, effective_rate)
    """
    # Determine rate
    rate = DEFAULT_RATE_MAP.get(lease_company)
    if rate is None:
        rate = lease_rate_pct / 100.0

    finance_type = get_finance_type(lease_company)

    if finance_type == FINANCE_TYPE_LOAN:
        # For bank loan, return year-1 payment as representative
        loan_schedule = calc_bank_loan_annual_payments(principal, rate, lease_years)
        annual_payment = loan_schedule[0]["total"] if loan_schedule else 0.0
    else:
        annual_payment = pmt(rate, lease_years, principal)

    return annual_payment, rate


def calc_annual_om_cost(
    system_kw: float,
    maintenance_yen_per_kw: float = DEFAULT_MAINTENANCE_YEN_PER_KW,
    insurance_yen_fixed: float = DEFAULT_INSURANCE_YEN_FIXED,
) -> float:
    """Calculate annual O&M cost (保守費 + 保険料).

    Args:
        system_kw:              System capacity (kW) = min(panel_kw, pcs_kw)
        maintenance_yen_per_kw: Maintenance fee per kW per year
        insurance_yen_fixed:    Fixed annual insurance/management fee

    Returns:
        Annual O&M cost (yen)
    """
    return system_kw * maintenance_yen_per_kw + insurance_yen_fixed


def calc_min_ppa_price(
    self_consumption_y1_kwh: float,
    surplus_y1_kwh: float,
    annual_lease_payment: float,
    lease_years: int,
    fit_price: float = 0.0,
    include_surplus: bool = False,
    target_dscr: float = 1.30,
    degradation: float = DEGRADATION_RATE,
    annual_om_cost: float = 0.0,
    finance_type: str = FINANCE_TYPE_LEASE,
    loan_payment_schedule: list[dict] | None = None,
    fire_insurance_annual: int = 0,
    depreciation_tax_schedule: list[int] | None = None,
) -> float:
    """Calculate minimum PPA unit price to achieve target_dscr in all years.

    For lease: DSCR = Revenue / (Lease Payment + O&M)
      - Worst year is always the final year (fixed cost, declining generation)

    For bank loan: DSCR = Revenue / (Loan Payment + O&M + Insurance + DepreciationTax)
      - Must check ALL years since costs decline but generation also declines

    Args:
        self_consumption_y1_kwh: Year-1 self-consumption (kWh)
        surplus_y1_kwh:          Year-1 surplus electricity (kWh)
        annual_lease_payment:    Fixed annual lease payment (yen) -- used for lease only
        lease_years:             Finance term
        fit_price:               FIT price for surplus (yen/kWh)
        include_surplus:         Whether to include surplus revenue
        target_dscr:             Minimum acceptable DSCR
        degradation:             Annual degradation rate (0.005 = 0.5%)
        annual_om_cost:          Annual O&M cost (yen) -- added to denominator
        finance_type:            FINANCE_TYPE_LEASE or FINANCE_TYPE_LOAN
        loan_payment_schedule:   Year-by-year bank loan payments (for LOAN only)
        fire_insurance_annual:   Annual fire insurance (for LOAN only)
        depreciation_tax_schedule: Year-by-year depreciation tax (for LOAN only)

    Returns:
        Minimum PPA unit price (yen/kWh), rounded up to nearest 0.5 yen
    """
    if self_consumption_y1_kwh <= 0 or annual_lease_payment <= 0:
        # For bank loan, check loan schedule instead
        if finance_type == FINANCE_TYPE_LOAN and loan_payment_schedule:
            pass  # proceed with loan schedule
        elif self_consumption_y1_kwh <= 0:
            return 0.0
        else:
            return 0.0

    if finance_type == FINANCE_TYPE_LOAN and loan_payment_schedule:
        # Bank loan: check every year to find worst DSCR scenario
        dep_schedule = depreciation_tax_schedule or [0] * lease_years
        max_required_price = 0.0

        for year in range(1, lease_years + 1):
            decay = (1 - degradation) ** (year - 1)
            sc = self_consumption_y1_kwh * decay
            sur = surplus_y1_kwh * decay if include_surplus else 0.0

            # Year's total cost
            loan_pay = loan_payment_schedule[year - 1]["total"] if year <= len(loan_payment_schedule) else 0
            dep_tax = dep_schedule[year - 1] if year <= len(dep_schedule) else 0
            total_cost = loan_pay + annual_om_cost + fire_insurance_annual + dep_tax

            required_revenue = total_cost * target_dscr
            surplus_revenue = sur * fit_price
            required_ppa_revenue = required_revenue - surplus_revenue

            if required_ppa_revenue <= 0 or sc <= 0:
                continue

            price_this_year = required_ppa_revenue / sc
            if price_this_year > max_required_price:
                max_required_price = price_this_year

        if max_required_price <= 0:
            return 0.0

        # Round up to nearest 0.5 yen
        return math.ceil(max_required_price * 2) / 2

    else:
        # Lease: worst year is final year (fixed cost, declining generation)
        decay = (1 - degradation) ** (lease_years - 1)
        min_self_consume = self_consumption_y1_kwh * decay
        min_surplus = surplus_y1_kwh * decay if include_surplus else 0.0

        total_cost = annual_lease_payment + annual_om_cost
        required_revenue = total_cost * target_dscr
        surplus_revenue = min_surplus * fit_price
        required_ppa_revenue = required_revenue - surplus_revenue

        if required_ppa_revenue <= 0:
            return 0.0

        raw_price = required_ppa_revenue / min_self_consume

        # Round up to nearest 0.5 yen
        return math.ceil(raw_price * 2) / 2


def calc_cashflow_table(
    self_consumption_y1_kwh: float,
    surplus_y1_kwh: float,
    ppa_unit_price: float,
    fit_price: float,
    annual_lease_payment: float,
    contract_years: int,
    lease_years: int,
    include_surplus: bool = False,
    degradation: float = DEGRADATION_RATE,
    annual_om_cost: float = 0.0,
    finance_type: str = FINANCE_TYPE_LEASE,
    loan_payment_schedule: list[dict] | None = None,
    fire_insurance_annual: int = 0,
    depreciation_tax_schedule: list[int] | None = None,
    company: str = "",
) -> list[dict]:
    """Generate year-by-year cashflow table for the PPA period.

    For lease: total_cost = lease_payment + om_cost (constant)
    For bank loan: total_cost = loan_payment[year] + om_cost + fire_insurance + depreciation_tax[year]

    DSCR = Revenue / Total Cost

    Returns list of dicts with keys:
        year, self_consumption_kwh, surplus_kwh, ppa_revenue, surplus_revenue,
        total_revenue, lease_payment, om_cost, fire_insurance, depreciation_tax,
        total_cost, net_cashflow, dscr
    """
    rows = []
    dep_schedule = depreciation_tax_schedule or []

    for year in range(1, contract_years + 1):
        decay = (1 - degradation) ** (year - 1)
        sc = self_consumption_y1_kwh * decay
        sur = surplus_y1_kwh * decay if include_surplus else 0.0

        ppa_rev = sc * ppa_unit_price
        sur_rev = sur * fit_price
        total_rev = ppa_rev + sur_rev

        # Determine year's costs based on finance type
        if year <= lease_years:
            om = annual_om_cost

            if finance_type == FINANCE_TYPE_LOAN and loan_payment_schedule:
                lp = loan_payment_schedule[year - 1]["total"] if year <= len(loan_payment_schedule) else 0
                fi = fire_insurance_annual
                dt = dep_schedule[year - 1] if year <= len(dep_schedule) else 0
            else:
                lp = annual_lease_payment
                fi = 0
                dt = 0
        else:
            # After finance term expires
            om = 0.0
            fi = 0
            dt = 0
            # Re-lease for みずほリース
            if finance_type == FINANCE_TYPE_LEASE and company == "みずほリース":
                lp = annual_lease_payment * MIZUHO_RELEASE_RATIO
            else:
                lp = 0.0

        total_cost = lp + om + fi + dt

        # Revenue share for シーエナジー (post-lease only)
        revenue_share = 0.0
        if year > lease_years and company == "シーエナジー" and CE_REVENUE_SHARE_RATE > 0:
            revenue_share = (total_rev - om) * CE_REVENUE_SHARE_RATE
            total_cost += revenue_share

        net_cf = total_rev - total_cost
        dscr = total_rev / total_cost if total_cost > 0 else float("inf")

        row = {
            "year": year,
            "self_consumption_kwh": round(sc),
            "surplus_kwh": round(sur),
            "ppa_revenue": round(ppa_rev),
            "surplus_revenue": round(sur_rev),
            "total_revenue": round(total_rev),
            "lease_payment": round(lp),
            "om_cost": round(om),
            "total_cost": round(total_cost),
            "net_cashflow": round(net_cf),
            "dscr": round(dscr, 3) if total_cost > 0 else None,
        }

        # Add bank loan specific fields
        if finance_type == FINANCE_TYPE_LOAN:
            row["fire_insurance"] = round(fi)
            row["depreciation_tax"] = round(dt)
            if loan_payment_schedule and year <= len(loan_payment_schedule):
                row["loan_principal"] = loan_payment_schedule[year - 1]["principal"]
                row["loan_interest"] = loan_payment_schedule[year - 1]["interest"]

        rows.append(row)
    return rows


def auto_calc_ppa(
    self_consumption_y1_kwh: float,
    surplus_y1_kwh: float,
    selling_price: float,
    subsidy_amount: float,
    lease_company: str,
    lease_rate_pct: float,
    lease_years: int,
    contract_years: int,
    system_kw: float = 0.0,
    fit_price: float = 0.0,
    include_surplus: bool = False,
    target_dscr: float = 1.30,
    maintenance_yen_per_kw: float = DEFAULT_MAINTENANCE_YEN_PER_KW,
    insurance_yen_fixed: float = DEFAULT_INSURANCE_YEN_FIXED,
    finance_company: str | None = None,
) -> dict:
    """Full PPA auto-calculation: financing payment -> O&M -> minimum PPA price -> cashflow table.

    Supports both lease and bank loan financing. The finance type is determined
    from the company name via FINANCE_TYPE_MAP.

    DSCR = Revenue / Total Annual Cost >= target_dscr

    Args:
        self_consumption_y1_kwh:  Year-1 self-consumption from iPals (kWh)
        surplus_y1_kwh:           Year-1 surplus from iPals (kWh)
        selling_price:            Equipment selling price (yen)
        subsidy_amount:           Subsidy amount (yen)
        lease_company:            Finance company name (backward compat)
        lease_rate_pct:           Manual rate override (%) -- used for unknown companies
        lease_years:              Finance term (years)
        contract_years:           PPA contract duration (years)
        system_kw:                System capacity kW (for O&M calculation)
        fit_price:                FIT price for surplus (yen/kWh)
        include_surplus:          Whether surplus revenue is counted
        target_dscr:              Minimum acceptable DSCR (default 1.30)
        maintenance_yen_per_kw:   Maintenance fee per kW (default 1,200 yen/kW/yr)
        insurance_yen_fixed:      Fixed annual insurance fee (default 120,000 yen/yr)
        finance_company:          Explicit finance company name (overrides lease_company if set)

    Returns dict with:
        principal, effective_rate_pct, annual_lease_payment, annual_om_cost,
        total_annual_cost, min_ppa_price, cashflow_table, min_dscr, warnings,
        finance_type,
        (bank loan only): fire_insurance_annual, depreciation_tax_y1,
                          annual_interest_y1, annual_principal
    """
    warnings_list: list[str] = []

    # Use finance_company if provided, otherwise fall back to lease_company
    company = finance_company if finance_company else lease_company
    finance_type = get_finance_type(company)

    principal = max(selling_price - subsidy_amount, 0.0)
    if principal <= 0:
        warnings_list.append("販売価格・補助金額を確認してください（元本が0以下です）")

    if self_consumption_y1_kwh <= 0:
        warnings_list.append("iPalsデータがありません。自家消費量を入力してください")

    annual_payment, rate = calc_lease_payment(principal, company, lease_rate_pct, lease_years)

    # O&M cost
    om_cost = calc_annual_om_cost(system_kw, maintenance_yen_per_kw, insurance_yen_fixed)

    # Bank loan specific calculations
    loan_schedule: list[dict] | None = None
    fire_ins = 0
    dep_tax_schedule: list[int] | None = None

    if finance_type == FINANCE_TYPE_LOAN and principal > 0:
        loan_schedule = calc_bank_loan_annual_payments(principal, rate, lease_years)
        fire_ins = calc_fire_insurance_annual(selling_price)
        dep_tax_schedule = calc_depreciation_tax_schedule(selling_price, lease_years)

    min_price = 0.0
    cashflow_table: list[dict] = []
    min_dscr: float | None = None

    if principal > 0 and self_consumption_y1_kwh > 0:
        min_price = calc_min_ppa_price(
            self_consumption_y1_kwh=self_consumption_y1_kwh,
            surplus_y1_kwh=surplus_y1_kwh,
            annual_lease_payment=annual_payment,
            lease_years=lease_years,
            fit_price=fit_price,
            include_surplus=include_surplus,
            target_dscr=target_dscr,
            annual_om_cost=om_cost,
            finance_type=finance_type,
            loan_payment_schedule=loan_schedule,
            fire_insurance_annual=fire_ins,
            depreciation_tax_schedule=dep_tax_schedule,
        )

        cashflow_table = calc_cashflow_table(
            self_consumption_y1_kwh=self_consumption_y1_kwh,
            surplus_y1_kwh=surplus_y1_kwh,
            ppa_unit_price=min_price,
            fit_price=fit_price,
            annual_lease_payment=annual_payment,
            contract_years=contract_years,
            lease_years=lease_years,
            include_surplus=include_surplus,
            annual_om_cost=om_cost,
            finance_type=finance_type,
            loan_payment_schedule=loan_schedule,
            fire_insurance_annual=fire_ins,
            depreciation_tax_schedule=dep_tax_schedule,
            company=company,
        )

        dscr_values = [r["dscr"] for r in cashflow_table if r["dscr"] is not None]
        min_dscr = min(dscr_values) if dscr_values else None

    # Calculate IRR and NPV from cashflow table
    ppa_irr: float | None = None
    ppa_npv: float | None = None
    if cashflow_table and principal > 0:
        # IRR: initial outflow = -principal, then annual net cashflows
        cf_for_irr = [-principal] + [r["net_cashflow"] for r in cashflow_table]
        try:
            ppa_irr = round(irr(cf_for_irr) * 100, 2)  # as percentage
        except (ZeroDivisionError, ValueError, OverflowError):
            ppa_irr = None

        # NPV at discount rate = effective financing rate
        try:
            annual_cfs = [r["net_cashflow"] for r in cashflow_table]
            ppa_npv = round(npv(rate, annual_cfs) - principal)
        except (ZeroDivisionError, ValueError, OverflowError):
            ppa_npv = None

    # Re-lease annual amount (みずほリース)
    re_lease_annual = 0
    if company == "みずほリース" and annual_payment > 0:
        re_lease_annual = round(annual_payment * MIZUHO_RELEASE_RATIO)

    result = {
        "principal": round(principal),
        "effective_rate_pct": round(rate * 100, 2),
        "annual_lease_payment": round(annual_payment),
        "annual_om_cost": round(om_cost),
        "total_annual_cost": round(annual_payment + om_cost),
        "min_ppa_price": min_price,
        "cashflow_table": cashflow_table,
        "min_dscr": min_dscr,
        "warnings": warnings_list,
        "finance_type": finance_type,
        "irr_pct": ppa_irr,
        "npv_yen": ppa_npv,
        "re_lease_annual": re_lease_annual,
    }

    # Add bank loan specific fields
    if finance_type == FINANCE_TYPE_LOAN:
        result["fire_insurance_annual"] = fire_ins
        result["depreciation_tax_y1"] = dep_tax_schedule[0] if dep_tax_schedule else 0
        result["annual_interest_y1"] = loan_schedule[0]["interest"] if loan_schedule else 0
        result["annual_principal"] = loan_schedule[0]["principal"] if loan_schedule else 0

    return result
