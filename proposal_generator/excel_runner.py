"""
excel_runner.py
Writes customer variables into the Excel calculation engine (hidden),
triggers the VBA macro, and reads back all output values needed for slides.

Workflow:
    1. Open Excel silently (xlwings, visible=False)
    2. Write CustomerInput values to PPAリース sheet
    3. Write iPals data to PalsDATA sheet
    4. Run the calculation macro
    5. Read output cells and return as dict
    6. Close Excel (save=False to avoid overwriting the template)
"""

from __future__ import annotations

import csv
import io
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Optional

import openpyxl

# ---------------------------------------------------------------------------
# Customer input model
# ---------------------------------------------------------------------------

@dataclass
class CustomerInput:
    """All variables that map to the yellow input cells in PPAリース sheet."""
    company_name: str = ""          # C1 企業名
    office_name: str = ""           # F1 事業所名
    address: str = ""               # I1 設置先住所
    snow_depth: float = 0.0         # L1 垂直積雪量(m)
    proposal_date: str = ""         # O1 提案日 (YYYY-MM-DD or display string)
    company_size: str = ""          # P1 企業規模
    site_survey: str = ""           # R1 現地調査
    tax_display: str = "税抜"       # U1 提案書税表記

    # System info (1種類目)
    panel_watt: float = 0.0         # D5 パネル出力(W)
    panel_count: int = 0            # D6 パネル枚数(枚)
    system_capacity_kw: float = 0.0 # D7 システム容量(kW)
    pcs_output_kw: float = 0.0      # D8 PCS出力(kW)
    battery_kwh: float = 0.0        # D10 蓄電池容量(kWh)

    # Contract info
    contract_years: int = 20        # Q5 契約期間(年)
    ppa_unit_price: float = 0.0     # L8 PPA単価(円/kWh)

    # Subsidy
    subsidy_name: str = ""          # 補助金名（活用補助金）
    subsidy_amount: float = 0.0     # 補助金額(円)

    # Surplus electricity
    surplus_price: float = 0.0      # 余剰売電単価(円/kWh)

    # Demand reduction
    demand_reduction_kw: float = 0.0  # 削減デマンド(kW)

    # Lease info
    lease_company: str = "シーエナジー"
    lease_years: int = 20
    lease_rate: float = 6.0

    # EPC (if applicable)
    epc_list_price: float = 0.0
    epc_cost: float = 0.0


# ---------------------------------------------------------------------------
# Cell address mapping (PPAリース sheet)
# Label cells omitted – these are the VALUE cells
# ---------------------------------------------------------------------------
INPUT_CELL_MAP = {
    "company_name":        ("PPAリース", "C1"),
    "office_name":         ("PPAリース", "F1"),
    "address":             ("PPAリース", "I1"),
    "snow_depth":          ("PPAリース", "L1"),
    "proposal_date":       ("PPAリース", "O1"),
    "company_size":        ("PPAリース", "P1"),
    "site_survey":         ("PPAリース", "R1"),
    "tax_display":         ("PPAリース", "U1"),
    "panel_watt":          ("PPAリース", "D5"),
    "panel_count":         ("PPAリース", "D6"),
    "system_capacity_kw":  ("PPAリース", "D7"),
    "pcs_output_kw":       ("PPAリース", "D8"),
    "battery_kwh":         ("PPAリース", "D10"),
    "contract_years":      ("PPAリース", "Q5"),
    "ppa_unit_price":      ("PPAリース", "L8"),
    "lease_company":       ("PPAリース", "U5"),
    "lease_years":         ("PPAリース", "U6"),
    "lease_rate":          ("PPAリース", "U8"),
    "surplus_price":       ("PPAリース", "O13"),
    "demand_reduction_kw": ("PPAリース", "S20"),
    "epc_list_price":      ("PPAリース", "X17"),
    "epc_cost":            ("PPAリース", "X18"),
}

# Output cells to collect after calculation (referenced by slide sheets)
# Format: { "variable_name": ("SheetName", "CellAddress") }
OUTPUT_CELL_MAP = {
    # Basic info
    "company_name":          ("PPAリース", "C1"),
    "office_name":           ("PPAリース", "F1"),
    "proposal_date":         ("PPAリース", "O1"),
    "system_capacity_kw":    ("PPAリース", "D7"),
    "panel_count":           ("PPAリース", "D6"),
    "contract_years":        ("PPAリース", "Q5"),
    "ppa_unit_price":        ("PPAリース", "L8"),

    # Financial outputs (CE sheet)
    "annual_electricity_kwh": ("CE", "D12"),
    "annual_cost_saving":     ("CE", "D14"),
    "co2_reduction_t":        ("CE", "D16"),
    "investment_recovery_yr": ("CE", "D18"),
    "irr":                    ("CE", "D20"),
    "npv":                    ("CE", "D22"),

    # PPA outputs
    "ppa_20yr_revenue":      ("PPAリース", "J25"),
    "ppa_normal_dscr":       ("PPAリース", "J27"),
    "subsidy_amount":        ("PPAリース", "P14"),
    "surplus_price":         ("PPAリース", "O13"),

    # EPC outputs
    "epc_list_price":        ("PPAリース", "X17"),
    "epc_cost":              ("PPAリース", "X18"),

    # CO2
    "co2_annual_t":          ("CO2計算結果", "B3"),
    "co2_20yr_t":            ("CO2計算結果", "B4"),
}


def run_excel_calculation(
    excel_path: Path,
    customer: CustomerInput,
    ipals_csv: Optional[str] = None,
    ipals_sheet: str = "PalsDATA（自家消費量）",
    macro_name: str = "データ貼付",
) -> dict:
    """
    Open Excel hidden, write inputs, run macro, read outputs.

    Args:
        excel_path: Path to the .xlsm file
        customer: CustomerInput instance
        ipals_csv: Raw CSV text from iPals simulation export (optional)
        ipals_sheet: Which PalsDATA sheet to write to
        macro_name: VBA macro name to run after writing iPals data

    Returns:
        dict of output variable name -> value
    """
    try:
        import xlwings as xw
    except ImportError:
        raise RuntimeError("xlwings is required. Run: python -m pip install xlwings")

    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(excel_path))

        # --- Write customer input cells ---
        for field_name, (sheet_name, cell_addr) in INPUT_CELL_MAP.items():
            value = getattr(customer, field_name, None)
            if value is not None and value != "" and value != 0.0 and value != 0:
                wb.sheets[sheet_name].range(cell_addr).value = value

        # --- Write iPals data ---
        if ipals_csv:
            _write_ipals_data(wb, ipals_csv, ipals_sheet)
            # Run the macro that copies iPals data to ①使用電力量 etc.
            try:
                app.macro(macro_name)()
            except Exception as e:
                print(f"Warning: macro '{macro_name}' failed: {e}")

        # --- Force recalculation ---
        wb.app.calculate()

        # --- Read output cells ---
        output = {}
        for var_name, (sheet_name, cell_addr) in OUTPUT_CELL_MAP.items():
            try:
                output[var_name] = wb.sheets[sheet_name].range(cell_addr).value
            except Exception:
                output[var_name] = None

        return output

    finally:
        wb.close()
        app.quit()


def _write_ipals_data(wb, csv_text: str, sheet_name: str) -> None:
    """Parse iPals CSV and write rows to PalsDATA sheet starting at B2."""
    reader = csv.reader(io.StringIO(csv_text))
    rows = list(reader)
    if not rows:
        return
    ws = wb.sheets[sheet_name]
    # Write from row 2 (row 1 = header kept as-is)
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            try:
                val = float(val)
            except (ValueError, TypeError):
                pass
            ws.range((r_idx + 2, c_idx + 2)).value = val


# ---------------------------------------------------------------------------
# Fallback: read cached values from saved Excel (no macros, read-only)
# Used when xlwings is unavailable or Excel is not installed.
# ---------------------------------------------------------------------------

def read_cached_excel(excel_path: Path) -> dict:
    """
    Read the last-saved calculated values from Excel using openpyxl.
    Does NOT recalculate – assumes the file was saved after running macros.

    Returns the same output dict structure as run_excel_calculation().
    """
    wb = openpyxl.load_workbook(str(excel_path), data_only=True, read_only=True)
    output = {}
    for var_name, (sheet_name, cell_addr) in OUTPUT_CELL_MAP.items():
        try:
            ws = wb[sheet_name]
            output[var_name] = ws[cell_addr].value
        except Exception:
            output[var_name] = None
    wb.close()
    return output
