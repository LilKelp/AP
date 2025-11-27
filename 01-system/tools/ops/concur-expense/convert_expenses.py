#!/usr/bin/env python3
"""Convert Concur Synchronized Accounting extracts into SAP-ready tables."""

from __future__ import annotations

import sys
from pathlib import Path
from datetime import datetime, date
from typing import Iterable

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[4]
INPUT_ROOT = BASE_DIR / "02-inputs" / "Concur"
OUTPUT_ROOT = BASE_DIR / "03-outputs" / "concur-expense"
VENDOR_ROOT = BASE_DIR / "02-inputs" / "Payment run raw"

REGIONS = [
    {
        "code": "AU",
        "data_dir": INPUT_ROOT / "AU",
        "vendor_file": VENDOR_ROOT / "AU Vendor list.xlsx",
        "employee_map": {
            "path": INPUT_ROOT / "AU NAME ID.xlsx",
            "sheet": None,
        },
    },
    {
        "code": "NZ",
        "data_dir": INPUT_ROOT / "NZ",
        "vendor_file": VENDOR_ROOT / "NZ Vendor list.xlsx",
        "employee_map": {
            "path": INPUT_ROOT / "NZ NAME ID.xlsx",
            "sheet": None,
        },
        "cost_center_transform": lambda value: (
            f"81{str(value)[2:]}" if isinstance(value, str) and value.startswith("80") else
            f"81{str(int(value))[2:]}" if isinstance(value, (int, float)) and str(int(value)).startswith("80") else value
        ),
    },
]

SKIP_KEYWORDS = {"EXAMPLE", "~$"}
MIXED_FLAG_COL = "Mixed_Tax_Flag"
TAXABLE_AMT_COL = "Taxable_Amount"
NONTAXABLE_AMT_COL = "Nontaxable_Amount"
MIXED_NOTE_COL = "Mixed_Note"
MIXED_COLS = [MIXED_FLAG_COL, TAXABLE_AMT_COL, NONTAXABLE_AMT_COL, MIXED_NOTE_COL]
MIXED_TAXABLE_DERIVED_COL = "Mixed_Taxable_Gross"
MIXED_NONTAXABLE_DERIVED_COL = "Mixed_Nontaxable_Gross"

GST_EXPECTED_RATE = 1 / 11
GST_RATE_TOLERANCE = 0.002
GST_ZERO_TOLERANCE = 0.01
MIXED_TOLERANCE = 0.05
EXPECTED_GST_RATE = {"AU": 0.10, "NZ": 0.15}

def normalize_account(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(int(round(value)))
    return str(value).strip()

def format_cost_center(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(int(round(value)))
    return str(value).strip()

def build_display_account(code: str) -> str:
    upper = code.upper()
    if upper.startswith("FB"):
        return f"{upper}-620120"
    return code

def map_sap_account(code: str) -> str:
    upper = code.upper()
    if upper.startswith("FB"):
        return "620120"
    return code


def normalize_name(text: str) -> str:
    if not isinstance(text, str):
        return ""
    return "".join(ch for ch in text.upper() if ch.isalnum())


def normalize_employee_id(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().lower()

def normalize_mixed_flag(value) -> bool:
    if isinstance(value, str):
        return value.strip().upper() in {"Y", "YES", "TRUE", "1"}
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return bool(value)
    if isinstance(value, bool):
        return value
    return False


def coerce_positive_number(value) -> float:
    try:
        num = float(value)
        if pd.isna(num):
            return 0.0
        return abs(num)
    except Exception:
        return 0.0


def ensure_mixed_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Guarantee mixed-tax helper columns exist for downstream splitting."""
    if df is None or df.empty:
        return df
    if MIXED_FLAG_COL not in df.columns:
        df[MIXED_FLAG_COL] = ""
    if TAXABLE_AMT_COL not in df.columns:
        df[TAXABLE_AMT_COL] = 0.0
    if NONTAXABLE_AMT_COL not in df.columns:
        df[NONTAXABLE_AMT_COL] = 0.0
    if MIXED_NOTE_COL not in df.columns:
        df[MIXED_NOTE_COL] = ""
    if MIXED_TAXABLE_DERIVED_COL not in df.columns:
        df[MIXED_TAXABLE_DERIVED_COL] = 0.0
    if MIXED_NONTAXABLE_DERIVED_COL not in df.columns:
        df[MIXED_NONTAXABLE_DERIVED_COL] = 0.0
    return df


def classify_line(row: pd.Series, region: str) -> pd.Series:
    """Classify lines into L0/L1/mixed using gross and GST; derive splits for mixed (AU/NZ)."""
    region_upper = region.upper()
    expected_rate = EXPECTED_GST_RATE.get(region_upper)
    if expected_rate is None:
        return row

    row = row.copy()
    gross = float(row.get("gross_amount", 0.0))
    gst = float(row.get("gst_amount", 0.0))
    gross_abs = abs(gross)
    gst_abs = abs(gst)
    note = str(row.get(MIXED_NOTE_COL, "") or "")

    row[MIXED_FLAG_COL] = str(row.get(MIXED_FLAG_COL, "")).upper() or "N"
    row[TAXABLE_AMT_COL] = coerce_positive_number(row.get(TAXABLE_AMT_COL, 0.0))
    row[NONTAXABLE_AMT_COL] = coerce_positive_number(row.get(NONTAXABLE_AMT_COL, 0.0))
    row[MIXED_TAXABLE_DERIVED_COL] = row[TAXABLE_AMT_COL]
    row[MIXED_NONTAXABLE_DERIVED_COL] = row[NONTAXABLE_AMT_COL]

    expected_ratio = expected_rate / (1 + expected_rate)  # GST / gross

    # Case A: effectively zero GST -> pure L0
    if gst_abs <= GST_ZERO_TOLERANCE:
        row[MIXED_FLAG_COL] = "N"
        row[MIXED_NOTE_COL] = note
        row["tax_code"] = "L0"
        return row

    # Case B: roughly expected GST on gross -> pure L1
    if gross_abs > 0.009:
        rate = gst_abs / gross_abs
        if abs(rate - expected_ratio) <= GST_RATE_TOLERANCE:
            row[MIXED_FLAG_COL] = "N"
            row[MIXED_NOTE_COL] = note
            row["tax_code"] = "L1"
            return row

        # Case C: GST present but materially below expected -> mixed candidate
        if rate < (expected_ratio - GST_RATE_TOLERANCE):
            taxable = round(gst_abs / expected_ratio, 2)  # gross portion that carries GST
            nontaxable = round(gross_abs - taxable, 2)
            if taxable <= gross_abs + MIXED_TOLERANCE and nontaxable >= -MIXED_TOLERANCE:
                row[MIXED_FLAG_COL] = "Y"
                row[TAXABLE_AMT_COL] = taxable
                row[NONTAXABLE_AMT_COL] = max(nontaxable, 0.0)
                row[MIXED_TAXABLE_DERIVED_COL] = taxable
                row[MIXED_NONTAXABLE_DERIVED_COL] = max(nontaxable, 0.0)
                row[MIXED_NOTE_COL] = note or "Auto-split mixed (GST below full rate)"
                row["tax_code"] = "L1"
                return row
            row[MIXED_FLAG_COL] = "CHECK"
            row[MIXED_NOTE_COL] = "Mixed candidate but derived split invalid; review GST/gross."
            return row

    row[MIXED_FLAG_COL] = "N"
    row[MIXED_NOTE_COL] = note
    return row

def load_vendor_lookup(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    df = pd.read_excel(path, usecols=[0, 1])
    df = df.dropna()
    lookup: dict[str, str] = {}
    for _, row in df.iterrows():
        supplier_id = row.iloc[0]
        name = row.iloc[1]
        try:
            supplier_id_str = str(int(round(float(supplier_id))))
        except Exception:
            supplier_id_str = str(supplier_id).strip()
        lookup[normalize_name(str(name))] = supplier_id_str
    return lookup


def load_employee_map(path: Path | None, sheet: str | None = None) -> dict[str, str]:
    if not path or not path.exists():
        return {}
    if sheet:
        try:
            df = pd.read_excel(path, sheet_name=sheet)
        except ValueError:
            df = pd.read_excel(path)
    else:
        df = pd.read_excel(path)
    # assume first two columns correspond to employee ID and supplier ID
    if df.shape[1] < 2:
        return {}
    df = df.iloc[:, :2].dropna()
    mapping: dict[str, str] = {}
    for _, row in df.iterrows():
        emp_id = normalize_employee_id(row.iloc[0])
        supplier_id = row.iloc[1]
        if not emp_id:
            continue
        try:
            supplier_str = str(int(round(float(supplier_id))))
        except Exception:
            supplier_str = str(supplier_id).strip()
        if supplier_str:
            mapping[emp_id] = supplier_str
    return mapping


def read_concur_file(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path)
    return pd.read_excel(path, sheet_name=0)


def map_employee_to_vendor(first: str, last: str, lookup: dict[str, str]) -> str:
    primary = normalize_name(f"{first} {last}")
    if primary in lookup:
        return lookup[primary]
    alternate = normalize_name(f"{last} {first}")
    return lookup.get(alternate, "")


def resolve_vendor_id(
    employee_id,
    first: str,
    last: str,
    employee_lookup: dict[str, str],
    vendor_lookup: dict[str, str],
) -> str:
    emp_key = normalize_employee_id(employee_id)
    if emp_key and emp_key in employee_lookup:
        return employee_lookup[emp_key]
    fallback = map_employee_to_vendor(first, last, vendor_lookup)
    return fallback


def determine_tax_code(gst_amount: float) -> str:
    if abs(gst_amount) > 0.009:
        return "L1"
    return "L0"


def normalize_key_value(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (datetime, date)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    return str(value).strip().upper()


def build_merge_key(row: pd.Series) -> tuple:
    emp = normalize_key_value(row.get("Employee ID"))
    report = normalize_key_value(row.get("Report ID"))
    transaction_date = normalize_key_value(row.get("Report Entry Transaction Date"))
    expense_type = normalize_key_value(row.get("Report Entry Expense Type Name"))
    vendor_name = normalize_key_value(row.get("Report Entry Vendor Name"))
    account = normalize_account(row.get("Journal Account Code")).upper()
    if transaction_date and expense_type and vendor_name:
        return ("KEY1", emp, report, transaction_date, expense_type, vendor_name, account)
    if transaction_date:
        return ("KEY2", emp, report, transaction_date, account)
    return ("KEY3", emp, report, account)


def format_merge_key(key: tuple) -> str:
    return " | ".join(str(part) for part in key)


def numeric_series(frame: pd.DataFrame, columns: list[str], default: float = 0.0) -> pd.Series:
    """Return the first available column converted to float, or a default-filled series."""
    for column in columns:
        if column in frame.columns:
            return pd.to_numeric(frame[column], errors="coerce").fillna(default)
    return pd.Series(default, index=frame.index, dtype=float)


def merge_gst_lines(expense_df: pd.DataFrame, gst_df: pd.DataFrame) -> tuple[pd.DataFrame, list[dict]]:
    """Merge standalone GST lines (DR) back into expense lines (CR) using deterministic keys."""
    expense_df = expense_df.copy()
    expense_df["merge_key"] = expense_df.apply(build_merge_key, axis=1) if not expense_df.empty else pd.Series(dtype=object)
    unmatched: list[dict] = []

    gst_totals = {}
    if gst_df is not None and not gst_df.empty:
        gst_df = gst_df.copy()
        gst_df["merge_key"] = gst_df.apply(build_merge_key, axis=1)
        gst_df["gst_value"] = numeric_series(
            gst_df,
            ["Report Entry Total Tax Posted Amount", "Report Entry Tax Posted Amount"],
        )
        gst_totals = gst_df.groupby("merge_key")["gst_value"].sum()
    else:
        gst_df = pd.DataFrame()

    for key, gst_total in gst_totals.items():
        mask = expense_df["merge_key"] == key
        if mask.any():
            share_base = expense_df.loc[mask, "gross_amount"].abs()
            total_share = share_base.sum()
            if total_share <= 0:
                allocation = pd.Series([1.0 / len(share_base)] * len(share_base), index=share_base.index)
            else:
                allocation = share_base / total_share
            expense_df.loc[mask, "gst_amount"] = gst_total * allocation
        else:
            sample = gst_df.loc[gst_df["merge_key"] == key].iloc[0]
            unmatched.append({
                "Employee ID": sample.get("Employee ID", ""),
                "Report ID": sample.get("Report ID", ""),
                "Report Submit Date": sample.get("Report Submit Date", ""),
                "Key": format_merge_key(key),
                "GST Found": round(float(gst_total), 2),
                "Expense Matched": 0,
                "Action": "Unmatched GST line",
            })
            print(f"[WARN] GST line unmatched for key {format_merge_key(key)}: gst={gst_total:.2f}")

    expense_df = expense_df.drop(columns=["merge_key"])
    return expense_df, unmatched

def split_mixed_lines(df: pd.DataFrame) -> pd.DataFrame:
    """Split flagged mixed-tax lines into separate L1/L0 rows using provided amounts."""
    if df.empty:
        return df
    rows = []
    tolerance = 0.05
    for _, row in df.iterrows():
        flag = str(row.get(MIXED_FLAG_COL, "")).upper()
        if flag != "Y":
            row_copy = row.copy()
            row_copy[MIXED_FLAG_COL] = "N" if flag == "" else flag
            row_copy["Mixed_Segment"] = "" if flag != "CHECK" else "Mixed candidate - review"
            rows.append(row_copy)
            continue

        taxable = coerce_positive_number(row.get(TAXABLE_AMT_COL, 0.0))
        non_taxable = coerce_positive_number(row.get(NONTAXABLE_AMT_COL, 0.0))
        mixed_note = row.get(MIXED_NOTE_COL, "")

        if taxable <= 0 and non_taxable <= 0:
            row_copy = row.copy()
            row_copy[MIXED_FLAG_COL] = "CHECK"
            row_copy["Mixed_Segment"] = "Mixed candidate - missing amounts"
            rows.append(row_copy)
            continue

        gross_abs = abs(float(row.get("gross_amount", 0.0)))
        total_specified = taxable + non_taxable
        if abs(total_specified - gross_abs) > tolerance:
            row_copy = row.copy()
            row_copy[MIXED_FLAG_COL] = "CHECK"
            row_copy["Mixed_Segment"] = "Mixed candidate - totals mismatch"
            rows.append(row_copy)
            continue

        sign = -1.0 if float(row.get("gross_amount", 0.0)) < 0 else 1.0

        # L1 portion
        l1_row = row.copy()
        l1_row["gross_amount"] = round(sign * taxable, 2)
        l1_row["gst_amount"] = round(sign * taxable / 11, 2)
        l1_row["net_amount"] = round(l1_row["gross_amount"] - l1_row["gst_amount"], 2)
        l1_row["tax_code"] = "L1"
        l1_row[MIXED_FLAG_COL] = "Y"
        l1_row["Mixed_Segment"] = "L1 portion"
        l1_row[MIXED_NOTE_COL] = mixed_note
        l1_row[MIXED_TAXABLE_DERIVED_COL] = taxable
        l1_row[MIXED_NONTAXABLE_DERIVED_COL] = non_taxable
        rows.append(l1_row)

        # L0 portion
        l0_row = row.copy()
        l0_row["gross_amount"] = round(sign * non_taxable, 2)
        l0_row["gst_amount"] = 0.0
        l0_row["net_amount"] = round(l0_row["gross_amount"], 2)
        l0_row["tax_code"] = "L0"
        l0_row[MIXED_FLAG_COL] = "Y"
        l0_row["Mixed_Segment"] = "L0 portion"
        l0_row[MIXED_NOTE_COL] = mixed_note
        l0_row[MIXED_TAXABLE_DERIVED_COL] = taxable
        l0_row[MIXED_NONTAXABLE_DERIVED_COL] = non_taxable
        rows.append(l0_row)
    return pd.DataFrame(rows)


def validate_gst_rates(df: pd.DataFrame, region: str) -> None:
    """Ensure GST rates align with AU (10% or 0) and NZ (15% or 0); skip derived mixed rows."""
    if df.empty:
        return
    region_upper = region.upper()
    expected_rate = {"AU": 0.10, "NZ": 0.15}.get(region_upper)
    if expected_rate is None:
        return
    tolerance = 0.005
    df = df.copy()
    df[MIXED_FLAG_COL] = df.get(MIXED_FLAG_COL, "").fillna("")
    non_mixed_mask = ~df[MIXED_FLAG_COL].isin(["Y", "CHECK"])
    df = df.loc[non_mixed_mask]
    if df.empty:
        return
    gst = df["gst_amount"].astype(float).abs()
    net = df["net_amount"].astype(float).abs()
    rate = pd.Series(0.0, index=df.index)
    nonzero_net = net > 0.009
    rate.loc[nonzero_net] = (gst.loc[nonzero_net] / net.loc[nonzero_net]).astype(float)
    valid_zero = gst <= 0.009
    valid_expected = (rate >= expected_rate - tolerance) & (rate <= expected_rate + tolerance)
    invalid_mask = ~(valid_zero | valid_expected)
    if invalid_mask.any():
        sample_cols = [col for col in ["Employee ID", "Report ID", "gross_amount", "gst_amount", "net_amount"] if col in df.columns]
        sample = df.loc[invalid_mask, sample_cols].head(5)
        print(
            f"[WARN] {region_upper}: GST rate check flagged {invalid_mask.sum()} of {len(df)} rows "
            f"(expected {expected_rate:.1%} +/- {tolerance*100:.1f}%, or zero)."
            f"\nSample:\n{sample.to_string(index=False)}"
        )


def iter_region_files(region_dir: Path) -> Iterable[Path]:
    for path in sorted(region_dir.iterdir()):
        if not path.is_file():
            continue
        suffix = path.suffix.lower()
        if suffix not in {".xlsx", ".xls", ".xlsm", ".csv"}:
            continue
        upper_name = path.name.upper()
        if any(key in upper_name for key in SKIP_KEYWORDS):
            continue
        yield path

def prepare_company_rows(
    df: pd.DataFrame,
    vendor_lookup: dict[str, str],
    employee_lookup: dict[str, str],
    region: str,
    cost_center_transform=None,
) -> tuple[pd.DataFrame, list[dict]]:
    df = ensure_mixed_columns(df.copy())
    payer = df.get("Journal Payer Payment Type Name", pd.Series(dtype=str)).fillna("").astype(str)
    payment_code = df.get("Report Entry Payment Code Name", pd.Series(dtype=str)).fillna("").astype(str)
    mask_company = payer.str.upper().eq("COMPANY")
    mask_cash = payment_code.str.upper().eq("CASH")
    comp = df.loc[mask_company & mask_cash].copy()
    comp = comp[comp["Journal Account Code"].notna()].copy()
    comp["Report Submit Date"] = pd.to_datetime(comp["Report Submit Date"], errors="coerce", dayfirst=True).dt.date
    comp["Report Entry Transaction Date"] = pd.to_datetime(
        comp.get("Report Entry Transaction Date"), errors="coerce", dayfirst=True
    ).dt.date
    comp["Department"] = comp["Department"].apply(format_cost_center)
    if cost_center_transform:
        comp["Department"] = comp["Department"].apply(cost_center_transform)
    comp["gross_amount"] = numeric_series(comp, ["Journal Amount"])
    comp["gst_amount"] = numeric_series(
        comp,
        ["Report Entry Total Tax Posted Amount", "Report Entry Tax Posted Amount"],
    )
    comp["Journal Debit Or Credit"] = comp.get("Journal Debit Or Credit", pd.Series(dtype=str)).fillna("").astype(str).str.upper().str.strip()
    comp["Report Entry Tax Code"] = comp.get("Report Entry Tax Code", pd.Series(dtype=str)).fillna("").astype(str).str.upper().str.strip()
    comp["normalized_account"] = comp["Journal Account Code"].apply(normalize_account)
    comp["display_account"] = comp["normalized_account"].apply(build_display_account)
    comp["sap_account"] = comp["normalized_account"].apply(map_sap_account)
    comp["SAP Vendor ID"] = [
        resolve_vendor_id(emp_id, first, last, employee_lookup, vendor_lookup)
        for emp_id, first, last in zip(
            comp.get("Employee ID", ""),
            comp.get("Employee First Name", ""),
            comp.get("Employee Last Name", ""),
        )
    ]
    gst_mask = comp["Report Entry Tax Code"].eq("GST") & comp["Journal Debit Or Credit"].eq("DR")
    expense_mask = comp["Journal Debit Or Credit"].eq("CR") & ~comp["Report Entry Tax Code"].eq("GST")
    gst_lines = comp.loc[gst_mask].copy()
    expense_lines = comp.loc[expense_mask].copy()
    expense_lines = ensure_mixed_columns(expense_lines)
    expense_lines, unmatched = merge_gst_lines(expense_lines, gst_lines)
    expense_lines["net_amount"] = expense_lines["gross_amount"] - expense_lines["gst_amount"]
    expense_lines["tax_code"] = expense_lines["gst_amount"].apply(determine_tax_code)
    expense_lines = expense_lines.apply(lambda row: classify_line(row, region), axis=1)
    expense_lines = split_mixed_lines(expense_lines)
    return expense_lines, unmatched

def aggregate_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=[
            "Employee ID",
            "Report ID",
            "Report Submit Date",
            "Department",
            "SAP Vendor ID",
            "display_account",
            "sap_account",
            "gross_amount",
            "net_amount",
            "gst_amount",
            "tax_code",
            "sap_amount",
            MIXED_FLAG_COL,
            "Mixed_Segment",
            MIXED_NOTE_COL,
            MIXED_TAXABLE_DERIVED_COL,
            MIXED_NONTAXABLE_DERIVED_COL,
        ])
    if MIXED_FLAG_COL not in df.columns:
        df[MIXED_FLAG_COL] = "N"
    if "Mixed_Segment" not in df.columns:
        df["Mixed_Segment"] = ""
    if MIXED_NOTE_COL not in df.columns:
        df[MIXED_NOTE_COL] = ""
    if MIXED_TAXABLE_DERIVED_COL not in df.columns:
        df[MIXED_TAXABLE_DERIVED_COL] = 0.0
    if MIXED_NONTAXABLE_DERIVED_COL not in df.columns:
        df[MIXED_NONTAXABLE_DERIVED_COL] = 0.0
    group_cols = [
        "Employee ID",
        "Report ID",
        "Report Submit Date",
        "Department",
        "SAP Vendor ID",
        "display_account",
        "sap_account",
        "tax_code",
        MIXED_FLAG_COL,
    ]
    agg = (
        df.groupby(group_cols, dropna=False)
        .agg({
            "gross_amount": "sum",
            "gst_amount": "sum",
            "Mixed_Segment": "first",
            MIXED_NOTE_COL: "first",
            MIXED_TAXABLE_DERIVED_COL: "first",
            MIXED_NONTAXABLE_DERIVED_COL: "first",
        })
        .reset_index()
    )
    agg["gross_amount"] = agg["gross_amount"].round(2)
    agg["gst_amount"] = agg["gst_amount"].round(2)
    agg["net_amount"] = (agg["gross_amount"] - agg["gst_amount"]).round(2)
    agg["sap_amount"] = agg["gross_amount"].abs().round(2)
    agg = agg.sort_values(
        ["SAP Vendor ID", "Report ID", "Employee ID", "Report Submit Date", "Department", "display_account", "tax_code"]
    ).reset_index(drop=True)
    return agg


def apply_region_tax_display(agg: pd.DataFrame, region: str) -> pd.DataFrame:
    if agg.empty:
        agg["tax_code_display"] = pd.Series(dtype=str)
        return agg
    if region.upper() == "NZ":
        mapping = {"L1": "Q2", "L0": "Q0"}
        agg["tax_code_display"] = agg["tax_code"].map(mapping).fillna(agg["tax_code"])
    else:
        agg["tax_code_display"] = agg["tax_code"]
    return agg

def build_gst_check(agg: pd.DataFrame, unmatched: list[dict] | None = None) -> pd.DataFrame:
    unmatched = unmatched or []
    base_columns = [
        "Employee ID",
        "SAP Vendor ID",
        "Report ID",
        "Report Submit Date",
        "Gross Amount",
        "Net Amount",
        "GST Amount",
        "Calculated GST (Gross-Net)",
        "Difference",
        "Status",
        MIXED_FLAG_COL,
        MIXED_NOTE_COL,
        MIXED_TAXABLE_DERIVED_COL,
        MIXED_NONTAXABLE_DERIVED_COL,
        "Key",
        "GST Found",
        "Expense Matched",
        "Action",
    ]
    frames: list[pd.DataFrame] = []

    if not agg.empty:
        group_cols = ["Employee ID", "SAP Vendor ID", "Report ID", "Report Submit Date"]
        mixed_lookup = (
            agg.groupby(group_cols, dropna=False)[MIXED_FLAG_COL]
            .apply(lambda s: "Y" if (s == "Y").any() else "N")
        )
        note_lookup = (
            agg.groupby(group_cols, dropna=False)[MIXED_NOTE_COL]
            .apply(lambda s: next((val for val in s if isinstance(val, str) and val.strip()), ""))
        )
        taxable_lookup = (
            agg.groupby(group_cols, dropna=False)
            .apply(lambda g: g.loc[
                (g[MIXED_FLAG_COL] == "Y") & (g["tax_code"].str.upper() == "L1"), "gross_amount"
            ].abs().sum())
        )
        nontaxable_lookup = (
            agg.groupby(group_cols, dropna=False)
            .apply(lambda g: g.loc[
                (g[MIXED_FLAG_COL] == "Y") & (g["tax_code"].str.upper() == "L0"), "gross_amount"
            ].abs().sum())
        )
        recon = (
            agg.groupby(group_cols, dropna=False)
            .agg({
                "gross_amount": "sum",
                "net_amount": "sum",
                "gst_amount": "sum",
            })
            .reset_index()
        )
        recon["Gross Amount"] = recon["gross_amount"].abs().round(2)
        recon["Net Amount"] = recon["net_amount"].abs().round(2)
        recon["GST Amount"] = recon["gst_amount"].abs().round(2)
        recon["Calculated GST (Gross-Net)"] = (recon["Gross Amount"] - recon["Net Amount"]).round(2)
        recon["Difference"] = (recon["GST Amount"] - recon["Calculated GST (Gross-Net)"]).round(2)
        recon["Status"] = recon["Difference"].abs().lt(0.01).map({True: "OK", False: "CHECK"})
        recon["Key"] = ""
        recon["GST Found"] = ""
        recon["Expense Matched"] = ""
        recon["Action"] = ""
        recon[MIXED_FLAG_COL] = recon.apply(
            lambda row: mixed_lookup.get(
                (row["Employee ID"], row["SAP Vendor ID"], row["Report ID"], row["Report Submit Date"]),
                "N",
            ),
            axis=1,
        )
        recon[MIXED_NOTE_COL] = recon.apply(
            lambda row: note_lookup.get(
                (row["Employee ID"], row["SAP Vendor ID"], row["Report ID"], row["Report Submit Date"]),
                "",
            ),
            axis=1,
        )
        recon[MIXED_TAXABLE_DERIVED_COL] = recon.apply(
            lambda row: taxable_lookup.get(
                (row["Employee ID"], row["SAP Vendor ID"], row["Report ID"], row["Report Submit Date"]),
                0.0,
            ),
            axis=1,
        )
        recon[MIXED_NONTAXABLE_DERIVED_COL] = recon.apply(
            lambda row: nontaxable_lookup.get(
                (row["Employee ID"], row["SAP Vendor ID"], row["Report ID"], row["Report Submit Date"]),
                0.0,
            ),
            axis=1,
        )
        frames.append(recon[base_columns])

    if unmatched:
        diag_rows = []
        for item in unmatched:
            diag_rows.append({
                "Employee ID": item.get("Employee ID", ""),
                "SAP Vendor ID": "",
                "Report ID": item.get("Report ID", ""),
                "Report Submit Date": item.get("Report Submit Date", ""),
                "Gross Amount": "",
                "Net Amount": "",
                "GST Amount": item.get("GST Found", ""),
                "Calculated GST (Gross-Net)": "",
                "Difference": "",
                "Status": "CHECK",
                MIXED_FLAG_COL: "",
                MIXED_NOTE_COL: "",
                "Key": item.get("Key", ""),
                "GST Found": item.get("GST Found", ""),
                "Expense Matched": item.get("Expense Matched", ""),
                "Action": item.get("Action", "Unmatched GST line"),
            })
        frames.append(pd.DataFrame(diag_rows))

    if not frames:
        return pd.DataFrame(columns=base_columns)

    recon_full = pd.concat(frames, ignore_index=True)
    return recon_full[base_columns]

def build_sap_view(agg: pd.DataFrame) -> pd.DataFrame:
    rows = []
    group_fields = ["Employee ID", "SAP Vendor ID", "Report ID", "Report Submit Date"]
    for _, group in agg.groupby(group_fields, sort=False):
        first = True
        for _, row in group.iterrows():
            text_value = ""
            if row.get(MIXED_FLAG_COL, "") == "Y":
                segment = row.get("Mixed_Segment", "").strip()
                note = row.get(MIXED_NOTE_COL, "").strip()
                parts = [part for part in [segment or "Mixed item", note] if part]
                text_value = " | ".join(parts)
            prefix = {
                "Concur Employee ID": row["Employee ID"] if first else "",
                "SAP Supplier ID": row["SAP Vendor ID"] if first else "",
                "Report ID": row["Report ID"] if first else "",
                "Report Submit Date": row["Report Submit Date"] if first else "",
            }
            rows.append({
                **prefix,
                "Account (I)": row["sap_account"],
                "Assignment (J)": "",
                "Amount (K)": row["sap_amount"],
                "Tax Code": row["tax_code_display"],
                "Text (M)": text_value,
                "Cost Center (N)": row["Department"],
            })
            first = False
        total_amount = round(group["sap_amount"].sum(), 2)
        rows.append({
            "Concur Employee ID": "",
            "SAP Supplier ID": "",
            "Report ID": "",
            "Report Submit Date": "",
            "Account (I)": "REPORT TOTAL",
            "Assignment (J)": "",
            "Amount (K)": total_amount,
            "Tax Code": "",
            "Text (M)": "Report total (validation only)",
            "Cost Center (N)": "",
        })
    return pd.DataFrame(rows)



def process_file(
    region: str,
    path: Path,
    vendor_lookup: dict[str, str],
    employee_lookup: dict[str, str],
    cost_center_transform=None,
) -> tuple[Path, pd.DataFrame]:
    raw_df = read_concur_file(path)
    raw_df = ensure_mixed_columns(raw_df)
    comp, unmatched_gst = prepare_company_rows(raw_df.copy(), vendor_lookup, employee_lookup, region, cost_center_transform)
    validate_gst_rates(comp, region)
    agg = aggregate_rows(comp)
    agg = apply_region_tax_display(agg, region)
    sap_view = build_sap_view(agg)
    gst_check = build_gst_check(agg, unmatched_gst)
    output_dir = OUTPUT_ROOT / region
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"SAP_{region}_{path.stem}.xlsx"
    if output_path.exists():
        try:
            output_path.unlink()
        except PermissionError:
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            output_path = output_dir / f"SAP_{region}_{path.stem}_{timestamp}.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        agg.rename(columns={
            "display_account": "Journal Account Code",
            "sap_account": "SAP GL",
            "gross_amount": "Journal Amount (Gross)",
            "net_amount": "Net Amount",
            "gst_amount": "GST Amount",
            "SAP Vendor ID": "SAP Supplier ID",
            "tax_code_display": "Tax Code",
        }).to_excel(writer, sheet_name="Summary", index=False)
        sap_view.to_excel(writer, sheet_name="SAP_Paste", index=False)
        if not gst_check.empty:
            gst_check.to_excel(writer, sheet_name="GST_Check", index=False)
        raw_df.to_excel(writer, sheet_name="Raw_Input", index=False)
    return output_path, agg

def process_region(
    region: str,
    region_dir: Path,
    vendor_lookup: dict[str, str],
    employee_lookup: dict[str, str],
    cost_center_transform=None,
) -> list[Path]:
    outputs: list[Path] = []
    for file_path in iter_region_files(region_dir):
        print(f"[INFO] {region}: transforming {file_path.name}")
        output_path, _ = process_file(region, file_path, vendor_lookup, employee_lookup, cost_center_transform)
        outputs.append(output_path)
    return outputs

def main() -> int:
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    if not INPUT_ROOT.exists():
        print(f"Input folder not found: {INPUT_ROOT}")
        return 1
    regions_to_process = []
    if len(sys.argv) > 1:
        requested = {arg.upper() for arg in sys.argv[1:]}
        regions_to_process = [conf for conf in REGIONS if conf["code"].upper() in requested]
    else:
        regions_to_process = REGIONS

    generated: list[Path] = []
    for region_conf in regions_to_process:
        region_dir = region_conf["data_dir"]
        if not region_dir.exists():
            continue
        vendor_lookup = load_vendor_lookup(region_conf["vendor_file"])
        emp_map_conf = region_conf.get("employee_map", {})
        employee_lookup = load_employee_map(emp_map_conf.get("path"), emp_map_conf.get("sheet"))
        cost_center_transform = region_conf.get("cost_center_transform")
        outputs = process_region(
            region_conf["code"],
            region_dir,
            vendor_lookup,
            employee_lookup,
            cost_center_transform,
        )
        generated.extend(outputs)

    if not generated:
        print("No Concur extracts were processed.")
        return 1

    print("\nCreated the following SAP-formatted files:")
    for path in generated:
        print(f"  - {path.relative_to(BASE_DIR)}")
    return 0

if __name__ == "__main__":
    sys.exit(main())
