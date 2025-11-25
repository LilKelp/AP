#!/usr/bin/env python3
"""Convert Concur Synchronized Accounting extracts into SAP-ready tables."""

from __future__ import annotations

import sys
from pathlib import Path
from datetime import datetime
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
GST_HINT_TAXABLE = {"Q2", "Q15"}
GST_HINT_NONTAX = {"Q0"}

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


def determine_tax_code(hint: str, gst_amount: float) -> str:
    if hint in GST_HINT_TAXABLE or hint.startswith("Q2"):
        return "L1"
    if hint in GST_HINT_NONTAX or hint.startswith("Q0"):
        return "L0"
    if abs(gst_amount) > 0.009:
        return "L1"
    return "L0"


def numeric_series(frame: pd.DataFrame, columns: list[str], default: float = 0.0) -> pd.Series:
    """Return the first available column converted to float, or a default-filled series."""
    for column in columns:
        if column in frame.columns:
            return pd.to_numeric(frame[column], errors="coerce").fillna(default)
    return pd.Series(default, index=frame.index, dtype=float)


def validate_gst_rates(df: pd.DataFrame, region: str) -> None:
    """Ensure GST rates align with AU (10% or 0) and NZ (15% or 0)."""
    if df.empty:
        return
    region_upper = region.upper()
    expected_rate = {"AU": 0.10, "NZ": 0.15}.get(region_upper)
    if expected_rate is None:
        return
    tolerance = 0.005
    gst = df["gst_amount"].abs()
    net = df["net_amount"].abs()
    rate = pd.Series(0.0, index=df.index)
    nonzero_net = net > 0.009
    rate.loc[nonzero_net] = (gst.loc[nonzero_net] / net.loc[nonzero_net]).astype(float)
    valid_zero = gst <= 0.009
    valid_expected = (rate >= expected_rate - tolerance) & (rate <= expected_rate + tolerance)
    invalid_mask = ~(valid_zero | valid_expected)
    if invalid_mask.any():
        sample = (
            df.loc[invalid_mask, ["Employee ID", "Report ID", "Journal Amount", "gst_amount", "net_amount"]]
            .head(5)
        )
        raise ValueError(
            f"{region_upper}: GST rate must be 0 or {expected_rate:.0%}; found {invalid_mask.sum()} rows outside tolerance."
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
    cost_center_transform=None,
) -> pd.DataFrame:
    payer = df.get("Journal Payer Payment Type Name", pd.Series(dtype=str)).fillna("").astype(str)
    payment_code = df.get("Report Entry Payment Code Name", pd.Series(dtype=str)).fillna("").astype(str)
    mask_company = payer.str.upper().eq("COMPANY")
    mask_cash = payment_code.str.upper().eq("CASH")
    comp = df.loc[mask_company & mask_cash].copy()
    comp = comp[comp["Journal Account Code"].notna()].copy()
    comp["Report Submit Date"] = pd.to_datetime(comp["Report Submit Date"], errors="coerce").dt.date
    comp["Department"] = comp["Department"].apply(format_cost_center)
    if cost_center_transform:
        comp["Department"] = comp["Department"].apply(cost_center_transform)
    comp["gross_amount"] = numeric_series(comp, ["Journal Amount"])
    comp["gst_amount"] = numeric_series(
        comp,
        ["Report Entry Total Tax Posted Amount"],
    )
    comp["net_amount"] = numeric_series(comp, ["Net Tax Amount"])
    if "Net Tax Amount" not in comp.columns:
        comp["net_amount"] = comp["gross_amount"] - comp["gst_amount"]
    hints = comp.get("Report Entry Tax Code", pd.Series(dtype=str)).fillna("").astype(str).str.upper().str.strip()
    comp["tax_code"] = [
        determine_tax_code(hint, gst_amount)
        for hint, gst_amount in zip(hints, comp["gst_amount"])
    ]
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
    return comp

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
            "sap_amount"
        ])
    group_cols = [
        "Employee ID",
        "Report ID",
        "Report Submit Date",
        "Department",
        "SAP Vendor ID",
        "display_account",
        "sap_account",
        "tax_code",
    ]
    agg = (
        df.groupby(group_cols, dropna=False)
        .agg({
            "gross_amount": "sum",
            "net_amount": "sum",
            "gst_amount": "sum",
        })
        .reset_index()
    )
    agg["sap_amount"] = agg["gross_amount"].abs().round(2)
    agg["gross_amount"] = agg["gross_amount"].round(2)
    agg["net_amount"] = agg["net_amount"].round(2)
    agg["gst_amount"] = agg["gst_amount"].round(2)
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

def build_gst_check(agg: pd.DataFrame) -> pd.DataFrame:
    if agg.empty:
        return pd.DataFrame()
    group_cols = ["Employee ID", "SAP Vendor ID", "Report ID", "Report Submit Date"]
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
    return recon[
        [
            "Employee ID",
            "SAP Vendor ID",
            "Report ID",
            "Report Submit Date",
            "Gross Amount",
            "Net Amount",
            "GST Amount",
            "Calculated GST (Gross-Net)",
            "Difference",
        ]
    ]

def build_sap_view(agg: pd.DataFrame) -> pd.DataFrame:
    rows = []
    group_fields = ["Employee ID", "SAP Vendor ID", "Report ID", "Report Submit Date"]
    for _, group in agg.groupby(group_fields, sort=False):
        first = True
        for _, row in group.iterrows():
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
                "Text (M)": "",
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
) -> Path:
    raw_df = read_concur_file(path)
    comp = prepare_company_rows(raw_df.copy(), vendor_lookup, employee_lookup, cost_center_transform)
    validate_gst_rates(comp, region)
    agg = aggregate_rows(comp)
    agg = apply_region_tax_display(agg, region)
    sap_view = build_sap_view(agg)
    gst_check = build_gst_check(agg)
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
