# SAP FBL1N Export
**Category**: ops  
**Version**: v0.1 (Released: 2025-12-02)

## What it does
- Uses SAP GUI scripting to run FBL1N for specified company codes and exports the ALV grid to Excel.

## Inputs
- `SapLogonEntry`: Name of the SAP Logon entry (as in SAP Logon pad).
- `Client`: SAP client (default 800).
- `User`: SAP user ID.
- `Password`: Prompted if not provided (not stored).
- `CompanyCodes`: Defaults to `8000, 8100`.
- `KeyDate`: Key date `dd/MM/yyyy` (e.g., `05/12/2025`).
- `OpenItemsOnly`: Set by default.
- `LayoutVariant`: Optional ALV variant name.
- `OutputDir`: Defaults to `02-inputs/downloads/` (files overwrite).

## Steps (routine)
1. Ensure SAP GUI scripting is enabled (client and server).
2. Run:
   - `powershell -ExecutionPolicy Bypass -File 01-system/tools/ops/sap-fbl1n/sap_fbl1n_export.ps1 -SapLogonEntry "<ENTRY>" -Client "800" -User "<user>" -CompanyCodes "8000","8100" -KeyDate "05/12/2025" -OpenItemsOnly -OutputDir "02-inputs/downloads"`
3. Enter your SAP password when prompted (not logged).
4. Collect the exported files from `02-inputs/downloads/`.

## Outputs
- `02-inputs/downloads/FBL1N_<bukrs>_<yyyymmdd>.xlsx`

## Notes
- Field IDs are based on standard FBL1N. If a control is not found, record a short SAP GUI script and update the IDs inside `sap_fbl1n_export.ps1`.
- Files are overwritten on each run.

## Troubleshooting
- If scripting is blocked, enable “Allow Scripting” in SAP GUI options and ensure the server permits scripting.
- If control IDs differ, use SAP GUI Script Recorder on FBL1N and adjust the IDs in the script.

## Change Log
- v0.1 (2025-12-02): initial version
