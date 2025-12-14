# SAP Login Helper
**Category**: ops
**Version**: v0.1 (Released: 2025-12-14)

## What it does
- Opens a SAP Logon connection (by entry name) and ensures the session is logged in via SAP GUI scripting.
- Produces a machine-readable status file so other pipelines can depend on “SAP is ready”.

## Inputs
- Config (default): `01-system/configs/apis/API-Keys.md`
  - `SAP_LOGON_ENTRY`: SAP Logon entry/description (example: `Nova Biomedical Production System (ECP)`)
  - `SAP_CLIENT`: client (example: `800`)
  - `SAP_USER`: username (example: `AZHAO`)
  - `SAP_PASSWORD`: optional (prefer SSO; do not share in chat)

## Steps (routine)
1. Ensure SAP GUI scripting is enabled (client + server).
2. From the repo root run: `python 01-system/tools/ops/sap-login/sap_login.py`.
3. If successful, proceed to downstream tools (for example `sap-fbl1n`, payment runs, etc.).

## Outputs
- **Latest status**: `03-outputs/sap-login/latest.json`
- **Run history**: `03-outputs/sap-login/runs/<YYYYMMDD_HHMMSS>/result.json`

## Notes
- Never commit real passwords/keys; keep `API-Keys.md` redacted in git.
- If SAP is already logged in, the tool exits successfully and records `mode=already_logged_in`.

## Change Log
- v0.1 (2025-12-14): Initial release.

