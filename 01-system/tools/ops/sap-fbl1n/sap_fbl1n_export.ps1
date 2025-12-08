<#
SAP FBL1N export helper (SAP GUI scripting).

Usage (example):
  powershell -ExecutionPolicy Bypass -File 01-system/tools/ops/sap-fbl1n/sap_fbl1n_export.ps1 `
    -SapLogonEntry "ECP(1)" -Client "800" -User "azhao" `
    -CompanyCodes "8000","8100" -KeyDate "05/12/2025" -OpenItemsOnly `
    -OutputDir "02-inputs/downloads" -LayoutVariant ""

Notes:
- Requires SAP GUI for Windows with scripting enabled (client + server).
- Does not store credentials; provide password at prompt or via -Password (not logged).
- Exports will overwrite existing files: FBL1N_<bukrs>_<yyyymmdd>.xlsx in OutputDir.
- Field IDs are based on standard FBL1N; if a control is not found, record a short
  SAP GUI script for your system and adjust the element IDs below.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SapLogonEntry,
    [string]$Client = "800",
    [Parameter(Mandatory = $true)]
    [string]$User,
    [string]$Password,
    [string[]]$CompanyCodes = @("8000", "8100"),
    [Parameter(Mandatory = $true)]
    [string]$KeyDate, # expected format dd/MM/yyyy
    [switch]$OpenItemsOnly = $true,
    [string]$LayoutVariant = "",
    [string]$OutputDir = "02-inputs/downloads"
)

function Get-PlainPassword {
    param([string]$Pwd)
    if ($Pwd) { return $Pwd }
    $secure = Read-Host -Prompt "Enter SAP password (not logged)" -AsSecureString
    return (New-Object System.Net.NetworkCredential("", $secure)).Password
}

function New-SapController {
    # Try direct COM controller first
    $progIds = @("Sapgui.ScriptingCtrl.1", "SAPGUI.ScriptingCtrl.1")
    foreach ($id in $progIds) {
        $ctrl = [Activator]::CreateInstance([type]::GetTypeFromProgID($id))
        if ($ctrl) { return $ctrl }
    }
    # Fallback to GetObject("SAPGUI") to attach to existing engine (when BindToMoniker shows 0 connections)
    try {
        $sap = [Runtime.InteropServices.Marshal]::GetActiveObject("SAPGUI")
        if ($sap) {
            return $sap.GetScriptingEngine()
        }
    } catch {
        # ignore, will throw below
    }
    throw "SAP GUI Scripting API not available. Enable scripting in SAP GUI options and ensure SAP GUI is installed."
}

function Get-SapSession {
    param(
        [string]$LogonEntry,
        [string]$Client,
        [string]$User,
        [string]$PlainPassword
    )

    $controller = New-SapController

    # Reuse existing session if possible (prefer active connections)
    try {
        if ($controller.Connections.Count -gt 0) {
            $connection = $controller.Connections.Item(0)
            if ($connection.Sessions.Count -gt 0) {
                return $connection.Sessions.Item(0)
            }
        }
    } catch { }

    throw "No active SAP GUI session detected. Please log into '$LogonEntry' (client $Client) in SAP GUI (same user, not elevated) and rerun."
}

function Export-Fbl1n {
    param(
        $Session,
        [string]$CompanyCode,
        [datetime]$KeyDateObj,
        [switch]$OpenItemsOnly,
        [string]$LayoutVariant,
        [string]$OutputDir
    )

    $dateFormatted = $KeyDateObj.ToString("dd.MM.yyyy")
    $outfile = Join-Path -Path $OutputDir -ChildPath ("{0}.xlsx" -f $KeyDateObj.ToString("dd.MM.yy"))

    # Go to FBL1N
    $Session.findById("wnd[0]/tbar[0]/okcd").text = "/nfbl1n"
    $Session.findById("wnd[0]/tbar[0]/btn[0]").press()

    # Set selection
    $Session.findById("wnd[0]/usr/ctxtRF05L-BUKRS").text = $CompanyCode
    if ($Session.findById("wnd[0]/usr/ctxtPA_STIDA", $false)) {
        $Session.findById("wnd[0]/usr/ctxtPA_STIDA").text = $dateFormatted
    } elseif ($Session.findById("wnd[0]/usr/ctxtRF05L-ALDAT", $false)) {
        $Session.findById("wnd[0]/usr/ctxtRF05L-ALDAT").text = $dateFormatted
    } elseif ($Session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW", $false)) {
        $Session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = $dateFormatted
    }

    if ($OpenItemsOnly) {
        $checkboxIds = @("wnd[0]/usr/chkRF05L-OPEN_ITEMS", "wnd[0]/usr/chkX_AKONT", "wnd[0]/usr/chkPARKED")
        foreach ($cid in $checkboxIds) {
            if ($Session.findById($cid, $false)) {
                $Session.findById($cid).selected = $true
                break
            }
        }
    }

    if ($LayoutVariant -and $Session.findById("wnd[0]/usr/ctxtLAYOUT_DYN", $false)) {
        $Session.findById("wnd[0]/usr/ctxtLAYOUT_DYN").text = $LayoutVariant
    }

    # Execute
    $Session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Grab the ALV grid
    $grid = $Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell", $false)
    if (-not $grid) { throw "ALV grid not found; adjust grid ID if needed." }

    # Export to Excel (XXL)
    $grid.contextMenu()
    $grid.selectContextMenuItem("&XXL")

    Handle-ExportDialogs -Session $Session -OutFile $outfile

    Write-Host "Exported $CompanyCode to $outfile"
}

# Main
$plainPwd = Get-PlainPassword -Pwd $Password
$keyDateObj = [datetime]::ParseExact($KeyDate, "dd/MM/yyyy", $null)

if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

$session = Get-SapSession -LogonEntry $SapLogonEntry -Client $Client -User $User -PlainPassword $plainPwd

foreach ($cc in $CompanyCodes) {
    Export-Fbl1n -Session $session -CompanyCode $cc -KeyDateObj $keyDateObj -OpenItemsOnly:$OpenItemsOnly -LayoutVariant $LayoutVariant -OutputDir $OutputDir
}

Write-Host "Done."

function Handle-ExportDialogs {
    param(
        $Session,
        [string]$OutFile
    )

    $resolvedDir = Resolve-Path (Split-Path -Parent $OutFile)
    $fileName = [System.IO.Path]::GetFileName($OutFile)

    $radioIds = @(
        "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]",
        "wnd[1]/usr/radRB_0",
        "wnd[1]/usr/radRB0"
    )

    for ($i = 0; $i -lt 5; $i++) {
        Start-Sleep -Milliseconds 300
        if (-not $Session.ActiveWindow) { break }
        if ($Session.ActiveWindow.Name -ne "wnd[1]") { break }

        $handled = $false

        if ($Session.findById("wnd[1]/usr/ctxtDY_PATH", $false)) {
            $Session.findById("wnd[1]/usr/ctxtDY_PATH").text = $resolvedDir
            $Session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = $fileName
            $Session.findById("wnd[1]/tbar[0]/btn[0]").press()
            $handled = $true
        }

        if (-not $handled) {
            foreach ($rid in $radioIds) {
                $rb = $Session.findById($rid, $false)
                if ($rb) {
                    $rb.select()
                    $Session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    $handled = $true
                    break
                }
            }
        }

        if (-not $handled -and $Session.findById("wnd[1]/tbar[0]/btn[0]", $false)) {
            $Session.findById("wnd[1]/tbar[0]/btn[0]").press()
            $handled = $true
        }

        if (-not $handled) { break }
    }
}
