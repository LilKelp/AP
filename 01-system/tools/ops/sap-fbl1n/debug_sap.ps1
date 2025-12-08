# Debug SAP GUI Scripting - Method 4 (BindToMoniker)

Write-Host "Attempting to grab running SAPGUI object via BindToMoniker..."

try {
    $SapGuiAuto = [System.Runtime.InteropServices.Marshal]::BindToMoniker("SAPGUI")
}
catch {
    Write-Host "Error: Could not get active SAPGUI object." -ForegroundColor Red
    Write-Host "Details: $_" -ForegroundColor Yellow
    exit
}

if (-not $SapGuiAuto) {
    Write-Host "Error: SAPGUI object is null." -ForegroundColor Red
    exit
}

try {
    $engine = $SapGuiAuto.GetScriptingEngine
}
catch {
    Write-Host "Error: Could not get Scripting Engine." -ForegroundColor Red
    exit
}

Write-Host "SAP GUI Scripting Engine found." -ForegroundColor Green
Write-Host "Connections count: $($engine.Connections.Count)"

# Note: SAP collections in PowerShell sometimes need special handling
$conns = $engine.Connections
for ($i = 0; $i -lt $conns.Count; $i++) {
    # Access via Children if Item fails, or just try Item
    try {
        $conn = $conns.Item($i)
    }
    catch {
        $conn = $conns.Children.Item($i)
    }
    
    Write-Host "  Connection [$i]: $($conn.Description)"
    
    $sessions = $conn.Sessions
    for ($j = 0; $j -lt $sessions.Count; $j++) {
        try {
            $sess = $sessions.Item($j)
        }
        catch {
            $sess = $sessions.Children.Item($j)
        }
        Write-Host "      Session [$j]: ID=$($sess.Id), Info=$($sess.Info.SystemName) client=$($sess.Info.Client)"
    }
}

Write-Host "`nDone."
