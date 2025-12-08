' SAP FBL1N XXL export (attach to existing SAP session).
' Usage: cscript //Nologo export_xxl.vbs <CompanyCode> <KeyDate dd/MM/yyyy> <OutputFullPath>

Option Explicit

If WScript.Arguments.Count < 3 Then
    WScript.Echo "Usage: cscript //Nologo export_xxl.vbs <CompanyCode> <KeyDate dd/MM/yyyy> <OutputFullPath>"
    WScript.Quit 1
End If

Dim CompanyCode, KeyDate, OutFile
CompanyCode = WScript.Arguments(0)
KeyDate = WScript.Arguments(1)
OutFile = WScript.Arguments(2)

Dim SapGuiAuto, application, connection, session, grid

On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
    WScript.Echo "Error: Cannot get SAPGUI object."
    WScript.Quit 1
End If
On Error GoTo 0

Set application = SapGuiAuto.GetScriptingEngine
If application Is Nothing Then
    WScript.Echo "Error: No scripting engine."
    WScript.Quit 1
End If

If application.Connections.Count = 0 Then
    WScript.Echo "Error: No active SAP connections."
    WScript.Quit 1
End If

Set connection = application.Connections.Item(0)
If connection.Sessions.Count = 0 Then
    WScript.Echo "Error: No active SAP session."
    WScript.Quit 1
End If
Set session = connection.Sessions.Item(0)

' Navigate to FBL1N
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nfbl1n"
session.findById("wnd[0]/tbar[0]/btn[0]").press

' Set company code
On Error Resume Next
If Not session.findById("wnd[0]/usr/ctxtRF05L-BUKRS", False) Is Nothing Then
    session.findById("wnd[0]/usr/ctxtRF05L-BUKRS").Text = CompanyCode
ElseIf Not session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW", False) Is Nothing Then
    session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").Text = CompanyCode
Else
    WScript.Echo "Error: Company code field not found."
    WScript.Quit 1
End If
On Error GoTo 0

' Set key date using common fields
Dim dateParts, dateFormatted
dateParts = Split(KeyDate, "/")
If UBound(dateParts) = 2 Then
    dateFormatted = dateParts(0) & "." & dateParts(1) & "." & dateParts(2)
Else
    WScript.Echo "Error: KeyDate must be dd/MM/yyyy"
    WScript.Quit 1
End If

If Not session.findById("wnd[0]/usr/ctxtPA_STIDA", False) Is Nothing Then
    session.findById("wnd[0]/usr/ctxtPA_STIDA").Text = dateFormatted
ElseIf Not session.findById("wnd[0]/usr/ctxtRF05L-ALDAT", False) Is Nothing Then
    session.findById("wnd[0]/usr/ctxtRF05L-ALDAT").Text = dateFormatted
ElseIf Not session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW", False) Is Nothing Then
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").Text = dateFormatted
End If

' Open items checkbox
Dim cid
For Each cid In Array("wnd[0]/usr/chkRF05L-OPEN_ITEMS", "wnd[0]/usr/chkX_AKONT", "wnd[0]/usr/chkPARKED")
    On Error Resume Next
    If Not session.findById(cid, False) Is Nothing Then
        session.findById(cid).selected = True
        Exit For
    End If
    On Error GoTo 0
Next

' Execute
session.findById("wnd[0]/tbar[1]/btn[8]").press

' Grab grid
Dim gridCandidates, gid
gridCandidates = Array( _
    "wnd[0]/usr/cntlGRID1/shellcont/shell", _
    "wnd[0]/usr/cntlGRID1/shellcont/shellcont[1]/shell" _
)
Set grid = Nothing
For Each gid In gridCandidates
    On Error Resume Next
    Set grid = session.findById(gid, False)
    On Error GoTo 0
    If Not grid Is Nothing Then Exit For
Next
If grid Is Nothing Then
    WScript.Echo "Error: ALV grid not found."
    WScript.Quit 1
End If

' Trigger XXL
On Error Resume Next
Dim colTry
For Each colTry In Array("BELNR", "BUKRS", "DMSHB")
    If Err.Number <> 0 Then Err.Clear
    grid.setCurrentCell 0, colTry
    If Err.Number = 0 Then Exit For
Next
Err.Clear
grid.contextMenu
grid.selectContextMenuItem "&XXL"
If Err.Number <> 0 Then
    Err.Clear
    WScript.Echo "Error: could not open XXL export from grid."
    WScript.Quit 1
End If
On Error GoTo 0

' Handle path/filename dialog if present
If session.ActiveWindow.Name = "wnd[1]" Then
    On Error Resume Next
    If Not session.findById("wnd[1]/usr/ctxtDY_PATH", False) Is Nothing Then
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = Left(OutFile, InStrRev(OutFile, "\") - 1)
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = Mid(OutFile, InStrRev(OutFile, "\") + 1)
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
    On Error GoTo 0
End If

' Spreadsheet format radio (recorded ID)
Dim radios, rid
radios = Array( _
    "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]", _
    "wnd[1]/usr/radRB_0", _
    "wnd[1]/usr/radRB0" _
)
If session.ActiveWindow.Name = "wnd[1]" Then
    For Each rid In radios
        On Error Resume Next
        If Not session.findById(rid, False) Is Nothing Then
            session.findById(rid).select
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            Exit For
        End If
        On Error GoTo 0
    Next
End If

' Final OK if another dialog
If session.ActiveWindow.Name = "wnd[1]" Then
    On Error Resume Next
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
End If

WScript.Echo "Export complete to: " & OutFile
