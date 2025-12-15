' SAP FBL1N Export Script (VBScript)
' Version: 13.0 (Local file or Spreadsheet save)
' Usage: cscript //Nologo sap_fbl1n_export.vbs <CompanyCode> <KeyDate> <OutputDir> [LayoutVariant] [dump] [radio=<id>] [mode=localfile|spreadsheet]

Option Explicit

Dim CompanyCode, KeyDate, OutputDir, LayoutVariant
Dim SapGuiAuto, Application, Connection, Session
Dim Grid, fso
Dim RadioOverride
Dim ExportMode

' Map company code to subfolder (AU/NZ) for payment run raw saves
Function GetRepoRoot()
    Dim d
    d = fso.GetParentFolderName(WScript.ScriptFullName) ' sap-fbl1n
    d = fso.GetParentFolderName(d) ' ops
    d = fso.GetParentFolderName(d) ' tools
    d = fso.GetParentFolderName(d) ' 01-system
    d = fso.GetParentFolderName(d) ' repo root
    GetRepoRoot = d
End Function

Function IsAbsolutePath(p)
    Dim t
    t = Trim(p)
    IsAbsolutePath = False
    If t = "" Then Exit Function
    If Left(t, 2) = "\\" Then
        IsAbsolutePath = True
        Exit Function
    End If
    If Len(t) >= 3 Then
        If Mid(t, 2, 2) = ":\\" Then
            IsAbsolutePath = True
            Exit Function
        End If
    End If
End Function

Function ToAbsolutePath(p)
    Dim t, root
    t = Trim(p)
    If IsAbsolutePath(t) Then
        ToAbsolutePath = fso.GetAbsolutePathName(t)
    Else
        root = GetRepoRoot()
        ToAbsolutePath = fso.GetAbsolutePathName(root & "\" & t)
    End If
End Function

Function ResolveSaveDir(baseDir, companyCode)
    Dim cc, target
    cc = Trim(companyCode)
    target = ToAbsolutePath(baseDir)
    If Right(target, 1) = "\" Then
        target = Left(target, Len(target) - 1)
    End If
    If LCase(Right(target, Len("payment run raw"))) = "payment run raw" Then
        If cc = "8000" Then
            target = target & "\AU"
        ElseIf cc = "8100" Then
            target = target & "\NZ"
        Else
            target = target & "\" & cc
        End If
    End If
    ResolveSaveDir = target
End Function

WScript.Echo "Starting VBScript Version 13.0"

' Arguments
If WScript.Arguments.Count < 3 Then
    WScript.Echo "Usage: cscript sap_fbl1n_export.vbs <CompanyCode> <KeyDate> <OutputDir> [LayoutVariant] [dump] [radio=<id>] [mode=localfile|spreadsheet]"
    WScript.Quit 1
End If

CompanyCode = WScript.Arguments(0)
KeyDate = WScript.Arguments(1)
OutputDir = WScript.Arguments(2)

LayoutVariant = ""
RadioOverride = ""
ExportMode = "localfile" ' default

Dim DumpRadios
DumpRadios = False

Dim idxArg, argVal, argLower
For idxArg = 3 To WScript.Arguments.Count - 1
    argVal = WScript.Arguments(idxArg)
    argLower = LCase(argVal)
    If argLower = "dump" Then
        DumpRadios = True
    ElseIf Left(argLower, 6) = "radio=" Or Left(argLower, 3) = "id=" Then
        If Left(argLower, 6) = "radio=" Then
            RadioOverride = Mid(argVal, 7)
        Else
            RadioOverride = Mid(argVal, 4)
        End If
    ElseIf Left(argLower, 5) = "mode=" Then
        ExportMode = LCase(Mid(argLower, 6))
    ElseIf LayoutVariant = "" Then
        LayoutVariant = argVal
    End If
Next

Set fso = CreateObject("Scripting.FileSystemObject")

' Format Date
Dim DateParts, DateFormatted, DateFile
DateParts = Split(KeyDate, "/")
If UBound(DateParts) <> 2 Then
    WScript.Echo "Error: Invalid date format. Use dd/MM/yyyy"
    WScript.Quit 1
End If
DateFormatted = DateParts(0) & "." & DateParts(1) & "." & DateParts(2)
DateFile = DateParts(0) & "." & DateParts(1) & "." & Right(DateParts(2), 2)

' Connect to SAP
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
    WScript.Echo "Error: Could not get SAPGUI."
    WScript.Quit 1
End If
Set Application = SapGuiAuto.GetScriptingEngine
On Error Goto 0

If Application.Connections.Count = 0 Then
    WScript.Echo "Error: No active SAP connection."
    WScript.Quit 1
End If

Set Connection = Application.Connections.Item(0)
If Connection.Sessions.Count = 0 Then
    WScript.Echo "Error: No active SAP session."
    WScript.Quit 1
End If
Set Session = Connection.Sessions.Item(0)

' Navigate to FBL1N
Session.findById("wnd[0]/tbar[0]/okcd").text = "/nfbl1n"
Session.findById("wnd[0]/tbar[0]/btn[0]").press

' Set Company Code
On Error Resume Next
Session.findById("wnd[0]/usr/ctxtRF05L-BUKRS").text = CompanyCode
If Err.Number <> 0 Then
    Err.Clear
    Session.findById("wnd[0]/usr/ctxtKD_BUKRS-LOW").text = CompanyCode
End If
On Error Goto 0

' Set Date - "Open at key date"
On Error Resume Next

' First, CLEAR other date fields that might be pre-filled
If Not Session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW", False) Is Nothing Then
    Session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = ""
End If
If Not Session.findById("wnd[0]/usr/ctxtRF05L-ALDAT", False) Is Nothing Then
    Session.findById("wnd[0]/usr/ctxtRF05L-ALDAT").text = ""
End If
' Now set the Open at Key Date (several possible field IDs)
Dim dateSet
dateSet = False
If Not Session.findById("wnd[0]/usr/ctxtPA_STIDA", False) Is Nothing Then
    Session.findById("wnd[0]/usr/ctxtPA_STIDA").text = DateFormatted
    dateSet = True
ElseIf Not Session.findById("wnd[0]/usr/ctxtRF05L-STIDA", False) Is Nothing Then
    Session.findById("wnd[0]/usr/ctxtRF05L-STIDA").text = DateFormatted
    dateSet = True
ElseIf Not Session.findById("wnd[0]/usr/ctxtRF05L-ALDAT", False) Is Nothing Then
    Session.findById("wnd[0]/usr/ctxtRF05L-ALDAT").text = DateFormatted
    dateSet = True
ElseIf Not Session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW", False) Is Nothing Then
    Session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = DateFormatted
    dateSet = True
End If
If Not dateSet Then
    WScript.Echo "Warning: Could not set key date (no matching field)."
End If
On Error Goto 0

' Open Items checkbox
Dim CheckboxIds, Cid
CheckboxIds = Array("wnd[0]/usr/chkRF05L-OPEN_ITEMS", "wnd[0]/usr/chkX_AKONT", "wnd[0]/usr/chkPARKED")
For Each Cid In CheckboxIds
    On Error Resume Next
    If Not Session.findById(Cid, False) Is Nothing Then
        Session.findById(Cid).selected = True
        Exit For
    End If
    On Error Goto 0
Next

' Layout
If LayoutVariant <> "" Then
    On Error Resume Next
    If Not Session.findById("wnd[0]/usr/ctxtLAYOUT_DYN", False) Is Nothing Then
        Session.findById("wnd[0]/usr/ctxtLAYOUT_DYN").text = LayoutVariant
    End If
    On Error Goto 0
End If

' Execute
Session.findById("wnd[0]/tbar[1]/btn[8]").press

' Wait for Grid
Dim gridTries
gridTries = 60
On Error Resume Next
Do While gridTries > 0
    Set Grid = Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell", False)
    If Not Grid Is Nothing Then Exit Do
    WScript.Sleep 5000
    gridTries = gridTries - 1
Loop
On Error Goto 0

If Grid Is Nothing Then
    WScript.Echo "Error: ALV Grid not found."
    WScript.Quit 1
End If

WScript.Echo "Grid found. Ready to export..."

' Trigger export
If ExportMode <> "spreadsheet" And ExportMode <> "localfile" And ExportMode <> "file" Then
    WScript.Echo "Warning: Unknown export mode '" & ExportMode & "'. Defaulting to localfile."
    ExportMode = "localfile"
End If

If ExportMode = "spreadsheet" Then
    WScript.Echo "Export mode: Spreadsheet -> Excel"
    If Not ExportSpreadsheet(Session, OutputDir, CompanyCode, DateParts, DumpRadios, RadioOverride) Then
        WScript.Echo "Error: Spreadsheet export failed."
        WScript.Quit 1
    End If
Else
    WScript.Echo "Export mode: Local File -> Excel"
    If Not ExportLocalFile(Session, OutputDir, CompanyCode, DateParts, DumpRadios, RadioOverride) Then
        WScript.Echo "Warning: Local file export failed, falling back to Spreadsheet export."
        If Not ExportSpreadsheet(Session, OutputDir, CompanyCode, DateParts, DumpRadios, RadioOverride) Then
            WScript.Echo "Error: Spreadsheet export failed."
            WScript.Quit 1
        End If
    End If
End If

WScript.Echo "SAP_EXPORT_DONE"
WScript.Quit 0

' -------- Helpers --------
Function ExportSpreadsheet(Session, OutputDir, CompanyCode, DateParts, DumpRadios, RadioOverride)
    ExportSpreadsheet = False
    Dim dateStamp, fileName, saveDir, dialogFound, idx
    dateStamp = DateParts(2) & DateParts(1) & DateParts(0) ' yyyymmdd
    fileName = "FBL1N_" & CompanyCode & "_" & dateStamp & ".xlsx"
    saveDir = ResolveSaveDir(OutputDir, CompanyCode)

    If Not fso.FolderExists(saveDir) Then
        On Error Resume Next
        fso.CreateFolder saveDir
        On Error Goto 0
    End If

    On Error Resume Next
    Session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select ' Spreadsheet
    On Error Goto 0

    ' Clear any immediate Information dialog (e.g., reminders) and continue
    HandleInfoDialog Session

    If Not SelectSpreadsheetRadio(Session, DumpRadios, RadioOverride) Then
        WScript.Echo "Warning: Spreadsheet selection dialog not found; continuing with SendKeys fallback."
    End If

    ' Some systems show a Processing Mode dialog (Table vs Pivot); prefer Table
    SelectProcessingMode Session

    ' Wait for Save dialog and fill path/file (handle repeat info popups)
    For idx = 1 To 60
        WScript.Sleep 250
        If Not Session.ActiveWindow Is Nothing Then
            If Session.ActiveWindow.Name = "wnd[1]" Then
                If InStr(LCase(Session.ActiveWindow.Text), "information") > 0 Then
                    On Error Resume Next
                    If Not Session.findById("wnd[1]/tbar[0]/btn[0]", False) Is Nothing Then
                        Session.findById("wnd[1]/tbar[0]/btn[0]").press
                    End If
                    On Error GoTo 0
                Else
                    ' Save dialog
                    If Not Session.findById("wnd[1]/usr/ctxtDY_PATH", False) Is Nothing Then
                        Session.findById("wnd[1]/usr/ctxtDY_PATH").text = saveDir
                    End If
                    If Not Session.findById("wnd[1]/usr/ctxtDY_FILENAME", False) Is Nothing Then
                        Session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fileName
                    End If
                    If Not Session.findById("wnd[1]/tbar[0]/btn[0]", False) Is Nothing Then
                        Session.findById("wnd[1]/tbar[0]/btn[0]").press
                        ExportSpreadsheet = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next

    ' Fallback to Windows Save As via SendKeys when no SAP save dialog appears.
    If TrySaveViaSendKeys(saveDir, fileName) Then
        ExportSpreadsheet = True
        Exit Function
    End If
End Function

Function ExportLocalFile(Session, OutputDir, CompanyCode, DateParts, DumpRadios, RadioOverride)
    ExportLocalFile = False
    Dim fileName, saveDir, fullPathLocal, idx
    fileName = DateParts(0) & "." & DateParts(1) & "." & Right(DateParts(2), 2) & ".xls"
    saveDir = ResolveSaveDir(OutputDir, CompanyCode)

    If Not fso.FolderExists(saveDir) Then
        On Error Resume Next
        fso.CreateFolder saveDir
        On Error Goto 0
    End If
    fullPathLocal = fso.GetAbsolutePathName(saveDir) & "\" & fileName

    On Error Resume Next
    ' Local file... (F9) under List -> Export
    Session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select
    On Error Goto 0

    ' If a format selection dialog appears first, try to pick Spreadsheet/Excel.
    If Not Session.ActiveWindow Is Nothing Then
        If Session.ActiveWindow.Name = "wnd[1]" Then
            If Session.findById("wnd[1]/usr/ctxtDY_PATH", False) Is Nothing Then
                Call SelectSpreadsheetRadio(Session, DumpRadios, RadioOverride)
                SelectProcessingMode Session
            End If
        End If
    End If

    ' Wait for Save dialog and fill DY_PATH/DY_FILENAME
    For idx = 1 To 80
        WScript.Sleep 250
        If Not Session.ActiveWindow Is Nothing Then
            If Session.ActiveWindow.Name = "wnd[1]" Then
                Dim lowerText
                lowerText = LCase(Session.ActiveWindow.Text)
                If InStr(lowerText, "information") > 0 Then
                    On Error Resume Next
                    If Not Session.findById("wnd[1]/tbar[0]/btn[0]", False) Is Nothing Then
                        Session.findById("wnd[1]/tbar[0]/btn[0]").press
                    End If
                    On Error GoTo 0
                ElseIf Not Session.findById("wnd[1]/usr/ctxtDY_PATH", False) Is Nothing Then
                    Session.findById("wnd[1]/usr/ctxtDY_PATH").text = saveDir
                    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fileName
                    Session.findById("wnd[1]/tbar[0]/btn[0]").press
                    WScript.Sleep 1000
                    If LooksLikeAsciiList(fullPathLocal) Then
                        WScript.Echo "Info: Local file export produced a text list (still usable by payment-list)."
                    End If
                    ExportLocalFile = True
                    Exit Function
                End If
            End If
        End If
    Next

    ' Fallback: some SAP setups open a Windows "Save As" dialog (not visible to SAP scripting).
    If TrySaveViaSendKeys(saveDir, fileName) Then
        If LooksLikeAsciiList(fullPathLocal) Then
            WScript.Echo "Info: Local file export produced a text list (still usable by payment-list)."
        End If
        ExportLocalFile = True
        Exit Function
    End If

    WScript.Echo "Error: Local file save dialog not found."
End Function

Function TrySaveViaSendKeys(saveDir, fileName)
    TrySaveViaSendKeys = False
    Dim sh, fullPath, i
    Set sh = CreateObject("WScript.Shell")
    fullPath = fso.GetAbsolutePathName(saveDir) & "\" & fileName

    ' Try to bring the Save As dialog to foreground
    If Not sh.AppActivate("Save As") Then
        sh.AppActivate("Save list in file")
    End If
    WScript.Sleep 300

    ' Try to select Excel file type, then type full path into filename field.
    On Error Resume Next
    sh.SendKeys "%t" ' Alt+T -> Save as type
    WScript.Sleep 150
    sh.SendKeys "e"  ' EXCEL Files (*.xls)
    WScript.Sleep 150
    sh.SendKeys "%n" ' Alt+N -> File name
    WScript.Sleep 150
    On Error GoTo 0

    sh.SendKeys fullPath
    sh.SendKeys "{ENTER}"

    ' Give SAP/Windows a moment to write the file
    For i = 1 To 20
        WScript.Sleep 250
        If fso.FileExists(fullPath) Then
            TrySaveViaSendKeys = True
            Exit Function
        End If
    Next
End Function

Function LooksLikeAsciiList(fullPath)
    LooksLikeAsciiList = False
    On Error Resume Next
    Dim xl, wb, ws, colCount, r, val
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Open(fullPath)
    If Not wb Is Nothing Then
        Set ws = wb.Worksheets(1)
        colCount = ws.UsedRange.Columns.Count
        If colCount = 1 Then
            For r = 1 To 30
                val = CStr(ws.Cells(r, 1).Value)
                If InStr(val, "|") > 0 Then
                    LooksLikeAsciiList = True
                    Exit For
                End If
            Next
        End If
        wb.Close False
    End If
    xl.Quit
    Set wb = Nothing
    Set xl = Nothing
    On Error GoTo 0
End Function

Function SelectSpreadsheetRadio(Session, DumpRadios, RadioOverride)
    Dim dialog, idx, targetId, fallbackIds, fid, rb
    SelectSpreadsheetRadio = False
    
    ' Wait for dialog (skip transient Information popups)
    idx = 0
    Do While idx < 30
        idx = idx + 1
        WScript.Sleep 250
        If Not Session.ActiveWindow Is Nothing Then
            If Session.ActiveWindow.Name = "wnd[1]" Then
                Dim lowerText
                lowerText = LCase(Session.ActiveWindow.Text)
                If InStr(lowerText, "information") > 0 Or InStr(lowerText, "export list object") > 0 Then
                    On Error Resume Next
                    If Not Session.findById("wnd[1]/tbar[0]/btn[0]", False) Is Nothing Then
                        Session.findById("wnd[1]/tbar[0]/btn[0]").press
                    End If
                    On Error GoTo 0
                Else
                    Exit Do
                End If
            End If
        End If
    Loop
    
    If Session.ActiveWindow Is Nothing Or Session.ActiveWindow.Name <> "wnd[1]" Then
        WScript.Echo "Error: Export dialog not found."
        Exit Function
    End If

    Set dialog = Session.ActiveWindow
    WScript.Echo "Dialog: " & dialog.Text
    If DumpRadios Then DumpControls dialog

    ' User override first
    If RadioOverride <> "" Then
        targetId = RadioOverride
        On Error Resume Next
        Set rb = Session.findById(targetId, False)
        If Not rb Is Nothing Then
            rb.select
            If rb.selected = True Then
                Session.findById("wnd[1]/tbar[0]/btn[0]").press
                SelectSpreadsheetRadio = True
                Exit Function
            End If
        End If
        Err.Clear
        On Error GoTo 0
        WScript.Echo "Radio override failed: " & targetId
        Exit Function
    End If

    fallbackIds = Array( _
        "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]", _
        "wnd[1]/usr/radRB_0", "wnd[1]/usr/radRB0", _
        "wnd[1]/usr/radRB_1", "wnd[1]/usr/radRB1" _
    )

    For Each fid In fallbackIds
        On Error Resume Next
        Set rb = Session.findById(fid, False)
        If Not rb Is Nothing Then
            rb.select
            If rb.selected = True Then
                Session.findById("wnd[1]/tbar[0]/btn[0]").press
                SelectSpreadsheetRadio = True
                Exit Function
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next

    ' Fallback to "OK" if no radio found
    If Session.findById("wnd[1]/tbar[0]/btn[0]", False) Is Nothing Then
        WScript.Echo "Error: No radio button selected via fallback IDs."
    Else
        Session.findById("wnd[1]/tbar[0]/btn[0]").press
        SelectSpreadsheetRadio = True
    End If
End Function

Sub HandleInfoDialog(Session)
    Dim i
    For i = 1 To 5
        If Session.ActiveWindow Is Nothing Then Exit For
        If Session.ActiveWindow.Name = "wnd[1]" Then
            On Error Resume Next
            WScript.Echo "Info dialog: " & Session.ActiveWindow.Text
            If Not Session.findById("wnd[1]/tbar[0]/btn[0]", False) Is Nothing Then
                Session.findById("wnd[1]/tbar[0]/btn[0]").press
            End If
            On Error GoTo 0
        End If
        WScript.Sleep 200
    Next
End Sub

Sub SelectProcessingMode(Session)
    ' Optional dialog after spreadsheet selection to choose Table vs Pivot
    Dim idx, dialog, radios, rid, rb, picked
    Set dialog = Nothing
    For idx = 1 To 20
        WScript.Sleep 200
        If Session.ActiveWindow Is Nothing Then Exit For
        If Session.ActiveWindow.Name = "wnd[1]" Then
            Set dialog = Session.ActiveWindow
            Exit For
        End If
    Next
    If dialog Is Nothing Then Exit Sub
    On Error Resume Next
    radios = Array( _
        "wnd[1]/usr/radRB_0", "wnd[1]/usr/radRB0", _
        "wnd[1]/usr/radRB_1", "wnd[1]/usr/radRB1", _
        "wnd[1]/usr/radSPOPLI-SELFLAG[0,0]" _
    )
    ' Prefer Table/Microsoft Excel, avoid pivot text if present
    picked = False
    For Each rid In radios
        Set rb = Session.findById(rid, False)
        If Not rb Is Nothing Then
            Dim txt
            txt = LCase(rb.Text)
            If InStr(txt, "pivot") = 0 Then
                If InStr(txt, "table") > 0 Or InStr(txt, "microsoft excel") > 0 Or InStr(txt, "excel") > 0 Then
                    rb.select
                    picked = True
                    Exit For
                End If
            End If
        End If
    Next
    ' If nothing matched, pick first available radio
    If Not picked Then
        For Each rid In radios
            Set rb = Session.findById(rid, False)
            If Not rb Is Nothing Then
                rb.select
                picked = True
                Exit For
            End If
        Next
    End If
    If Not Session.findById("wnd[1]/tbar[0]/btn[0]", False) Is Nothing Then
        Session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
    On Error GoTo 0
End Sub

Sub DumpRadioButtons(dialog)
    Dim stack(), stackSize, node, i
    ReDim stack(0)
    stackSize = 1
    Set stack(0) = dialog
    WScript.Echo "Radio options in dialog:"
    
    Do While stackSize > 0
        Set node = stack(stackSize - 1)
        stackSize = stackSize - 1
        
        For i = 0 To node.Children.Count - 1
            On Error Resume Next
            Dim child, typeName, textVal
            Set child = node.Children.Item(i)
            typeName = ""
            textVal = ""
            typeName = child.Type
            textVal = child.Text
            On Error GoTo 0
            
            If InStr(typeName, "GuiRadioButton") > 0 Then
                WScript.Echo "  " & child.Id & " | " & textVal
            End If
            
            ' push child for traversal
            ReDim Preserve stack(stackSize)
            Set stack(stackSize) = child
            stackSize = stackSize + 1
        Next
    Loop
End Sub

Sub DumpControls(dialog)
    Dim stack(), stackSize, node, i
    ReDim stack(0)
    stackSize = 1
    Set stack(0) = dialog
    WScript.Echo "All controls in dialog:"
    
    Do While stackSize > 0
        Set node = stack(stackSize - 1)
        stackSize = stackSize - 1
        
        If (Not node Is Nothing) Then
            Dim childCount
            childCount = 0
            On Error Resume Next
            childCount = node.Children.Count
            If Err.Number <> 0 Then
                Err.Clear
                childCount = 0
            End If
            On Error GoTo 0
            
            If childCount > 0 Then
                For i = 0 To childCount - 1
                    On Error Resume Next
                    Dim child, typeName, textVal
                    Set child = node.Children.Item(i)
                    typeName = ""
                    textVal = ""
                    If Not (child Is Nothing) Then
                        typeName = child.Type
                        textVal = child.Text
                    End If
                    On Error GoTo 0
                    
                    If Not (child Is Nothing) Then
                        WScript.Echo "  " & child.Id & " | " & typeName & " | " & textVal
                        
                        ReDim Preserve stack(stackSize)
                        Set stack(stackSize) = child
                        stackSize = stackSize + 1
                    End If
                Next
            End If
        End If
    Loop
End Sub

Function FindRadioByText(dialog, keyword)
    Dim stack(), stackSize, node, i
    ReDim stack(0)
    stackSize = 1
    Set stack(0) = dialog
    keyword = LCase(keyword)
    Set FindRadioByText = Nothing
    
    Do While stackSize > 0
        Set node = stack(stackSize - 1)
        stackSize = stackSize - 1
        
        For i = 0 To node.Children.Count - 1
            On Error Resume Next
            Dim child, typeName, textVal
            Set child = node.Children.Item(i)
            typeName = ""
            textVal = ""
            typeName = child.Type
            textVal = child.Text
            On Error GoTo 0
            
            If InStr(typeName, "GuiRadioButton") > 0 Then
                If InStr(LCase(textVal), keyword) > 0 Then
                    Set FindRadioByText = child
                    Exit Function
                End If
            End If
            
            ReDim Preserve stack(stackSize)
            If Not (child Is Nothing) Then
                Set stack(stackSize) = child
                stackSize = stackSize + 1
            End If
        Next
    Loop
End Function
