' SAP FBL1N Export Script (VBScript)
' Version: 11.1 (Local File -> Clipboard)
' Usage: cscript //Nologo sap_fbl1n_export.vbs <CompanyCode> <KeyDate> <OutputDir> [LayoutVariant] [dump] [radio=<id>]

Option Explicit

Dim CompanyCode, KeyDate, OutputDir, LayoutVariant
Dim SapGuiAuto, Application, Connection, Session
Dim Grid, fso
Dim RadioOverride

WScript.Echo "Starting VBScript Version 11.0"

' Arguments
If WScript.Arguments.Count < 3 Then
    WScript.Echo "Usage: cscript sap_fbl1n_export.vbs <CompanyCode> <KeyDate> <OutputDir> [LayoutVariant] [dump]"
    WScript.Quit 1
End If

CompanyCode = WScript.Arguments(0)
KeyDate = WScript.Arguments(1)
OutputDir = WScript.Arguments(2)

LayoutVariant = ""
RadioOverride = ""

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

WScript.Echo "Grid found. Exporting to Clipboard..."

' Trigger Local File Export: List -> Export -> Local File
On Error Resume Next
Session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select
On Error Goto 0

WScript.Sleep 1000

' Select "In the clipboard" Format (robust selection with optional dump)
If Not SelectClipboardRadio(Session, DumpRadios, RadioOverride) Then
    WScript.Echo "Error: Could not select clipboard export option."
    WScript.Quit 1
End If

' Handle "Information" dialog (Data copied to clipboard)
WScript.Sleep 1000
Dim i
For i = 1 To 10
    On Error Resume Next
    If Session.ActiveWindow.Name = "wnd[1]" Then
        WScript.Echo "Handling Info Dialog: " & Session.ActiveWindow.Text
        Session.findById("wnd[1]/tbar[0]/btn[0]").press ' Continue (Green check)
        Exit For
    End If
    On Error Goto 0
    WScript.Sleep 500
Next

WScript.Echo "SAP_EXPORT_DONE"
WScript.Quit 0

' -------- Helpers --------
Function SelectClipboardRadio(Session, DumpRadios, RadioOverride)
    Dim dialog, idx, targetId, fallbackIds, fid, rb
    SelectClipboardRadio = False
    
    ' Wait briefly for dialog to appear
    For idx = 1 To 20
        WScript.Sleep 250
        If Not Session.ActiveWindow Is Nothing Then
            If Session.ActiveWindow.Name = "wnd[1]" Then Exit For
        End If
    Next
    
    If Session.ActiveWindow Is Nothing Or Session.ActiveWindow.Name <> "wnd[1]" Then
        WScript.Echo "Error: Export dialog not found."
        Exit Function
    End If
    
    Set dialog = Session.ActiveWindow
    WScript.Echo "Dialog: " & dialog.Text
    
    If DumpRadios Then DumpControls dialog
    
    ' If user provided a radio ID, try it first (no further scanning on failure)
    If RadioOverride <> "" Then
        targetId = RadioOverride
        On Error Resume Next
        Set rb = Session.findById(targetId, False)
        If Not rb Is Nothing Then
            rb.select
            If rb.selected = True Then
                Session.findById("wnd[1]/tbar[0]/btn[0]").press
                SelectClipboardRadio = True
                Exit Function
            End If
        End If
        Err.Clear
        On Error GoTo 0
        WScript.Echo "Radio override failed: " & targetId
        Exit Function ' do not continue scanning to avoid recursion errors; we can retry with dump output
    End If
    
    ' Try common radio IDs in order (underscore and no-underscore)
    fallbackIds = Array( _
        "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]", _
        "wnd[1]/usr/radRB_6", "wnd[1]/usr/radRB_5", "wnd[1]/usr/radRB_4", _
        "wnd[1]/usr/radRB_3", "wnd[1]/usr/radRB_2", "wnd[1]/usr/radRB_1", "wnd[1]/usr/radRB_0", _
        "wnd[1]/usr/radRB6", "wnd[1]/usr/radRB5", "wnd[1]/usr/radRB4", "wnd[1]/usr/radRB3", "wnd[1]/usr/radRB2", "wnd[1]/usr/radRB1", "wnd[1]/usr/radRB0" _
    )
    For Each fid In fallbackIds
        On Error Resume Next
        Set rb = Session.findById(fid, False)
        If Not rb Is Nothing Then
            rb.select
            If rb.selected = True Then
                Session.findById("wnd[1]/tbar[0]/btn[0]").press
                SelectClipboardRadio = True
                Exit Function
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next
    
    WScript.Echo "Error: No radio button selected via fallback IDs."
End Function

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
