If Not IsObject(Application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If

WScript.Echo "SAP GUI Scripting Engine found."
WScript.Echo "Connections count: " & application.Connections.Count

For Each Connection In application.Connections
    WScript.Echo "  Connection: " & Connection.Description
    WScript.Echo "    Sessions count: " & Connection.Sessions.Count
    
    For Each Session In Connection.Sessions
        WScript.Echo "      Session: ID=" & Session.Id & ", Info=" & Session.Info.SystemName & " (Client " & Session.Info.Client & ")"
    Next
Next

WScript.Echo "Done."
