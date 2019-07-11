On Error Resume Next
Const PAGE_LOADED = 4
Set objShell = CreateObject("WScript.shell")
objShell.AppActivate("Internet Explorer")
Set IE = CreateObject("InternetExplorer.Application")
Call IE.Navigate("WEBSITE GOES HERE")
IE.Visible = True
Do Until IE.ReadyState = PAGE_LOADED : Call WScript.Sleep(100) : Loop
IE.Document.all.User.Value = "USERNAME GOES HERE"
IE.Document.all.Password.Value = "PASSWORD GOES HERE"
If Err.Number <> 0 Then
msgbox "Error: " & err.Description
End If
Call IE.Document.Forms(0).Submit()
Set IE = Nothing