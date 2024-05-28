Set config = CreateObject("Scripting.Dictionary")

configFilePath = "config.txt"

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(configFilePath) Then
    Set file = fso.OpenTextFile(configFilePath, 1)
    
    Do Until file.AtEndOfStream
        line = Trim(file.ReadLine)
        If line <> "" And Left(line, 1) <> ";" Then
            arr = Split(line, "=")
            If UBound(arr) = 1 Then
                key = Trim(arr(0))
                value = Trim(arr(1))
                config(key) = value
            End If
        End If
    Loop
    
    file.Close
Else
    MsgBox "Configuration file not found: " & configFilePath
End If

' Check if username is set
If config("username") = "YOUR 6+2" Then
  MsgBox "Username is not set! Please change it in " & configFilePath
  WScript.Quit
End If
' Check if password is set
If config("password") = "YOUR PASSWORD" Then
  MsgBox "Password is not set! Please change it in " & configFilePath
  WScript.Quit
End If

Dim delays(2)

If config("delay") = "SHORT" Then
  delays(0) = 1000
  delays(1) = 1000
  delays(2) = 2000
ElseIf config("delay") = "MEDIUM" Then
  delays(0) = 1500
  delays(1) = 2000
  delays(2) = 4000
Else
  delays(0) = 2000
  delays(1) = 3000
  delays(2) = 6000
End If

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.Run """%PROGRAMFILES(x86)%\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"""

WScript.Sleep delays(0)

WshShell.AppActivate "Cisco AnyConnect Secure Mobility Client"

WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"

WScript.Sleep delays(1)
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"

WScript.Sleep delays(2)
WshShell.SendKeys config("username")
WshShell.SendKeys "{TAB}"
WScript.Sleep delays(0)
WshShell.SendKeys config("password")
WshShell.SendKeys "{ENTER}"
WScript.Sleep 500
