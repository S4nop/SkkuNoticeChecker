Attribute VB_Name = "Start_Program"
Public Function Startup()
Set Startup = CreateObject("WScript.Shell")
Startup.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\SKKU", App.Path & "\" & App.EXEName & ".exe"
End Function

Public Function StartUPDelete()
On Error Resume Next

    Dim stud
    Set stud = CreateObject("WScript.Shell")
    stud.RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\SKKU", App.Path & "\" & App.EXEName & ".exe"
End Function
