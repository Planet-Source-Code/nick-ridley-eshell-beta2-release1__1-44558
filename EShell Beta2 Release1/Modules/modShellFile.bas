Attribute VB_Name = "modShellFile"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function ShellFile(path As String)
ShellFile = ShellExecute(frmConsole.hWND, "open", path, "", "", 1)
End Function

