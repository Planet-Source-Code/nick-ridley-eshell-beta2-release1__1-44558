Attribute VB_Name = "modCfg"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Function ReadValue(Section As String, Key As String, Optional Default As String)
    Dim sReturn As String
    sReturn = String(255, Chr(0))
    ReadValue = Left(sReturn, GetPrivateProfileString(Section, Key, Default, sReturn, Len(sReturn), App.path & "\Settings.cfg"))
End Function

Public Sub SaveValue(Section As String, Key As String, Value As String)
    WritePrivateProfileString Section, Key, Value, App.path & "\Settings.cfg"
End Sub

