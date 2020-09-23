Attribute VB_Name = "modFuncExist"
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Public Function Function_Exist(ByVal sModule As String, ByVal sFunction As String) As Boolean

    Dim hHandle As Long
    
    hHandle = GetModuleHandle(sModule)
    If hHandle = 0 Then
        
        hHandle = LoadLibraryEx(sModule, 0&, 0&)
        
        If GetProcAddress(hHandle, sFunction) = 0 Then
            Function_Exist = False
        Else
            Function_Exist = True
        End If
        
        FreeLibrary hHandle
    Else
        If GetProcAddress(hHandle, sFunction) = 0 Then
            Function_Exist = Function_Exist
        Else
            Function_Exist = True
        End If
    End If
    
End Function
