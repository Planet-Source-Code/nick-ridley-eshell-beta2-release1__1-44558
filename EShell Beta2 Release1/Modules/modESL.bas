Attribute VB_Name = "modESL"

Public Function LoadESL(pth As String)

If Right(pth, 4) = ".lnk" Then ShellFile (pth): Exit Function

Dim ff As String
Dim file As String
Dim c As String, d As String, d2 As String

ff = FreeFile

Open pth For Input As #ff

Line Input #ff, file

Close #ff

If Left(file, 8) = "COMMAND," Then

    file = Right(file, Len(file) - 8)

    If Left(file, 5) = "CORE," Then
        
        file = Right(file, Len(file) - 5)
        c = Left(file, InStr(1, file, ",") - 1)
        d = Right(file, Len(file) - InStr(1, file, ","))
    
        d2 = Left(d, InStr(1, d, ",") - 1)
        d = Right(d, Len(d) - InStr(1, d, ","))
    
        Select Case c
            
            Case "WINDOW"
            
                Select Case d2
            
                Case "SHUTDOWN"
                frmShutdown.Show
            
                Case "CREATESHORT"
                frmCreateShortcut.SaveTo = d
                frmCreateShortcut.txtFile = ""
                frmCreateShortcut.txtName = ""
                frmCreateShortcut.Show
            
            End Select
        
        End Select
    
    Else
    
        
    
    End If

Else

    ShellFile file

End If

End Function
