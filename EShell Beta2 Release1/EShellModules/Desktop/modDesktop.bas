Attribute VB_Name = "modDesktop"
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_BOTTOM = 1
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

Private locx(255) As Long
Private locy(255) As Long
Private iname(255) As String
Private icount As Long

Public Sub SetDesktop(Whwnd As Long, WindowHwnd As Form)

    SetWindowPos Whwnd, HWND_BOTTOM, WindowHwnd.Top / Screen.TwipsPerPixelX, WindowHwnd.Left / Screen.TwipsPerPixelY, WindowHwnd.Width / Screen.TwipsPerPixelX, WindowHwnd.Height / Screen.TwipsPerPixelY, 0

End Sub

Public Function LoadDesktop()

On Error Resume Next

Dim i As Long
Dim pth As String
Dim ff As Long
Dim l As String
Dim p As Long
Dim path As String, icon As String
Dim x As Long, y As Long

If frmMain.imgIcon.UBound > 0 Then
    
    For i = 1 To frmMain.imgIcon.UBound
    
        Unload frmMain.imgIcon(i)
        Unload frmMain.lblCaption(i)
    
    Next i

End If

pth = frmMain.ERoot & "\Desktop\"

ff = FreeFile

Open pth & "desktop.cfg" For Input As #ff

Do Until EOF(ff)

    Line Input #ff, l
    p = InStr(1, l, ",")
    iname(i) = Left(l, p - 1)
    l = Right(l, Len(l) - p)
    p = InStr(1, l, ",")
    locx(i) = Left(l, p - 1)
    locy(i) = Right(l, Len(l) - p)
    
    i = i + 1

Loop

icount = i - 1

Close ff

frmMain.File1.path = "C:\"
frmMain.File1.path = pth

i = 0

For i = 0 To frmMain.File1.ListCount - 1

    If Right(frmMain.File1.List(i), 4) = ".lnk" Or Right(frmMain.File1.List(i), 4) = ".esl" Then
    
        Load frmMain.imgIcon(frmMain.imgIcon.UBound + 1)
        Load frmMain.lblCaption(frmMain.lblCaption.UBound + 1)
                
        With frmMain.imgIcon(frmMain.imgIcon.UBound)
        
            .Top = frmMain.imgIcon(frmMain.imgIcon.UBound - 1).Top + .Height + 210 + 180
            .ZOrder 0
            
        End With
        With frmMain.lblCaption(frmMain.imgIcon.UBound)
        
            .Top = frmMain.lblCaption(frmMain.imgIcon.UBound - 1).Top + frmMain.imgIcon(frmMain.imgIcon.UBound - 1).Height + 210 + 180
            .ZOrder 0
            
        End With
        
        If GetXYIcon(pth & frmMain.File1.List(i), x, y) = True Then
        
            frmMain.imgIcon(frmMain.imgIcon.UBound - 1).Left = x
            frmMain.imgIcon(frmMain.imgIcon.UBound - 1).Top = y
            frmMain.lblCaption(frmMain.imgIcon.UBound - 1).Top = y + frmMain.imgIcon(frmMain.imgIcon.UBound - 1).Height
        
        End If
        
        frmMain.imgIcon(frmMain.imgIcon.UBound - 1).Visible = True
        frmMain.lblCaption(frmMain.imgIcon.UBound - 1).Visible = True
        
        If Right(frmMain.File1.List(i), 4) = ".lnk" Then
            
            Load32Icon pth & frmMain.File1.List(i), 0, frmMain.imgIcon(frmMain.imgIcon.UBound - 1)
        
        Else
        
                ff = FreeFile
                
                Open pth & frmMain.File1.List(i) For Input As #ff
            
                Line Input #ff, path
                Line Input #ff, icon
                
                Close #ff
                
                If UCase(Left(icon, 4)) <> "APP," Then
                
                    icon = Replace(LCase(icon), "%root%", frmMain.ERoot)
                    
                    Load32Icon icon, 0, frmMain.imgIcon(frmMain.imgIcon.UBound - 1)
                    
                Else
                              
                    icon = Right(icon, Len(icon) - InStr(1, icon, ","))
                              
                    Load32Icon path, CLng(icon), frmMain.imgIcon(frmMain.imgIcon.UBound - 1)
                
                
                End If
                
        End If
        
        With frmMain.lblCaption(frmMain.lblCaption.UBound - 1)
        
        .Caption = Left(frmMain.File1.List(i), Len(frmMain.File1.List(i)) - 4)
        .Tag = pth & frmMain.File1.List(i)
        
rewidth:
        
        If .Width > 960 Then
                
            If Right(.Caption, 3) <> "..." Then .Caption = .Caption & "..."
            .Caption = Left(.Caption, Len(.Caption) - 4)
            .Caption = .Caption & "..."
            
            GoTo rewidth
            
        End If
        
        .Left = frmMain.imgIcon(frmMain.imgIcon.UBound - 1).Left + (frmMain.imgIcon(frmMain.imgIcon.UBound - 1).Width / 2) - (.Width / 2)
        
        End With
        
    End If

Next i

End Function

Public Function GetXYIcon(ina As String, x As Long, y As Long) As Boolean

Dim i As Long

For i = 0 To 255

    If iname(i) = ina Then
    
        x = locx(i)
        y = locy(i)
        GetXYIcon = True
        Exit Function
    
    End If

Next i

End Function

Public Function ChangeXYIcon(ina As String, x As Long, y As Long)

Dim i As Long

For i = 0 To 255

    If iname(i) = ina Then
    
        locx(i) = x
        locy(i) = y
        ReWriteXYFile
        Exit Function
    
    End If

Next i

icount = icount + 1
iname(icount) = ina
locx(icount) = x
locy(icount) = y

ReWriteXYFile

End Function

Public Function ReWriteXYFile()

Dim ff As Long
Dim i As Long

ff = FreeFile

Open frmMain.ERoot & "\Desktop\desktop.cfg" For Output As #ff

For i = 0 To icount

    Print #ff, iname(i) & "," & locx(i) & "," & locy(i)

Next i

Close #ff

End Function
