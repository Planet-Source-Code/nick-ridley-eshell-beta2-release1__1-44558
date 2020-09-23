Attribute VB_Name = "modMenu"

Public Function LoadMenu(Parent As Form, Folder As String, ttop As Long, tleft As Long, Optional IsRoot As Boolean = False) As Form

'On Error GoTo error

Dim x As New frmMenu
Dim c As Long
Dim i As Long, i2 As Long
Dim ff As Long, ff2 As Long
Dim data(1024) As String
Dim path As String
Dim icon As String
Dim t As Long
Dim p As Long

Load x

x.Folder = Folder
Set x.Parent = Parent
x.Root = IsRoot

x.Left = Screen.Width

'ff = FreeFile

'Open Folder & "\index.esm" For Input As #ff

'Do Until EOF(ff)

'    Line Input #ff, data(i)
    
'    If Left(data(i), 1) <> "#" Then i = i + 1
    
'Loop

frmMain.Dir1.path = Folder
frmMain.File1.path = Folder

Do Until i = frmMain.Dir1.ListCount

    data(i) = frmMain.Dir1.List(i)
    p = Len(data(i))
    
    Do Until Mid(data(i), p, 1) = "\"
        p = p - 1
    Loop
    
    data(i) = "\" & Right(data(i), Len(data(i)) - p)
    
    i = i + 1
    DoEvents
    
Loop

Do Until i2 = frmMain.File1.ListCount

    data(i) = frmMain.File1.List(i2)
    
    i2 = i2 + 1
    
    If Right(data(i), 4) = ".esl" Or Right(data(i), 4) = ".lnk" Then i = i + 1

Loop

For c = 0 To i - 1

    Load x.imgIcon(c + 1)
    Load x.lblItem(c + 1)
    Load x.lblFolder(c + 1)
    
noESL:
    
    If Left(data(c), 1) <> "\" Then
    
        If LCase(Right(data(c), 4)) <> ".esl" And LCase(Right(data(c), 4)) <> ".lnk" Then
        
            If c <> i - 1 Then
        
                c = c + 1
                GoTo noESL
    
            Else
            
                GoTo quickend
    
            End If
            
        Else
                
            x.lblFolder(c + 1).Visible = False
            
            If LCase(Right(data(c), 4)) = ".esl" Then
                
                ff2 = FreeFile
                
                Open Folder & "\" & data(c) For Input As #ff2
            
                Line Input #ff2, path
                Line Input #ff2, icon
                
                Close #ff2
                
                If UCase(Left(icon, 4)) <> "APP," Then
                
                    icon = Replace(LCase(icon), "%root%", frmMain.startroot)
                    
                    x.imgIcon(c + 1) = LoadPicture(icon)
                    
                Else
                              
                    icon = Right(icon, Len(icon) - InStr(1, icon, ","))
                              
                    DrawStartIcon path, frmMain.picIcon, True, CLng(icon)
                    SavePicture frmMain.picIcon.Image, App.path & "\temp.bmp"
                    DoEvents
                    x.imgIcon(c + 1) = LoadPicture(App.path & "\temp.bmp")
                    DoEvents
                    Kill App.path & "\temp.bmp"
                
                End If
                
            Else
            
                    DrawStartIcon Folder & "\" & data(c), frmMain.picIcon, True
                    SavePicture frmMain.picIcon.Image, App.path & "\temp.bmp"
                    DoEvents
                    x.imgIcon(c + 1) = LoadPicture(App.path & "\temp.bmp")
                    DoEvents
                    Kill App.path & "\temp.bmp"
                
            End If
                
            x.lblItem(c + 1) = Left(data(c), Len(data(c)) - 4)
            
            x.imgIcon(c + 1).Left = x.imgIcon(c).Left
            x.lblItem(c + 1).Left = x.lblItem(c).Left
            x.lblFolder(c + 1).Left = x.lblFolder(c).Left
            
            x.imgIcon(c + 1).Top = x.imgIcon(c).Top + 270
            Debug.Print x.imgIcon(c + 1).Top
            If x.imgIcon(c + 1).Top + 270 > Screen.Height Then
            
                x.imgIcon(c + 1).Top = 30
                x.imgIcon(c + 1).Left = x.Width + 30
                x.lblItem(c + 1).Left = x.Width + 360
                x.lblFolder(c + 1).Left = x.Width + 1980
                
                x.Width = x.Width + 2220
                
                Load x.Shape1(x.Shape1.UBound + 1)
                Load x.Shape2(x.Shape2.UBound + 1)
                
                x.Shape1(x.Shape1.UBound).Left = x.lblItem(c).Left + 1860
                x.Shape2(x.Shape2.UBound).Left = x.lblItem(c).Left + 1860
                x.Shape1(x.Shape1.UBound).ZOrder 0
                x.Shape2(x.Shape2.UBound).ZOrder 0
                x.Shape1(x.Shape1.UBound).Visible = True
                x.Shape2(x.Shape2.UBound).Visible = True
            
            End If
            
            x.lblItem(c + 1).Top = x.imgIcon(c + 1).Top
            x.lblFolder(c + 1).Top = x.imgIcon(c + 1).Top
        
            x.lblItem(c + 1).Visible = True
            x.imgIcon(c + 1).Visible = True
            
            x.lblItem(c + 1).Tag = Folder & "\" & data(c)
            
            x.lblItem(c + 1).ZOrder 0
            x.imgIcon(c + 1).ZOrder 0
            
        End If
        
        
    Else
        
        icon = frmMain.startroot & "\icon\programs.ico"
        
        x.imgIcon(c + 1) = LoadPicture(icon)
        x.lblItem(c + 1) = Right(data(c), Len(data(c)) - 1)
        
        x.imgIcon(c + 1).Left = x.imgIcon(c).Left
        x.lblItem(c + 1).Left = x.lblItem(c).Left
        x.lblFolder(c + 1).Left = x.lblFolder(c).Left
        
        x.imgIcon(c + 1).Top = x.imgIcon(c).Top + 270
        If x.imgIcon(c + 1).Top + 270 > Screen.Height Then
        
            x.imgIcon(c + 1).Top = 30
            x.imgIcon(c + 1).Left = x.Width + 30
            x.lblItem(c + 1).Left = x.Width + 360
            x.lblFolder(c + 1).Left = x.Width + 1980
            
            x.Width = x.Width + 2220
            
            Load x.Shape1(x.Shape1.UBound + 1)
            Load x.Shape2(x.Shape2.UBound + 1)
            
            x.Shape1(x.Shape1.UBound).Left = x.lblItem(c).Left + 1860
            x.Shape2(x.Shape2.UBound).Left = x.lblItem(c).Left + 1860
            x.Shape1(x.Shape1.UBound).ZOrder 0
            x.Shape2(x.Shape2.UBound).ZOrder 0
            x.Shape1(x.Shape1.UBound).Visible = True
            x.Shape2(x.Shape2.UBound).Visible = True
            
        End If
        
        x.lblItem(c + 1).Top = x.imgIcon(c + 1).Top
        x.lblFolder(c + 1).Top = x.imgIcon(c + 1).Top
    
        x.lblItem(c + 1).Visible = True
        x.imgIcon(c + 1).Visible = True
        x.lblFolder(c + 1).Visible = True
        
        x.lblItem(c + 1).Tag = data(c)
        
        x.lblItem(c + 1).ZOrder 0
        x.imgIcon(c + 1).ZOrder 0
        x.lblFolder(c + 1).ZOrder 0
    
    End If

    t = t + 1

Next c

quickend:

Close #ff

error:

If IsRoot Then
    
    t = t + 1
    
    Load x.imgIcon(t)
    Load x.lblItem(t)
    Load x.lblFolder(t)
    
    x.lblFolder(t).Visible = False
    x.lblItem(t).ZOrder 0
    x.imgIcon(t).ZOrder 0
    x.imgIcon(t).Visible = True
    x.lblItem(t).Visible = True
    
    x.imgIcon(t).Left = x.imgIcon(t - 1).Left
    x.lblItem(t).Left = x.lblItem(t - 1).Left
    x.lblFolder(t).Left = x.lblFolder(t - 1).Left
    
    x.imgIcon(t).Top = x.imgIcon(t - 1).Top + 270
    If x.imgIcon(t).Top + 270 > Screen.Height Then
    
        x.imgIcon(t).Top = 30
        x.imgIcon(t).Left = x.Width + 30
        x.lblItem(t).Left = x.Width + 360
        x.lblFolder(t).Left = x.Width + 1980
        
        x.Width = x.Width + 2220
        
        Load x.Shape1(x.Shape1.UBound + 1)
        Load x.Shape2(x.Shape2.UBound + 1)
        
        x.Shape1(x.Shape1.UBound).Left = x.lblItem(t - 1).Left + 1860
        x.Shape2(x.Shape2.UBound).Left = x.lblItem(t - 1).Left + 1860
        x.Shape1(x.Shape1.UBound).ZOrder 0
        x.Shape2(x.Shape2.UBound).ZOrder 0
        x.Shape1(x.Shape1.UBound).Visible = True
        x.Shape2(x.Shape2.UBound).Visible = True
        
    End If
    
    x.lblItem(t).Top = x.imgIcon(t).Top
    x.lblFolder(t).Top = x.imgIcon(t).Top
    
    x.lblItem(t) = "Shutdown"
    DrawStartIcon frmMain.startroot & "\icon\shutdown.ico", frmMain.picIcon, True, 0
    SavePicture frmMain.picIcon.Image, App.path & "\temp.bmp"
    DoEvents
    x.imgIcon(t) = LoadPicture(App.path & "\temp.bmp")
    DoEvents
    Kill App.path & "\temp.bmp"
    
    x.lblItem(t).Tag = "SHUTDOWN:"
    
    t = t + 1
    
    Load x.imgIcon(t)
    Load x.lblItem(t)
    Load x.lblFolder(t)
    
    x.lblFolder(t).Visible = False
    x.lblItem(t).ZOrder 0
    x.imgIcon(t).ZOrder 0
    x.imgIcon(t).Visible = True
    x.lblItem(t).Visible = True
    
    x.imgIcon(t).Left = x.imgIcon(t - 1).Left
    x.lblItem(t).Left = x.lblItem(t - 1).Left
    x.lblFolder(t).Left = x.lblFolder(t - 1).Left
    
    x.imgIcon(t).Top = x.imgIcon(t - 1).Top + 270
    If x.imgIcon(t).Top + 270 > Screen.Height Then
    
        x.imgIcon(t).Top = 30
        x.imgIcon(t).Left = x.Width + 30
        x.lblItem(t).Left = x.Width + 360
        x.lblFolder(t).Left = x.Width + 1980
        
        x.Width = x.Width + 2220
        
        Load x.Shape1(x.Shape1.UBound + 1)
        Load x.Shape2(x.Shape2.UBound + 1)
        
        x.Shape1(x.Shape1.UBound).Left = x.lblItem(t - 1).Left + 1860
        x.Shape2(x.Shape2.UBound).Left = x.lblItem(t - 1).Left + 1860
        x.Shape1(x.Shape1.UBound).ZOrder 0
        x.Shape2(x.Shape2.UBound).ZOrder 0
        x.Shape1(x.Shape1.UBound).Visible = True
        x.Shape2(x.Shape2.UBound).Visible = True
        
    End If
    
    x.lblItem(t).Top = x.imgIcon(t).Top
    x.lblFolder(t).Top = x.imgIcon(t).Top
    
    x.lblItem(t) = "Add Start Menu Folders"
    x.imgIcon(t).Picture = frmAddStart.Image1.Picture
    
    x.lblItem(t).Tag = "ADDSTART:"
    
End If

x.Height = t * 270 + 30

For i2 = 0 To x.Shape1.UBound

    x.Shape2(i2).Height = t * 270 + 30
    x.Shape1(i2).Height = t * 270 + 30

Next i2

x.Show

i2 = 0

Do Until i2 = i

    If x.lblItem(i2).Width + 580 > 2235 Then
        
        x.lblItem(i2) = x.lblItem(i2) & "..."
        
        Do Until x.lblItem(i2).Width + 580 < 2235
        
            x.lblItem(i2) = Left(x.lblItem(i2), Len(x.lblItem(i2)) - 4)
            x.lblItem(i2) = x.lblItem(i2) & "..."
        
        Loop
        
    End If

    i2 = i2 + 1

Loop

'If x.Width <> 2235 Then

'    i2 = 0

'    Do Until i2 = i
    
'        x.lblFolder(i2).Left = x.Width - 255
        'i2 = i2 + 1
'    Loop

'End If

If IsRoot = True Then

    If frmMain.Top - x.Height > 0 Then
        x.Top = frmMain.Top - x.Height + 15
    ElseIf frmMain.Top + x.Height < Screen.Height Then
        x.Top = frmMain.Top
    Else
        x.Top = 0
    End If
    
    If frmMain.Left - x.Width > 0 Then
        x.Left = frmMain.Left - x.Width + 15
    ElseIf frmMain.Left + x.Width < Screen.Width Then
        x.Left = frmMain.Left
    Else
        x.Left = 0
    End If
    
Else

    If Parent.Top + Parent.Li * 270 + 30 + x.Height < Screen.Height Then
        x.Top = ttop 'Parent.Top + Parent.Li * 270 - 270
    ElseIf Parent.Top + Parent.Li * 270 + 30 - x.Height > 0 Then
        x.Top = ttop - x.Height 'Parent.Top + Parent.Li * 270 - 270
    Else
        x.Top = 0
    End If
    
    If Parent.Left + Parent.Width + x.Width < Screen.Width Then
        x.Left = tleft 'Parent.Left + Parent.Width - 15
    ElseIf Parent.Left - x.Width > 0 Then
        x.Left = tleft - Parent.Width - x.Width + 15 'Parent.Left - x.Width + 15
    Else
        x.Left = 0
    End If
    
End If

Set LoadMenu = x

End Function
