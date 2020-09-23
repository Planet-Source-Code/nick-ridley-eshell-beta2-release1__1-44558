VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00B3ABAB&
   BorderStyle     =   0  'None
   Caption         =   "Tasklisting"
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDie 
      Interval        =   5000
      Left            =   120
      Top             =   1620
   End
   Begin VB.Timer tmrSize 
      Interval        =   10
      Left            =   600
      Top             =   1140
   End
   Begin VB.Timer tmrTaskUpdate 
      Interval        =   250
      Left            =   120
      Top             =   1140
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E7E7E7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   60
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox lstApps 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ListBox lstNames 
      Height          =   255
      Left            =   1380
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lstHwnd 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ListBox lstHwndNames 
      Height          =   255
      Left            =   1380
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock wsckModule 
      Left            =   2160
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   15
      LocalPort       =   16
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   0
      Top             =   0
      Width           =   2955
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E7E7E7&
      BackStyle       =   1  'Opaque
      Height          =   2175
      Left            =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   150
      Left            =   2640
      Picture         =   "frmMain.frx":0E42
      Top             =   1995
      Width           =   285
   End
   Begin VB.Label lblTask 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   60
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   60
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const mName = "Tasklist"
Public ERoot As String

Private H As Long
'Simple tasklisting example

Private Sub Form_Load()

If Command$ <> "-nocore" Then

    wsckModule.SendData "CORE,LOADED"

Else

    ERoot = "c:\vb\eshell beta2"
    tmrDie.Enabled = False
    
End If

Me.Icon = Nothing
WindowPos Me, 1
picIcon(0).Height = 16 * Screen.TwipsPerPixelY
picIcon(0).Width = 16 * Screen.TwipsPerPixelX
fEnumWindows Me.lstApps

picIcon(0).Top = picIcon(0).Top - picIcon(0).Height
lblTask(0).Top = lblTask(0).Top - lblTask(0).Height

Me.Left = ReadValue("Pos", "X", 120)
Me.Top = ReadValue("Pos", "Y", Screen.Height - 1200)

If Me.Top > Screen.Height Then Me.Top = Screen.Height - Me.Height
If Me.Left > Screen.Width Then Me.Left = Screen.Width - Me.Width

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0&

SaveValue "Pos", "X", Me.Left
SaveValue "Pos", "Y", Me.Top

End Sub

Private Sub lblTask_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then

    SetFGWindow lblTask(Index).Tag, True

ElseIf Button = 2 Then

    SetFGWindow lblTask(Index).Tag, False

End If

End Sub

Private Sub tmrDie_Timer()
End
End Sub

Private Sub tmrSize_Timer()

If Me.Height = H Then

    Exit Sub
    
Else

    If Me.Height < H Then
    
        Me.Height = Me.Height + 15
        
    Else
    
        Me.Height = Me.Height - 15
        
    End If

    Shape1.Height = Me.Height
    Shape2.Height = Me.Height
    Image2.Top = Me.Height - Image2.Height - 30

End If

End Sub

Private Sub tmrTaskUpdate_Timer()

ListApps

End Sub

Public Function ListApps()

On Error Resume Next

Dim i As Long, c As Long, i2 As Long
Dim d As Long
Dim e As Boolean

Me.lstApps.Clear
Me.lstNames.Clear

fEnumWindows Me.lstApps

DoEvents

i = lstApps.ListCount - 1
c = lstApps.ListCount

Do Until i < 0

d = 0
e = False

'check if window allready has an entry
Do Until d = lstHwnd.ListCount
If lstHwnd.List(d) = lstApps.List(i) Then e = True: Exit Do
d = d + 1
Loop

'Add it if its not there

If e = False Then

    Load lblTask(lblTask.UBound + 1)
    Load picIcon(picIcon.UBound + 1)
    
    lblTask(lblTask.UBound).Caption = lstNames.List(i)
    lblTask(lblTask.UBound).Top = lblTask(lblTask.UBound - 1).Top + picIcon(picIcon.UBound).Height + 30
    lblTask(lblTask.UBound).ZOrder 0
    lblTask(lblTask.UBound).Tag = lstApps.List(i)
    lblTask(lblTask.UBound).Visible = True
    lblTask(lblTask.UBound).AutoSize = True
    lblTask(lblTask.UBound).Refresh
            
    If lblTask(lblTask.UBound).Width > 2415 Or lblTask(lblTask.UBound).Height > 195 Then
        
        lblTask(lblTask.UBound) = lblTask(lblTask.UBound) & "..."
        
        Do Until lblTask(lblTask.UBound).Width <= 2415 And lblTask(lblTask.UBound).Height <= 195
        
            lblTask(lblTask.UBound) = Left(lblTask(lblTask.UBound), Len(lblTask(lblTask.UBound)) - 4)
            lblTask(lblTask.UBound) = lblTask(lblTask.UBound) & "..."
        
        Loop
        
    End If
    
    picIcon(picIcon.UBound).Top = picIcon(picIcon.UBound - 1).Top + picIcon(picIcon.UBound).Height + 30
    picIcon(picIcon.UBound).ZOrder 0
    picIcon(picIcon.UBound).AutoRedraw = True
    picIcon(picIcon.UBound).Visible = True
    
    lstHwnd.AddItem lstApps.List(i)
    lstHwndNames.AddItem lstNames.List(i)
    
    Call DrawIcon(picIcon(picIcon.UBound).hdc, lstApps.List(i), 0, 0)

End If

'Change the buttons text if the one on the form has changed

'If e = True Then

'End If

i = i - 1

Loop


i = 0
d = lstApps.ListCount

'Now check top see if windows that we have on the list still exits

Do Until i >= lstHwnd.ListCount

    c = 0
    e = False
    
    Do Until c = lstApps.ListCount
        
        If lstHwnd.List(i) = lstApps.List(c) Then
        
            lstNames.List(c) = lstHwndNames.List(i)
        
            For i2 = 0 To lstApps.ListCount
                
                If lstApps.List(i2) = lblTask(c).Tag Then
                
                If Right(lblTask(c), 3) = "..." Then
                    
                    If Left(lstNames.List(i2), Len(lblTask(c)) - 3) & "..." <> lblTask(c) Then
                    
                        lblTask(c).Caption = lstNames.List(i2)
                        
                    End If
                
                Else
                    
                    If lstNames.List(i2) <> lblTask(c) Then lblTask(c) = lstNames.List(i2)

                End If
                
                picIcon(c).Cls
                Call DrawIcon(picIcon(c).hdc, lstApps.List(i2), 0, 0)
                
                End If
                
            Next i2
            
            e = True: Exit Do
        
        End If
        
        c = c + 1
    
    Loop
    
    If e = False And c <> 0 Then
        
        'c = 0
        
        'Do Until lblTask(c).Caption = lstHwndNames.List(i)
            
        '    c = c + 1
        '    If c > lblTask.UBound Then GoTo kill
        
        'Loop
    
        RemTask c + 1
        DoEvents
        
        lstHwnd.RemoveItem i
        lstHwndNames.RemoveItem i
    
    End If
    
    For i2 = 0 To lstApps.ListCount
    
    If Right(lblTask(i2), 3) <> "..." Then
    
        If lblTask(i2).Width > 2340 And lblTask(i2).Height > 195 Then lblTask(i2) = lblTask(i2) & "..."
        
        Do 'Until
        
            If (lblTask(i2).Width <= 2340 And lblTask(i2).Height <= 195) Then Exit Do
        
            lblTask(i2) = Left(lblTask(i2), Len(lblTask(i2)) - 4)
            lblTask(i2) = lblTask(i2) & "..."
        
        Loop
    
    End If
    
    'Debug.Print lblTask(i2).Width

Next i2
    
kill:
    
    i = i + 1

Loop

H = lstHwnd.ListCount * 285 + 270

End Function

Public Function RemTask(i As Long)

Dim c As Long
c = i

Do Until c = lblTask.UBound
    
    lblTask(c).Caption = lblTask(c + 1).Caption
    lblTask(c).Tag = lblTask(c + 1).Tag
    picIcon(c).Picture = Nothing
    picIcon(c).Cls
    Call DrawIcon(picIcon(c).hdc, lblTask(c + 1).Tag, 0, 0)
    c = c + 1

Loop

Unload lblTask(lblTask.UBound)
Unload picIcon(picIcon.UBound)

End Function

Public Sub DrawIcon(hdc As Long, hWnd As Long, x As Integer, y As Integer)

ico = GetIcon(hWnd)
DrawIconEx hdc, x, y, ico, 16, 16, 0, 0, DI_NORMAL

End Sub

Public Function GetIcon(hWnd As Long) As Long

Call SendMessageTimeout(hWnd, WM_GETICON, 0, 0, 0, 1000, GetIcon)

If Not CBool(GetIcon) Then GetIcon = GetClassLong(hWnd, GCL_HICONSM)
If Not CBool(GetIcon) Then Call SendMessageTimeout(hWnd, WM_GETICON, 1, 0, 0, 1000, GetIcon)
If Not CBool(GetIcon) Then GetIcon = GetClassLong(hWnd, GCL_HICON)
If Not CBool(GetIcon) Then Call SendMessageTimeout(hWnd, WM_QUERYDRAGICON, 0, 0, 0, 1000, GetIcon)

End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim i As Long

Do Until i = lblTask.UBound + 1
    
    lblTask(i).FontUnderline = False
    i = i + 1

Loop

End Sub

Private Sub imgBG_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim i As Long

Do Until i = lblTask.UBound + 1

    lblTask(i).FontUnderline = False
    i = i + 1

Loop

End Sub

Private Sub lblTask_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim i As Long

Do Until i = lblTask.UBound + 1
    
    lblTask(i).FontUnderline = False
    i = i + 1

Loop

lblTask(Index).FontUnderline = True

End Sub

Private Sub wsckModule_DataArrival(ByVal bytesTotal As Long)

Dim data As String
Dim p As String, d As String

wsckModule.GetData data

DoEvents

Sleep 10

p = UCase(Left(data, InStr(1, data, ",") - 1))
d = Right(data, Len(data) - InStr(1, data, ","))

DoEvents

If p = "PORT" Then wsckModule.Close: wsckModule.Bind d: tmrDie.Enabled = False
If p = "ROOT" Then ERoot = d
If p = "KILL" Then End
If p = "COMMAND" Then

    Select Case d
    
        Case "REFRESH"
        ListApps
    
    End Select

End If

End Sub
