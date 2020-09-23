VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   450
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDie 
      Interval        =   5000
      Left            =   3120
      Top             =   0
   End
   Begin VB.Timer tmrSysTrayUpdate 
      Interval        =   500
      Left            =   1680
      Top             =   2280
   End
   Begin VB.PictureBox picSysTray 
      BackColor       =   &H00B3ABAB&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   60
      ScaleHeight     =   330
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   60
      Width           =   1875
      Begin VB.Image imgTrayIcon 
         Height          =   255
         Index           =   0
         Left            =   -255
         Stretch         =   -1  'True
         Top             =   -50
         Width           =   255
      End
   End
   Begin MSWinsockLib.Winsock wsckModule 
      Left            =   2640
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   15
      LocalPort       =   16
   End
   Begin VB.Image Image2 
      Height          =   150
      Left            =   2160
      Picture         =   "frmMain.frx":0E42
      Top             =   240
      Width           =   285
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00B3ABAB&
      BackStyle       =   1  'Opaque
      Height          =   450
      Left            =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const mName = "Systray"
Public ERoot As String

Private Sub Form_Load()

Call LoadTrayIconHandler
Call tmrSysTrayUpdate_Timer

If Command$ <> "-nocore" Then

    wsckModule.SendData "CORE,LOADED"

Else

    ERoot = "c:\vb\eshell beta2"
    tmrDie.Enabled = False
    
End If

Me.Left = ReadValue("Pos", "X", 120)
Me.Top = ReadValue("Pos", "Y", Screen.Height - 1200)

If Me.Top > Screen.Height Then Me.Top = Screen.Height - Me.Height
If Me.Left > Screen.Width Then Me.Left = Screen.Width - Me.Width

WindowPos Me, 1

End Sub

Private Sub Form_Unload(Cancel As Integer)

Call UnLoadTrayIconHandler

End Sub

Private Sub Image2_DblClick()
'Unload Me
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0&

SaveValue "Pos", "X", Me.Left
SaveValue "Pos", "Y", Me.Top

End Sub

Private Sub tmrDie_Timer()
Unload Me
End Sub

Private Sub tmrSysTrayUpdate_Timer()

Dim ET As Long
Dim dtrLeft As Long

For ET = 1 To imgTrayIcon.Count - 1
    If imgTrayIcon(ET).Tag <> "skip" Then dtrLeft = dtrLeft + 300
Next ET

picSysTray.Width = dtrLeft + 100

Me.Width = 465 + picSysTray.Width
Shape1.Width = Me.Width
Image2.Left = Me.Width - Image2.Width - 60

End Sub

Private Sub imgTrayicon_DblClick(index As Integer)
Dim msg As TrayIconMouseMessages
Dim ti As CTrayIcon
Dim lRet As Long
    msg = WM_LBUTTONDBLCLK
        On Error Resume Next
        Set ti = m_colTrayIcons(frmMain.imgTrayIcon(index).Tag)
        If Err.Number = 0 Then
            ti.PostCallbackMessage msg
        Else
            Err.Clear
        End If
        Set ti = Nothing
End Sub

Private Sub imgTrayicon_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim msg As TrayIconMouseMessages
Dim ti As CTrayIcon
Dim lRet As Long
        msg = WM_MOUSEMOVE
        On Error Resume Next
        Set ti = m_colTrayIcons(frmMain.imgTrayIcon(index).Tag)
        If Err.Number = 0 Then
            ti.PostCallbackMessage msg
        Else
            Err.Clear
        End If
        Set ti = Nothing
End Sub

Private Sub imgTrayicon_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim msg As TrayIconMouseMessages
Dim ti As CTrayIcon
Dim lRet As Long
 
     If Button = 1 Then
        msg = WM_LBUTTONDOWN
     ElseIf Button = 2 Then
        msg = WM_RBUTTONDOWN
     End If
       
    On Error Resume Next
    Set ti = m_colTrayIcons(frmMain.imgTrayIcon(index).Tag)
    If Err.Number = 0 Then
        ti.PostCallbackMessage msg
    Else
        Err.Clear
    End If
    Set ti = Nothing
End Sub

Private Sub imgTrayicon_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim msg As TrayIconMouseMessages
Dim ti As CTrayIcon
Dim lRet As Long
     If Button = 1 Then
        msg = WM_LBUTTONUP
     ElseIf Button = 2 Then
        msg = WM_RBUTTONUP
     End If
       
    On Error Resume Next
    Set ti = m_colTrayIcons(frmMain.imgTrayIcon(index).Tag)
    If Err.Number = 0 Then
        ti.PostCallbackMessage msg
    Else
        Err.Clear
    End If
    Set ti = Nothing
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
If p = "KILL" Then Unload Me
If p = "COMMAND" Then

    Select Case d
    
    
    End Select

End If

End Sub


