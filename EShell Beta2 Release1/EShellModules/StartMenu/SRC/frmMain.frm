VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00B2ACAD&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   720
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   720
   ScaleWidth      =   720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDie 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin VB.FileListBox File1 
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.DirListBox Dir1 
      Height          =   330
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E7E7E7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSWinsockLib.Winsock wsckModule 
      Left            =   4200
      Top             =   2760
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
      Left            =   405
      Picture         =   "frmMain.frx":1042
      Top             =   540
      Width           =   285
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmMain.frx":13B3
      Top             =   0
      Width           =   720
   End
   Begin VB.Shape Shape1 
      Height          =   720
      Left            =   0
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const mName = "Start Menu"
Private Root As Form
Public startroot As String

Private Sub Form_Load()

If Command$ <> "-nocore" Then

    'send confirm load to core
    wsckModule.SendData "CORE,LOADED"

Else

    startroot = "C:\VB\EShell Beta2"
    tmrDie.Enabled = False

End If

Me.Left = ReadValue("Pos", "X", 120)
Me.Top = ReadValue("Pos", "Y", Screen.Height - 1200)

If Me.Top > Screen.Height Then Me.Top = Screen.Height - Me.Height
If Me.Left > Screen.Width Then Me.Left = Screen.Width - Me.Width

WindowPos Me, 1

End Sub

Private Sub Image1_Click()
Set Root = LoadMenu(Me, startroot & "\startmenu", 0, 0, True)
End Sub

Private Sub Image1_DblClick()
End
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0&

SaveValue "Pos", "X", Me.Left
SaveValue "Pos", "Y", Me.Top

End Sub

Private Sub tmrDie_Timer()
End
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
If p = "ROOT" Then startroot = d
If p = "KILL" Then End

End Sub
