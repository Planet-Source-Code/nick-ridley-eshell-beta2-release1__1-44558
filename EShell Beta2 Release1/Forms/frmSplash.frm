VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrLoad 
      Interval        =   10
      Left            =   4140
      Top             =   2700
   End
   Begin MSComctlLib.ProgressBar prgLoad 
      Height          =   135
      Left            =   780
      TabIndex        =   0
      Top             =   4800
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lblLoad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "loading, please wait..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   780
      TabIndex        =   1
      Top             =   4560
      Width           =   1635
   End
   Begin VB.Shape shpBorder 
      Height          =   435
      Left            =   0
      Top             =   0
      Width           =   435
   End
   Begin VB.Image imgSplash 
      Height          =   5760
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Top             =   0
      Width           =   7650
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Byte
Dim d As Integer

Private Sub Form_Load()

If Function_Exist("user32", "SetLayeredWindowAttributes") = True Then SetLayered Me.hWND, True, t

'imgSplash.Picture = LoadPicture(App.path & "\gfx\splash2.gif")

Me.Width = imgSplash.Width
Me.Height = imgSplash.Height

shpBorder.Width = imgSplash.Width
shpBorder.Height = imgSplash.Height

Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2

'prgLoad.Width = Me.Width - 720
'prgLoad.Left = 360

'lblLoad.Top = prgLoad.Top - 240
'lblLoad.Left = Me.Width / 2 - lblLoad.Width / 2

End Sub

Private Sub tmrLoad_Timer()

If t < 255 And Function_Exist("user32", "SetLayeredWindowAttributes") = True Then

    SetLayered Me.hWND, True, t
    t = t + 5
    
Else

    prgLoad = prgLoad + 1
    
    If prgLoad = 100 Then prgLoad = 0: d = d + 1
    
    If d = 2 Then
            
        tmrLoad.Enabled = False
        Me.Hide
        frmConsole.Show
    
    End If

End If

End Sub
