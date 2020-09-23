VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desktop Properties"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "frmConfig"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkBGImage 
      Caption         =   "Background Image"
      Height          =   255
      Left            =   210
      TabIndex        =   8
      Top             =   3780
      Width           =   1695
   End
   Begin VB.TextBox txtBGImage 
      Height          =   255
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4080
      Width           =   2115
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   4080
      Width           =   375
   End
   Begin VB.ComboBox cmbStretch 
      Height          =   330
      ItemData        =   "frmConfig.frx":0E42
      Left            =   3240
      List            =   "frmConfig.frx":0E4C
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4380
      Width           =   1215
   End
   Begin VB.PictureBox picCol 
      Height          =   315
      Left            =   3300
      ScaleHeight     =   255
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdColour 
      Caption         =   "..."
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   5340
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3420
      TabIndex        =   1
      Top             =   5340
      Width           =   1155
   End
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1200
      ScaleHeight     =   1695
      ScaleWidth      =   2280
      TabIndex        =   0
      Top             =   540
      Width           =   2280
      Begin VB.Image imgBG 
         Height          =   615
         Left            =   720
         Stretch         =   -1  'True
         Top             =   540
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   5220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Use these settings to configure your desktop:"
      Height          =   255
      Left            =   180
      TabIndex        =   11
      Top             =   3060
      Width           =   4275
   End
   Begin VB.Label Label1 
      Caption         =   "Background Image:"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Background Colour:"
      Height          =   195
      Left            =   480
      TabIndex        =   9
      Top             =   4860
      Width           =   1515
   End
   Begin VB.Image imgMonitor 
      Height          =   2535
      Left            =   960
      Picture         =   "frmConfig.frx":0E61
      Top             =   300
      Width           =   2760
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkBGImage_Click()

If chkBGImage.Value = 1 Then
    
    txtBGImage.Locked = False
    txtBGImage = ReadValue("desktop", "bg", frmMain.ERoot & "\eshell.cfg", "")
    cmdBrowse.Enabled = True
    cmbStretch.Locked = False
    
Else

    txtBGImage.Locked = True
    txtBGImage = ""
    cmdBrowse.Enabled = False
    cmbStretch.Locked = True

End If

CalcDesktop

End Sub

Private Sub cmbStretch_Click()
CalcDesktop
End Sub

Private Sub cmdBrowse_Click()

CD.Filter = "Image Files|*.bmp;*.jpg;*.gif|All Files|*.*"
CD.ShowOpen
txtBGImage = CD.FileName

CalcDesktop

End Sub

Private Sub cmdCancel_Click()

Form_Load
Me.Hide

End Sub

Private Sub cmdColour_Click()

CD.Color = picBG.BackColor
CD.ShowColor
picCol.BackColor = CD.Color

CalcDesktop

End Sub

Private Sub cmdOK_Click()

If cmbStretch = "Stretch" Then

    SaveValue "desktop", "stretch", "True", frmMain.ERoot & "\eshell.cfg"

Else

    SaveValue "desktop", "stretch", "False", frmMain.ERoot & "\eshell.cfg"

End If

SaveValue "desktop", "bg", txtBGImage, frmMain.ERoot & "\eshell.cfg"
SaveValue "desktop", "bgcol", picCol.BackColor, frmMain.ERoot & "\eshell.cfg"

Form_Load
Me.Hide

frmMain.wsckModule.SendData "Desktop,COMMAND,REFRESH"

End Sub

Private Sub Form_Load()

On Error Resume Next

If ReadValue("desktop", "stretch", frmMain.ERoot & "\eshell.cfg", False) = False Then

cmbStretch = "Centre"

Else

cmbStretch = "Stretch"

End If

txtBGImage = ReadValue("desktop", "bg", frmMain.ERoot & "\eshell.cfg", "")

If ReadValue("desktop", "bg", frmMain.ERoot & "\eshell.cfg", "") <> "" Then

    chkBGImage.Value = 1
    txtBGImage.Locked = False
    cmdBrowse.Enabled = True
    cmbStretch.Locked = False
    
End If

picCol.BackColor = ReadValue("desktop", "bgcol", frmMain.ERoot & "\eshell.cfg", "")

CalcDesktop

End Sub

Public Function CalcDesktop()

On Error Resume Next

picBG.BackColor = picCol.BackColor

If txtBGImage <> "" Then
    
    imgBG.Visible = True
    imgBG.Stretch = False
    imgBG.Picture = LoadPicture(txtBGImage)
    imgBG.Stretch = True

    
    If cmbStretch = "Centre" Then

        imgBG.Width = (imgBG.Width / Screen.Width) * picBG.Width
        imgBG.Height = (imgBG.Height / Screen.Height) * picBG.Height

    Else
        
        imgBG.Width = picBG.Width
        imgBG.Height = picBG.Height
        imgBG.Top = 0
        imgBG.Left = 0

    End If

    imgBG.Top = picBG.Height / 2 - imgBG.Height / 2
    imgBG.Left = picBG.Width / 2 - imgBG.Width / 2
    
Else

    imgBG.Visible = False
    
End If

End Function
