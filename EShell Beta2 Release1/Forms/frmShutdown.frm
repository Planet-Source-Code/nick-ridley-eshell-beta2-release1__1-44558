VERSION 5.00
Begin VB.Form frmShutdown 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shutdown"
   ClientHeight    =   2220
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
   Icon            =   "frmShutdown.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   3420
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   435
      Left            =   2340
      TabIndex        =   0
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      ItemData        =   "frmShutdown.frx":1042
      Left            =   533
      List            =   "frmShutdown.frx":104C
      TabIndex        =   3
      Text            =   "Powerdown"
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "What do you want the computer to do?"
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmShutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Select Case Combo1

    Case "Powerdown"
    'Shutdown SM_Shutdown
    MsgBox "powerdown"
    
    Case "Restart"
    'Shutdown SM_Reboot
    MsgBox "powerdown"
    
End Select

End Sub

Private Sub Command2_Click()

Me.Hide

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()

Me.Show
DoEvents
Load32Icon App.path & "\icon\shutdown.ico", 0, Image1

End Sub


