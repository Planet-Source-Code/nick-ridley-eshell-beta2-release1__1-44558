VERSION 5.00
Begin VB.Form frmRelease 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EShell Beta2 Release1 - Details"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "NO, I have not compiled all the exe's and wish to exit and then do so"
      Height          =   555
      Left            =   3634
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK, I have read all instructions and compiled all exe's"
      Height          =   555
      Left            =   131
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   5055
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   $"frmRelease.frx":0000
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   6855
   End
End
Attribute VB_Name = "frmRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
frmSplash.Show
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()

Dim ff As Long
Dim l As String

ff = FreeFile

Open App.path & "\readme.txt" For Input As #ff

Do Until EOF(ff)

    Line Input #ff, l
    
    Text1 = Text1 & l & vbCrLf

Loop

Close ff

MsgBox "When compiling EXE's your instalation directory ($) is '" & App.path & "'.", vbOKOnly + vbInformation, "Please Note:"

End Sub

