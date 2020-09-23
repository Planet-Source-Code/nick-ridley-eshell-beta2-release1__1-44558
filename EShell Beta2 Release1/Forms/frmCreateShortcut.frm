VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCreateShortcut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Shortcut"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCreateShortcut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   60
      Top             =   1260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   255
      Left            =   4260
      TabIndex        =   5
      Top             =   300
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3420
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   900
      Width           =   3615
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   300
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   660
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Path:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   60
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frmCreateShortcut.frx":0E42
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmCreateShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SaveTo As String

Private Sub Command1_Click()

Dim ff As Long

ff = FreeFile

Open App.path & "\desktop\" & txtName & ".esl" For Output As #ff

    Print #ff, txtFile
    Print #ff, Image1.Tag
    
    frmConsole.wsckModule.RemotePort = FindPort("Desktop")
    frmConsole.wsckModule.RemoteHost = "127.0.0.1"
    DoEvents
    frmConsole.wsckModule.SendData "RELOADICONS,"

Close #ff

End Sub

Private Sub Command2_Click()

On Error Resume Next

CD1.Filter = "Applications|*.exe|All Files|*.*"
CD1.ShowOpen

If CD1.FileName = "" Then Exit Sub

If txtName = "" Then

    txtFile = CD1.FileName
    Dim p As Long
    p = Len(txtFile)
    
    Do Until Mid(txtFile, p, 1) = "\" Or p = 0
    
        p = p - 1
    
    Loop
    
    txtName = Right(txtFile, Len(textfile) - p)
    
    Load32Icon txtFile, 0, Image1
    iamge1.Tag = txtFile
    
End If

End Sub

Private Sub Image1_Click()

CD1.Filter = ""
CD1.ShowOpen

If CD1.FileName = "" Then Exit Sub

Image1.Tag = CD1.FileName
Load32Icon CD1.FileName, 0, Image1

End Sub

Private Sub txtFile_Change()

If Right(txtFile, 1) = "\" Then

    Load32Icon App.path & "\icon\folder.ico", 0, Image1
    Image1.Tag = App.path & "\icon\folder.ico"

End If

End Sub
