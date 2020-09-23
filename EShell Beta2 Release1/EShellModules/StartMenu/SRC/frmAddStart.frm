VERSION 5.00
Begin VB.Form frmAddStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add start menu items"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Done"
      Height          =   315
      Left            =   5160
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   315
      Left            =   5160
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   315
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtFolder 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   4995
   End
   Begin VB.Label Label2 
      Caption         =   "i.e. c:\windows\start menu\programs\"
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   960
      Width           =   4995
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frmAddStart.frx":0E42
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Please select a folder below to copy its contents to you EShell start menu"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5835
   End
End
Attribute VB_Name = "frmAddStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SH As New Shell  'reference to shell32.dll class
Dim ShBFF As Folder  'Shell Browse For Folder

Private Sub Command1_Click()

On Error Resume Next

Set ShBFF = SH.BrowseForFolder(hWnd, "Please select the folder you whish to copy to your start menu", 1)
            
With ShBFF.Items.Item
   
    txtFolder = .path
   
End With

End Sub

Private Sub Command2_Click()

    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.createfolder("c:\windowscopy")
    ' For Example:
    path1$ = txtFolder
    path2$ = frmMain.startroot & "\startmenu\"


    If fso.folderexists(path1$) Then


        If Not fso.folderexists("c:\windowscopy") Then
            'Generate Path
            Set fld = fso.createfolder("c:\windowscopy")
        End If
        'Copy now
        fso.copyfolder path1$, path2$, True
        'On Error:
    Else
        MsgBox "That folder does not exist"
    End If

    Set fso = Nothing

End Sub

Private Sub Command3_Click()
Me.Hide
End Sub
