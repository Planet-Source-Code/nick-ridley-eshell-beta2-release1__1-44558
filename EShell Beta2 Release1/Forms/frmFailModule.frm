VERSION 5.00
Begin VB.Form frmFailModule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Failed Loading Modules"
   ClientHeight    =   1425
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
   Icon            =   "frmFailModule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   375
      Left            =   2393
      TabIndex        =   2
      Top             =   840
      Width           =   795
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   375
      Left            =   1493
      TabIndex        =   1
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Some modules failed to load. Would you like to try loading these modules again?"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "frmFailModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdNo_Click()
Unload Me
End Sub

Private Sub cmdYes_Click()

Dim i As Integer

For i = 0 To modLoadModule.mCount

    If modLoadModule.failLoad(i) = True Then

        frmConsole.reloadModule i

    End If

Next i

Unload Me

End Sub

