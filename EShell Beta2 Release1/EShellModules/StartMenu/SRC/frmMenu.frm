VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2235
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   2235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTrans 
      Interval        =   5
      Left            =   1200
      Top             =   3120
   End
   Begin VB.Timer tmrClose 
      Interval        =   10
      Left            =   1740
      Top             =   3120
   End
   Begin VB.Label lblFolder 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1980
      TabIndex        =   1
      Top             =   -225
      Width           =   195
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Item]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   -210
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   0
      Left            =   30
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   240
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E7E7E7&
      BackStyle       =   1  'Opaque
      Height          =   3615
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00B2ACAD&
      BackStyle       =   1  'Opaque
      Height          =   3615
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   2235
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Parent As Form
Public ChildFrm As Form
Private GC As Boolean
Public Root As Boolean
Public Folder As String
Public Li As Long
Public MO As Boolean
Private t As Byte
Private rC As Boolean

Private Sub Form_Load()

Li = -1
If Function_Exist("user32", "SetLayeredWindowAttributes") = True Then SetLayered Me.hWnd, True, t

WindowPos Me, 1

End Sub

Private Sub lblItem_Click(index As Integer)

If lblFolder(index).Visible = True Then

    If GC Then ChildFrm.KillMenu
    Set ChildFrm = LoadMenu(Me, Folder & lblItem(index).Tag, Me.Top + Me.Li * 270 - 270, Me.Left + lblItem(index).Left + 1860)
        
    GC = True
    
Else

    If Not Root Then Parent.KillMenu
    If GC Then ChildFrm.KillMenu
    
    If lblItem(index).Tag = "ADDSTART:" Then
    
        frmAddStart.Show
        
    ElseIf lblItem(index).Tag = "SHUTDOWN:" Then
    
        frmMain.wsckModule.SendData "CORE,SHUTDOWN,"
    
    Else
    
        frmMain.wsckModule.SendData "CORE,LOADESL," & lblItem(index).Tag
    
    End If
    
    Unload Me
    
End If

End Sub

Private Sub lblItem_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Li = index Then Exit Sub

If Li <> -1 Then lblItem(Li).ForeColor = vbBlack: lblFolder(Li).ForeColor = vbBlack
Li = index
lblItem(Li).ForeColor = vbWhite
lblFolder(Li).ForeColor = vbWhite

End Sub

Private Sub tmrTrans_Timer()

If t < 255 And Function_Exist("user32", "SetLayeredWindowAttributes") = True Then

    SetLayered Me.hWnd, True, t
    t = t + 5

Else

    tmrTrans.Enabled = False

End If

End Sub

Private Sub tmrClose_Timer()

Dim x As Long, y As Long
Dim k As Boolean

x = GetX * 15: y = GetY * 15

If x < Me.Left Then k = True
If x > Me.Width + Me.Left Then k = True
If y < Me.Top Then k = True
If y > Me.Height + Me.Top Then k = True

If MO = False And k = False Then MO = True
If Not Root Then Parent.MO = True
If GC Then k = False

If k And MO And rC Then KillMenu (True)
If k And MO Then
    rC = True
Else
    rC = False
End If

End Sub

Public Function KillMenu(Optional force As Boolean = False)

On Error Resume Next

If GC Then ChildFrm.KillMenu

GC = False

If Not force And MO Then Exit Function

If Not Root Then Parent.KillMenu
Unload Me

End Function

Public Function SetChildFrm(frm As Form)

Set ChildFrm = frm

End Function
