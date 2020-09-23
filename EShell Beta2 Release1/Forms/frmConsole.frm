VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmConsole 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EShell - Console"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   6
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConsole.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00B3ABAB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4680
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSWinsockLib.Winsock wsckModule 
      Left            =   6360
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
   End
   Begin VB.TextBox txtConsole 
      Height          =   3675
      Left            =   53
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   780
      Width           =   7095
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   6840
      Picture         =   "frmConsole.frx":0E42
      ToolTipText     =   "add new start menu items from windows into EShell"
      Top             =   420
      Width           =   240
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7200
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmConsole.frx":11D0
      Top             =   0
      Width           =   7200
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim l As Boolean
Dim c As Long

Dim mL As Long

Private Sub Form_Load()

Dim i As Integer
Dim m As Boolean
Dim t As Integer
Dim tick As Long, lo As Long
Dim ttick As Long

Me.Show
DoEvents

ttick = GetTickCount

AddLine "EShell beta 2, loading, please wait"
AddLine "Current version is " & App.Major & "." & App.Minor & "." & App.Revision
AddLine "Many icons used in this program(s) are from the metrox icon set by Greg Fleming (aka Hell Dragon) of darkproject.com"
AddLine ""
AddLine "Loading Module Comm's Socket..."

wsckModule.Protocol = sckUDPProtocol
wsckModule.Bind 15
wsckModule.RemoteHost = "127.0.0.1"

AddLine ""
AddLine "Loading Modules:"
AddLine "Loading Core..."
AddLine "Reading Modules Config File..."
LoadModuleConfig
AddLine "Loading " & modLoadModule.mCount & " modules from module.cfg"

For i = 0 To modLoadModule.mCount - 1

    tick = GetTickCount

    AddLine "Loading " & modLoadModule.mName(i) & "..."
    c = i
    m = LoadModule(i)
    DoEvents
    
    Do
    
        If l Then l = False: Exit Do
        If tick + 1000 <= GetTickCount Then GoTo fail
        DoEvents
        DoEvents
        DoEvents
        
    Loop
    
    If m Then
    
        wsckModule.RemotePort = FindPort(modLoadModule.mName(i))
        DoEvents
        wsckModule.SendData "ROOT," & App.path
        DoEvents
        Sleep 25
        DoEvents
        AddLine "Done..."
        t = t + 1
        
    Else
    
fail:
        AddLine "Failed loading module " & modLoadModule.mName(i) & "!..."
        modLoadModule.failLoad(i) = True

    End If

    Sleep 25

Next i

AddLine ""
AddLine "Sucessfully loaded " & t & " of " & modLoadModule.mCount & " modules"

If t <> modLoadModule.mCount Then frmFailModule.Show

AddLine ""
AddLine "Loaded EShell in " & (GetTickCount - ttick) & "ms"

mL = t

End Sub

Public Function AddLine(txt As String)

txtConsole.Text = txtConsole.Text & txt & vbCrLf
txtConsole.SelStart = Len(txtConsole.Text)

End Function

Private Sub Form_Unload(Cancel As Integer)

Cancel = 1

Dim i As Long

For i = 0 To modLoadModule.mCount - 1

    wsckModule.RemoteHost = "127.0.0.1"
    wsckModule.RemotePort = 1200 + i
    DoEvents
    wsckModule.SendData "KILL,"
    DoEvents

Next i
Sleep 10
DoEvents
End

End Sub

Private Sub Image2_Click()
wsckModule.RemotePort = "1203"
wsckModule.RemoteHost = "127.0.0.1"
wsckModule.SendData "ROOT," & App.path
End Sub

Private Sub wsckModule_DataArrival(ByVal bytesTotal As Long)

On Error Resume Next

Dim data As String

wsckModule.GetData data

If Left(data, 5) = "CORE," Then

    data = Right(data, Len(data) - 5)
    
    If data = "LOADED" Then wsckModule.RemotePort = 16: DoEvents: l = True: wsckModule.SendData "Port," & 1200 + c: AddLine ("PORT: " & 1200 + c)
    If Left(data, 8) = "LOADESL," Then LoadESL (Right(data, Len(data) - 8))
    If Left(data, 8) = "MAKEESL," Then frmCreateShortcut.Show: frmCreateShortcut.SaveTo = Right(data, Len(data) - 8)
    If Left(data, 9) = "SHUTDOWN," Then frmShutdown.Show
    
Else

    Dim n As String
    Dim i As Long
    
    n = Left(data, InStr(1, data, ",") - 1)
    data = Right(data, Len(data) - InStr(1, data, ","))

    i = FindPort(n)
    If i <> 0 Then GoTo foundport

    Exit Sub

foundport:

    wsckModule.RemotePort = i
    wsckModule.RemoteHost = "127.0.0.1"
    DoEvents
    wsckModule.SendData data

End If

End Sub

Public Sub reloadModule(i As Integer)

Dim tick As Long
Dim m As Boolean

tick = GetTickCount

AddLine "Reloading " & modLoadModule.mName(i) & "..."
c = i
m = LoadModule(i)
DoEvents

Do

    If l Then l = False: Exit Do
    If tick + 1000 <= GetTickCount Then GoTo fail
    DoEvents
    DoEvents
    DoEvents
    
Loop

If m Then

    wsckModule.RemotePort = FindPort(modLoadModule.mName(i))
    DoEvents
    wsckModule.SendData "ROOT," & App.path
    AddLine "Done..."
    mL = mL + 1
    
Else

fail:
    AddLine "Failed loading module " & modLoadModule.mName(i) & "!..."
    failLoad(i) = True

End If

End Sub
