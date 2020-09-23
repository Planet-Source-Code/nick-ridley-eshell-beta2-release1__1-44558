VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDie 
      Interval        =   5000
      Left            =   2760
      Top             =   420
   End
   Begin MSWinsockLib.Winsock wsckModule 
      Left            =   3240
      Top             =   420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   15
      LocalPort       =   16
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const mName = "DesktopAddOn"
Public ERoot As String

Private Sub Form_Load()

On Error Resume Next

If Command$ <> "-nocore" Then

    'wsckModule.RemotePort = 15
    'wsckModule.RemoteHost = "127.0.0.1"
    wsckModule.SendData "CORE,LOADED"
    DoEvents
    
Else

    ERoot = "c:\vb\eshell beta2"
    tmrDie.Enabled = False
    
End If

Me.Visible = False

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
If p = "ROOT" Then ERoot = d
If p = "KILL" Then End
If p = "WINDOW" Then

    Select Case d
    
        Case "CONFIG"
        frmConfig.Show

    End Select

End If

End Sub
