VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuDesk 
      Caption         =   "mnuDesk"
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuShort 
         Caption         =   "Create New Shortcut"
      End
   End
   Begin VB.Menu mnuIcon 
      Caption         =   "mnuIcon"
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cIcon As Long

Private Sub mnuDelete_Click()

Dim r As Long

r = MsgBox("Do you want to delete the shortcut file " & frmMain.lblCaption(cIcon).Tag & " ?", vbQuestion + vbYesNo, "Delete Icon")

If r = vbYes Then Kill frmMain.lblCaption(cIcon).Tag: LoadDesktop

End Sub

Private Sub mnuProperties_Click()

frmMain.wsckModule.SendData "DesktopAddOn,WINDOW,CONFIG"

End Sub

Private Sub mnuShort_Click()

frmMain.wsckModule.SendData "CORE,MAKEESL," & frmMain.ERoot & "\Desktop\"

End Sub
