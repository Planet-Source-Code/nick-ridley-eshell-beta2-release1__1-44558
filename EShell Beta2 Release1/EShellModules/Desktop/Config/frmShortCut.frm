VERSION 5.00
Begin VB.Form frmShortCut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Shortcut"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "frmShortCut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   2400
      X2              =   2400
      Y1              =   3600
      Y2              =   0
   End
   Begin VB.Image Image1 
      Height          =   3600
      Left            =   0
      Picture         =   "frmShortCut.frx":0E42
      Top             =   0
      Width           =   2400
   End
End
Attribute VB_Name = "frmShortCut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
