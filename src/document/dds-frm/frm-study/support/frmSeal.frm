VERSION 5.00
Begin VB.Form frmSeal 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmSeal.frx":0000
   ScaleHeight     =   3645
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmSeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    
    'Set the main toolbar to not see the institution dropdown or report icon
    fMainForm.SetMainToolbar False
    
End Sub

Private Sub Form_Load()
    Hourglass True
End Sub
