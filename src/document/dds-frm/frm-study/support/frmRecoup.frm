VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecoup 
   Caption         =   "Validating Record for Recoupment"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   4545
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar proStatus 
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblPerforming 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmRecoup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
