VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   5700
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   9360
      Begin VB.PictureBox picLogo 
         Height          =   5205
         Left            =   150
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   5145
         ScaleWidth      =   5235
         TabIndex        =   1
         Top             =   255
         Width           =   5295
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Direct Deposit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5955
         TabIndex        =   7
         Tag             =   "Product"
         Top             =   2265
         Width           =   2580
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "HEARTS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   5670
         TabIndex        =   6
         Tag             =   "CompanyProduct"
         Top             =   1740
         Width           =   1965
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5985
         TabIndex        =   5
         Tag             =   "Version"
         Top             =   2865
         Width           =   2070
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning"
         Height          =   195
         Left            =   300
         TabIndex        =   2
         Tag             =   "Warning"
         Top             =   3720
         Width           =   6855
      End
      Begin VB.Label lblCompany 
         Height          =   255
         Left            =   5625
         TabIndex        =   4
         Tag             =   "Company"
         Top             =   5310
         Width           =   2415
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright 1999, 2000"
         Height          =   255
         Left            =   5610
         TabIndex        =   3
         Tag             =   "Copyright"
         Top             =   5100
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    frmSplash.MousePointer = vbHourglass
    DoEvents
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
        
End Sub

