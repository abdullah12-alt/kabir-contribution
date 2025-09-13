VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{8CD222DF-7752-11D3-9D1E-00105A19BCF2}#1.0#0"; "OAOTBAR.OCX"
Begin VB.Form frmEditDeposits 
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9975
   ControlBox      =   0   'False
   Icon            =   "frmEditDeposits.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   9975
   WindowState     =   2  'Maximized
   Begin OAOTitleBar.OutlookTitleBar OutlookTitle1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   979
      ForeColor       =   16777215
      Caption         =   "Validate Transactions"
   End
   Begin MSComctlLib.ProgressBar proStatus 
      Height          =   390
      Left            =   4230
      TabIndex        =   5
      Top             =   4290
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdViewTrans 
      Caption         =   "Valid &Transactions"
      Height          =   465
      Left            =   3765
      TabIndex        =   4
      Top             =   6375
      Width           =   1815
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   1110
      Top             =   6885
      _ExtentX        =   10266
      _ExtentY        =   5345
      _Version        =   393216
      Rows            =   2
      Cols            =   3
      Picture         =   "frmEditDeposits.frx":000C
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   3885
      TabIndex        =   2
      Top             =   1740
      Width           =   5715
      Begin VB.Label lblPerforming 
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   4455
      End
      Begin VB.Label lblEditStatus 
         Caption         =   "Press the validate button below to start the validation process."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   2115
         TabIndex        =   3
         Top             =   570
         Width           =   3450
      End
      Begin VB.Image Image1 
         Height          =   1680
         Left            =   105
         Top             =   435
         Width           =   1815
      End
   End
   Begin VB.Timer AnimationTimer 
      Interval        =   10
      Left            =   3660
      Top             =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   7695
      TabIndex        =   1
      Top             =   6375
      Width           =   1815
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Validate"
      Height          =   465
      Left            =   5760
      TabIndex        =   0
      Top             =   6375
      Width           =   1815
   End
   Begin VB.Image imgImage 
      BorderStyle     =   1  'Fixed Single
      Height          =   6375
      Left            =   75
      Picture         =   "frmEditDeposits.frx":13686
      Stretch         =   -1  'True
      Top             =   765
      Width           =   3510
   End
End
Attribute VB_Name = "frmEditDeposits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Toggle As Integer
Public iPic As Integer

Private Const MODULE As String = "Validate Transactions - "

Public Sub ShowStatus(iPercent As Integer)
    
    proStatus.Value = iPercent
    DoEvents
    
End Sub

Private Sub AnimationTimer_Timer()

Static iPercent As Integer
    If Toggle = 1 Then
        iPic = iPic + 1
        If iPic = 6 Then
            iPic = 0
        End If
        Image1.Picture = PictureClip1.GraphicCell(iPic)
        DoEvents
    End If
     
    
End Sub



Private Sub cmdCancel_Click()
Hourglass True
Toggle = 0
Unload Me
Hourglass False

End Sub

Public Sub cmdEdit_Click()
    
    giProcess = VALIDATE_ONLY
    ValidateTransactionsNew

End Sub

Private Sub cmdViewTrans_Click()
    Hourglass True
    Load frmValidated
    'frmValidated.adcValid.ConnectionString = gobjLoginInfo.ConnectString
    'frmValidated.adcValid.RecordSource = "SELECT * FROM DD_VALID_REC WHERE CREATED_DATETIME >= '" & Format$(Now, "MM/dd/yyyy") & " 00:00:00'"
    'frmValidated.adcValid.Refresh
    frmValidated.Show
    Hourglass False
End Sub



Private Sub Form_Load()
Set OutlookTitle1.Picture = fMainForm.imlToolbarIcons.ListImages("Validate Transactions").Picture
Image1.Picture = PictureClip1.GraphicCell(0)
End Sub


Public Sub PrintReports()

On Error GoTo Report1Err
miPrintMode = 0
'View the Deceased report
ViewPrintReport 1, ""
'View the pre edit report
ViewPrintReport 2, ""

Xit:
    Exit Sub
Report1Err:
    Hourglass False
    MsgBox MODULE & "Print Reports" & vbCrLf & Error, vbInformation
    Resume Xit
End Sub

Private Sub OutlookTitle1_IconClick()
    If cmdCancel.Enabled = True Then
        Unload Me
    End If

End Sub
