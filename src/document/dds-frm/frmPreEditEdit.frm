VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmPreEditEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit a Pre-Edit record"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmPreEditEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo dbcIncomeSource 
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
      DataFieldList   =   "Column 0"
      AllowInput      =   0   'False
      _Version        =   196616
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldSeparator  =   ","
      ForeColorEven   =   0
      BackColorOdd    =   8454143
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   2381
      Columns(0).Caption=   "Income Source"
      Columns(0).Name =   "Income Source"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4895
      Columns(1).Caption=   "Description"
      Columns(1).Name =   "Description"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   4048
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtMemo 
      Height          =   1335
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2040
      Width           =   6615
   End
   Begin VB.TextBox txtDDNum 
      Height          =   375
      Left            =   720
      MaxLength       =   11
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Frame fraPatientInfo 
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6615
      Begin VB.Label lblAccountNumber 
         Height          =   255
         Left            =   4200
         TabIndex        =   16
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Account #:"
         Height          =   255
         Left            =   3240
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblInstitutionCode 
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblBasicInfo 
         Caption         =   "MRUN:"
         Height          =   255
         Index           =   6
         Left            =   4395
         TabIndex        =   10
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblBasicInfo 
         Caption         =   "Name:"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblBasicInfo 
         Caption         =   "Institution Code:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblName 
         Height          =   255
         Left            =   660
         TabIndex        =   7
         Top             =   240
         Width           =   3330
      End
      Begin VB.Label lblMRUN 
         Height          =   255
         Left            =   5025
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Income Source Type:"
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Memo:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "DD #:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmPreEditEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sIncomeSourceSql As String
Private rsIncomeSource As New ADODB.Recordset
Private moPatientInfo As clsPatientPreEditInfo
Private mnVBMsgBoxResult As VbMsgBoxResult
Private Const MODULE = "frmPreEditEdit"

Public Property Get MsgBoxResult() As VbMsgBoxResult
    MsgBoxResult = mnVBMsgBoxResult
End Property

Public Sub setPatientPreEditInfo(ByRef PatientInfo As clsPatientPreEditInfo)
    Set moPatientInfo = PatientInfo
    With moPatientInfo
        lblName.Caption = .Name
        lblMRUN.Caption = .MRUN
        lblInstitutionCode.Caption = .InstitutionCode
        lblAccountNumber.Caption = .AccountNum
        txtDDNum.Text = .DDNum
        dbcIncomeSource.Text = .IncomeSource
        txtMemo.Text = .Memo
    End With
    cmdOK.Enabled = False
End Sub
Private Sub RefreshIncomeSourceCombo(ByRef sSql As String)

On Error GoTo RefreshIncomeSourceComboErr

    Hourglass True
    With dbcIncomeSource
        If rsIncomeSource.State = adStateOpen Then
            rsIncomeSource.Close
        End If
        rsIncomeSource.Open sSql, gcnDDS, adOpenDynamic, adLockOptimistic
        Do Until rsIncomeSource.EOF
            .AddItem rsIncomeSource.Fields(0)
            rsIncomeSource.MoveNext
        Loop
    End With
    Hourglass False

Xit:
    Exit Sub

RefreshIncomeSourceComboErr:
    ShowUnexpectedError MODULE + " RefreshIncomeSourceCombo", Err
    Resume Xit


End Sub


Private Sub cmdCancel_Click()
    mnVBMsgBoxResult = vbCancel
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    mnVBMsgBoxResult = vbOK
    With moPatientInfo
        .DDNum = txtDDNum.Text
        .IncomeSource = dbcIncomeSource.Text
        .Memo = txtMemo.Text
    End With
    Me.Hide
End Sub

Private Sub dbcIncomeSource_Change()
    cmdOK.Enabled = True
End Sub

Private Sub dbcIncomeSource_Click()
    
    cmdOK.Enabled = True

End Sub

Private Sub Form_Load()
    CenterMe Me
    sIncomeSourceSql = "SELECT FUNB_INCOME_SRC_TYPE + ',' + INCOME_SRC_TYPE_DESCR " _
                            & "FROM DD_INCOME_SOURCE_TYPE ORDER BY FUNB_INCOME_SRC_TYPE"
    RefreshIncomeSourceCombo (sIncomeSourceSql)
    
End Sub

Private Sub txtDDNum_Change()
    cmdOK.Enabled = True
End Sub

Private Sub txtDDNum_GotFocus()
 SetSelected
 
End Sub

Private Sub txtMemo_Change()
    cmdOK.Enabled = True
End Sub

Private Sub txtMemo_GotFocus()
SetSelected
End Sub
