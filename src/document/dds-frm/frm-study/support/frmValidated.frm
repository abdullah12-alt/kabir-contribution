VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmValidated 
   Caption         =   "Validated FUNB Transactions"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   10050
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
      Height          =   4245
      Left            =   210
      TabIndex        =   1
      Top             =   135
      Width           =   9600
      _Version        =   196616
      UseGroups       =   -1  'True
      ForeColorEven   =   0
      BackColorOdd    =   8454143
      Levels          =   2
      RowHeight       =   847
      Groups(0).Width =   16219
      Groups(0).Caption=   "Valid Records"
      Groups(0).Columns.Count=   8
      Groups(0).Columns(0).Width=   2143
      Groups(0).Columns(0).Caption=   "Trans Date"
      Groups(0).Columns(0).Name=   "AS_OF_DATETIME"
      Groups(0).Columns(0).Alignment=   1
      Groups(0).Columns(0).CaptionAlignment=   1
      Groups(0).Columns(0).DataField=   "AS_OF_DATETIME"
      Groups(0).Columns(0).DataType=   135
      Groups(0).Columns(0).NumberFormat=   "MM/dd/yyyy"
      Groups(0).Columns(0).FieldLen=   256
      Groups(0).Columns(1).Width=   3360
      Groups(0).Columns(1).Caption=   "Affinity Acct Num"
      Groups(0).Columns(1).Name=   "AFFINITY_ACCT_NUM"
      Groups(0).Columns(1).CaptionAlignment=   0
      Groups(0).Columns(1).DataField=   "AFFINITY_ACCT_NUM"
      Groups(0).Columns(1).DataType=   8
      Groups(0).Columns(1).FieldLen=   256
      Groups(0).Columns(2).Width=   4736
      Groups(0).Columns(2).Caption=   "Patient Name"
      Groups(0).Columns(2).Name=   "PATIENT_NAME"
      Groups(0).Columns(2).CaptionAlignment=   0
      Groups(0).Columns(2).DataField=   "PATIENT_NAME"
      Groups(0).Columns(2).DataType=   8
      Groups(0).Columns(2).FieldLen=   256
      Groups(0).Columns(3).Width=   2884
      Groups(0).Columns(3).Caption=   "Personal Funds Amt "
      Groups(0).Columns(3).Name=   "PF_DISTRIBUTION_AMOUNT"
      Groups(0).Columns(3).Alignment=   1
      Groups(0).Columns(3).CaptionAlignment=   1
      Groups(0).Columns(3).DataField=   "PF_DISTRIBUTION_AMT"
      Groups(0).Columns(3).DataType=   5
      Groups(0).Columns(3).FieldLen=   256
      Groups(0).Columns(4).Width=   3096
      Groups(0).Columns(4).Caption=   "Patient Account Amt"
      Groups(0).Columns(4).Name=   "PA_DISTRIBUTION_AMOUNT"
      Groups(0).Columns(4).Alignment=   1
      Groups(0).Columns(4).CaptionAlignment=   1
      Groups(0).Columns(4).DataField=   "PA_DISTRIBUTION_AMT"
      Groups(0).Columns(4).DataType=   5
      Groups(0).Columns(4).FieldLen=   256
      Groups(0).Columns(5).Width=   5503
      Groups(0).Columns(5).Caption=   "Direct Deposit Number"
      Groups(0).Columns(5).Name=   "DD_NUM"
      Groups(0).Columns(5).CaptionAlignment=   0
      Groups(0).Columns(5).DataField=   "DD_NUM"
      Groups(0).Columns(5).DataType=   8
      Groups(0).Columns(5).Level=   1
      Groups(0).Columns(5).FieldLen=   256
      Groups(0).Columns(6).Width=   7620
      Groups(0).Columns(6).Caption=   "Medical Record Number"
      Groups(0).Columns(6).Name=   "MEDICAL_RECORD_NUM"
      Groups(0).Columns(6).CaptionAlignment=   0
      Groups(0).Columns(6).DataField=   "MEDICAL_RECORD_NUM"
      Groups(0).Columns(6).DataType=   8
      Groups(0).Columns(6).Level=   1
      Groups(0).Columns(6).NumberFormat=   "0-00-00-00"
      Groups(0).Columns(6).FieldLen=   256
      Groups(0).Columns(7).Width=   3096
      Groups(0).Columns(7).Caption=   "Institution"
      Groups(0).Columns(7).Name=   "INSTITUTION_CODE"
      Groups(0).Columns(7).CaptionAlignment=   0
      Groups(0).Columns(7).DataField=   "INSTITUTION_CODE"
      Groups(0).Columns(7).DataType=   8
      Groups(0).Columns(7).Level=   1
      Groups(0).Columns(7).FieldLen=   256
      _ExtentX        =   16933
      _ExtentY        =   7488
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   8490
      TabIndex        =   0
      Top             =   4620
      Width           =   1215
   End
End
Attribute VB_Name = "frmValidated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mrsValid As New ADODB.Recordset

Private Sub cmdClose_Click()
    
    frmEditDeposits.Refresh
    Unload Me
    
    
End Sub

Private Sub Form_Load()
Dim sSql As String
On Error GoTo LoadErr

sSql = "SELECT * FROM DD_VALID_REC"
mrsValid.Open sSql, gcnDDS, adOpenStatic
Set SSOleDBGrid1.DataSource = mrsValid
SSOleDBGrid1.Refresh
Exit Sub

LoadErr:
MsgBox Error

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mrsValid = Nothing
End Sub

