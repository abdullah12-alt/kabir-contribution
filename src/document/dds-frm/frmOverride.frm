VERSION 5.00
Begin VB.Form frmOverride 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Force a Direct Deposit"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   600
      TabIndex        =   42
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   6705
      TabIndex        =   6
      Top             =   5820
      Width           =   1590
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resulting Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   120
      TabIndex        =   12
      Top             =   2505
      Width           =   8250
      Begin VB.Label lblFields 
         Height          =   270
         Index           =   0
         Left            =   1605
         TabIndex        =   41
         Top             =   390
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DD Number:"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   390
         Width           =   1350
      End
      Begin VB.Label lblFields 
         Height          =   270
         Index           =   1
         Left            =   1605
         TabIndex        =   39
         Top             =   765
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DD Number:"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   765
         Width           =   1350
      End
      Begin VB.Label lblFields 
         Height          =   270
         Index           =   2
         Left            =   1605
         TabIndex        =   37
         Top             =   1140
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DD Number:"
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   1140
         Width           =   1350
      End
      Begin VB.Label lblFields 
         Height          =   270
         Index           =   3
         Left            =   1605
         TabIndex        =   35
         Top             =   1515
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DD Number:"
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   34
         Top             =   1515
         Width           =   1350
      End
      Begin VB.Label lblFields 
         Height          =   270
         Index           =   4
         Left            =   1605
         TabIndex        =   33
         Top             =   1890
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DD Number:"
         Height          =   270
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   1890
         Width           =   1350
      End
      Begin VB.Label lblFields 
         Height          =   270
         Index           =   5
         Left            =   1605
         TabIndex        =   31
         Top             =   2265
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DD Number:"
         Height          =   270
         Index           =   5
         Left            =   120
         TabIndex        =   30
         Top             =   2265
         Width           =   1350
      End
      Begin VB.Label lblFields 
         Height          =   300
         Index           =   6
         Left            =   1605
         TabIndex        =   29
         Top             =   2655
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DD Number:"
         Height          =   270
         Index           =   6
         Left            =   120
         TabIndex        =   28
         Top             =   2655
         Width           =   1350
      End
      Begin VB.Label lblFields 
         Height          =   270
         Index           =   7
         Left            =   5655
         TabIndex        =   27
         Top             =   390
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DD Number:"
         Height          =   270
         Index           =   7
         Left            =   4170
         TabIndex        =   26
         Top             =   390
         Width           =   1350
      End
      Begin VB.Label lblFields 
         Height          =   270
         Index           =   8
         Left            =   5655
         TabIndex        =   25
         Top             =   765
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DD Number:"
         Height          =   270
         Index           =   8
         Left            =   4170
         TabIndex        =   24
         Top             =   765
         Width           =   1350
      End
      Begin VB.Label lblFields 
         Height          =   270
         Index           =   9
         Left            =   5655
         TabIndex        =   23
         Top             =   1140
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DD Number:"
         Height          =   270
         Index           =   9
         Left            =   4170
         TabIndex        =   22
         Top             =   1140
         Width           =   1350
      End
      Begin VB.Label lblFields 
         Height          =   270
         Index           =   10
         Left            =   5655
         TabIndex        =   21
         Top             =   1515
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DD Number:"
         Height          =   270
         Index           =   10
         Left            =   4170
         TabIndex        =   20
         Top             =   1515
         Width           =   1350
      End
      Begin VB.Label lblFields 
         Height          =   270
         Index           =   11
         Left            =   5655
         TabIndex        =   19
         Top             =   1890
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DD Number:"
         Height          =   270
         Index           =   11
         Left            =   4170
         TabIndex        =   18
         Top             =   1890
         Width           =   1350
      End
      Begin VB.Label lblFields 
         Height          =   270
         Index           =   12
         Left            =   5655
         TabIndex        =   17
         Top             =   2265
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DD Number:"
         Height          =   270
         Index           =   12
         Left            =   4170
         TabIndex        =   16
         Top             =   2265
         Width           =   1350
      End
      Begin VB.Label lblFields 
         Height          =   330
         Index           =   13
         Left            =   5655
         TabIndex        =   15
         Top             =   2655
         Width           =   2340
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "DD Number:"
         Height          =   270
         Index           =   13
         Left            =   4170
         TabIndex        =   14
         Top             =   2655
         Width           =   1350
      End
      Begin VB.Label lblFields 
         Height          =   240
         Index           =   14
         Left            =   5655
         TabIndex        =   13
         Top             =   2265
         Width           =   2340
      End
   End
   Begin VB.CommandButton cmdForce 
      Caption         =   "Force deposit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5010
      TabIndex        =   5
      Top             =   5820
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter the account number and amounts you want to override"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   150
      TabIndex        =   7
      Top             =   105
      Width           =   8205
      Begin VB.TextBox txtAccountNumber 
         Height          =   315
         Left            =   2670
         TabIndex        =   0
         Top             =   375
         Width           =   2070
      End
      Begin VB.TextBox txtAccountAmount 
         Height          =   315
         Left            =   2670
         TabIndex        =   1
         Top             =   815
         Width           =   2070
      End
      Begin VB.TextBox txtPFAmount 
         Height          =   315
         Left            =   2670
         TabIndex        =   2
         Top             =   1255
         Width           =   2070
      End
      Begin VB.TextBox txtATPPML 
         Height          =   315
         Left            =   2670
         TabIndex        =   3
         Top             =   1695
         Width           =   2070
      End
      Begin VB.CommandButton cmdValidate 
         Caption         =   "Validate"
         Height          =   450
         Left            =   5145
         TabIndex        =   4
         Top             =   375
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Account Number:"
         Height          =   285
         Left            =   555
         TabIndex        =   11
         Top             =   405
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Account Amount:"
         Height          =   285
         Left            =   555
         TabIndex        =   10
         Top             =   855
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Personal Funds Amount:"
         Height          =   285
         Left            =   600
         TabIndex        =   9
         Top             =   1305
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "ATP/PML:"
         Height          =   285
         Left            =   630
         TabIndex        =   8
         Top             =   1755
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmOverride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrs As New ADODB.Recordset
Dim msVisitId As String
Dim msMemo As String
Dim msIncomeTypeID As String

Private Function DetermineHosptialByAcct(ByVal sAcct As String, ByRef sHospitalCode As String, ByRef sHospitalDSN As String) As Boolean

sHospitalCode = ""
sHospitalDSN = ""

Select Case Left$(sAcct, 1)
Case "0"
    sHospitalCode = "0"
    sHospitalDSN = "JUH"
Case "1"
    sHospitalCode = "1"
    sHospitalDSN = "CHERRY"
Case "2"
    sHospitalCode = "2"
    sHospitalDSN = "BROUGHT"
Case "3"
    sHospitalCode = "3"
    sHospitalDSN = "DIX"
Case "4"
    sHospitalCode = "4"
    sHospitalDSN = "MURDOCH"
Case "5"
    sHospitalCode = "5"
    sHospitalDSN = "OBERRY"
Case "6"
    sHospitalCode = "6"
    sHospitalDSN = "CASWELL"
Case "7"
    sHospitalCode = "7"
    sHospitalDSN = "WCAROLINA"
Case "8"
    sHospitalCode = "9"
    sHospitalDSN = "NCSCC"
Case "9"
    Select Case Mid$(sAcct, 2, 1)
    Case "0"
        sHospitalCode = "E"
        sHospitalDSN = "BLACKMNT"
    Case "6"
        sHospitalCode = "H"
        sHospitalDSN = "JFKADATC"
    Case "2"
        sHospitalCode = "Q"
        sHospitalDSN = "WBJADATC"
    End Select
End Select

If sHospitalCode = vbNullString Then
    DetermineHosptialByAcct = False
Else
    DetermineHosptialByAcct = True
End If


End Function

Private Sub cmdForce_Click()
Hourglass True
If LockApplication = True Then
    ForceToValid
    UnlockApplication
End If
Hourglass False
Unload Me

End Sub

Private Sub cmdValidate_Click()
On Error GoTo cmdValidateErr
Dim cnn As New ADODB.Connection
Dim rsAff As New ADODB.Recordset
Dim rsADO As New ADODB.Recordset
Dim sHospitalCode As String
Dim sHospitalDSN As String
Dim sSql As String
Dim iTries As Integer
Dim strConnect As String
Hourglass True

'Check to see if the user put in an A or P
If txtATPPML = "A" Or txtATPPML = "P" Then
    'Passed Check
Else
    Err.Raise 54768, , "ATP PML must be either A or P"
End If

'Check to see if account amount and personal funds amount are not zero and equal the total amount
If txtAccountAmount = vbNullString Then
    Err.Raise 56789, , "Account Amount must be filled in"
End If

If txtPFAmount = vbNullString Then
    Err.Raise 56790, , "Personal Funds Amount must be filled in"
End If

If CCur(txtPFAmount) + CCur(txtAccountAmount) <> CCur(lblFields(11)) Then
    Err.Raise 54787, , "The Account Amount + PF Amount must equal the FUNB Total Amount"
End If
'Check to see if the account exists in affinity
If Len(txtAccountNumber) <> 8 Or Not IsNumeric(txtAccountNumber) Then
    Err.Raise 54654, , "Account number must be eight digits"
End If
If DetermineHosptialByAcct(txtAccountNumber, sHospitalCode, sHospitalDSN) = False Then
    Err.Raise 54654, , "Cannot determine the institution"
End If

'Set db = OpenDatabase(sHospitalDSN, dbDriverNoPrompt, False, "ODBC;UID=dhhsHearts;PWD=;DSN=" & sHospitalDSN & ";DATABASE=" & sHospitalDSN)
'AS-2/27/2014 - Using ADODB to open database
'Set db = OpenDatabase(App.Path & "\affdata\" & sHospitalDSN & ".mdb")

'strConnect = "DRIVER={InterSystems ODBC};SERVER=hes001.dhr.state.nc.us;PORT=1972;DATABASE=" & sHospitalDSN & ";STATIC CURSORS=1;AUTHENTICATION METHOD=0;UID=dhhsHearts;PWD=apollo30;"
strConnect = "DRIVER={InterSystems IRIS ODBC35};SERVER=hes001.dhr.state.nc.us;PORT=1972;DATABASE=" & sHospitalDSN & ";STATIC CURSORS=1;AUTHENTICATION METHOD=0;UID=dhhsHearts;PWD=apollo30;"
cnn.CursorLocation = adUseClient
cnn.ConnectionTimeout = 1000
cnn.CommandTimeout = 1000
'Open the main connection
cnn.ConnectionString = strConnect
cnn.Open

sSql = "SELECT VISIT.PATIENT_ACCOUNT_NUMBER, VISIT.VISIT_ID, PATIENT.MRUN, PATIENT.NAME, PATIENT.DEATH_FLAG"
sSql = sSql & " FROM VISIT INNER JOIN PATIENT ON VISIT.PATIENT_ID = PATIENT.PATIENT_ID"
sSql = sSql & " WHERE VISIT.PATIENT_ACCOUNT_NUMBER='" & txtAccountNumber & "'"
rsAff.Open sSql, cnn
If rsAff.EOF Then
    Err.Raise 54654, , "Account was not found in Affinity"
Else
    If txtPFAmount > 0 Then
        rsADO.Open "SELECT PATIENT_ID FROM PF_PATIENT WHERE MEDICAL_RECORD_NUM = '" & Format$(rsAff!MRUN, "0000000") & "'", gcnPFS
        If rsADO.EOF Then
            Err.Raise 56765, , "Personal Funds Account not established"
        End If
    End If
    lblFields(3) = sHospitalCode
    lblFields(2) = rsAff!Name
    lblFields(4) = Format$(rsAff!PATIENT_ACCOUNT_NUMBER, "00000000")
    lblFields(5) = Format$(rsAff!MRUN, "0000000")
    msVisitId = rsAff!VISIT_ID
    If rsAff!DEATH_FLAG = "Y" Then
        lblFields(10) = "Y"
    Else
        lblFields(10) = "N"
    End If
    lblFields(7) = txtATPPML
    lblFields(12) = Format$(txtAccountAmount, "Currency")
    lblFields(13) = Format$(txtPFAmount, "Currency")
    
End If
    
cmdForce.Enabled = True
MsgBox "The record is now ready to be forced", vbInformation
Xit:
Set rsAff = Nothing
Set rsADO = Nothing
If cnn.State <> 0 Then
    cnn.Close
End If
Set cnn = Nothing
Hourglass False
Exit Sub

cmdValidateErr:
If Err = 3146 Or Err = 13 Then
    iTries = iTries + 1
    If iTries >= 5 Then
        MsgBox "Could not get Affinity Information"
        Resume Xit
    Else
        Resume
    End If
Else
    MsgBox Error, vbInformation
    Resume Xit
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo cmdTestErr
Dim cnn As New ADODB.Connection
Dim rsADO As New ADODB.Recordset
Dim rsAff As New ADODB.Recordset
Dim strConnect As String
Dim sSql As String
'strConnect = "DSN=JUH;Uid=dhhsHearts;Pwd=apollo30;"
'strConnect = "Provider=MSDASQL.1;Extended Properties='SERVER=hes001.dhr.state.nc.us;PORT=1972;DATABASE=JUH;AUTHENTICATION METHOD=0;UID=dhhsHearts;PWD=apollo30;STATIC CURSORS=0;QUERY TIMEOUT=0;UNICODE SQLTYPES=0'"
'strConnect = "DRIVER={InterSystems ODBC};SERVER=hes001.dhr.state.nc.us;PORT=1972;DATABASE=JUH;STATIC CURSORS=1;AUTHENTICATION METHOD=0;UID=dhhsHearts;PWD=apollo30;"
strConnect = "DRIVER={InterSystems IRIS ODBC35};SERVER=hes001.dhr.state.nc.us;PORT=1972;DATABASE=JUH;STATIC CURSORS=1;AUTHENTICATION METHOD=0;UID=dhhsHearts;PWD=apollo30;"
cnn.CursorLocation = adUseClient
cnn.ConnectionTimeout = 1000
cnn.CommandTimeout = 1000
'Open the main connection
cnn.ConnectionString = strConnect
cnn.Open
sSql = sSql & " SELECT VISIT.PATIENT_ACCOUNT_NUMBER, VISIT.VISIT_ID, PATIENT.MRUN, PATIENT.NAME, PATIENT.DEATH_FLAG"
sSql = sSql & " FROM VISIT INNER JOIN PATIENT ON VISIT.PATIENT_ID = PATIENT.PATIENT_ID"
sSql = sSql & " WHERE VISIT.PATIENT_ACCOUNT_NUMBER)= " & txtAccountNumber & ";"
rsAff.Open sSql, cnn

MsgBox "The record is now ready to be forced", vbInformation
Xit:

Set rsADO = Nothing
Set rsAff = Nothing
cnn.Close
Set cnn = Nothing

Hourglass False
Exit Sub

cmdTestErr:
    MsgBox Error, vbInformation
    Resume


End Sub

Private Sub Form_Activate()
Dim sSql As String
sSql = "SELECT BAI_FILE_ID,DD_NUM,DD_INVALID_REC.FUNB_INCOME_SRC_TYPE,DD_INCOME_SOURCE_TYPE.INCOME_SOURCE_TYPE_ID,PATIENT_NAME,INSTITUTION_CODE,AFFINITY_ACCT_NUM,MEDICAL_RECORD_NUM,DR_CR_FLAG,COMMENT,TOT_FUNB_BENEFIT_AMT,AS_OF_DATETIME,DD_INVALID_REC.CREATED_DATETIME, DECEASED_IND, SHARED_DD_NUM_IND,INVALID_RECORD_ID"
sSql = sSql & " FROM DD_INVALID_REC,DD_INCOME_SOURCE_TYPE"
sSql = sSql & " WHERE DD_INVALID_REC.FUNB_INCOME_SRC_TYPE = DD_INCOME_SOURCE_TYPE.FUNB_INCOME_SRC_TYPE"
sSql = sSql & " AND INVALID_RECORD_ID = " & lblFields(14)
mrs.Open sSql, gcnDDS
If mrs.EOF Then
    MsgBox "Error getting Income Source record", vbInformation
    Unload Me
    Exit Sub
End If
    
lblFields(0).Caption = mrs.Fields("DD_NUM")
lblFields(1).Caption = mrs.Fields("FUNB_INCOME_SRC_TYPE")
lblFields(2).Caption = CNull(mrs.Fields("PATIENT_NAME"))
lblFields(3).Caption = CNull(mrs.Fields("INSTITUTION_CODE"))
If IsNull(mrs!AFFINITY_ACCT_NUM) Then
    lblFields(4).Caption = ""
Else
    lblFields(4).Caption = Format$(mrs.Fields!AFFINITY_ACCT_NUM, "00000000")
End If
If IsNull(mrs!MEDICAL_RECORD_NUM) Then
    lblFields(5).Caption = ""
Else
    lblFields(5).Caption = Format$(mrs!MEDICAL_RECORD_NUM, "0000000")
End If
lblFields(6).Caption = CNull(mrs.Fields("DR_CR_FLAG"))
lblFields(7).Caption = ""
lblFields(8).Caption = CNull(mrs.Fields("AS_OF_DATETIME"))
lblFields(9).Caption = CNull(mrs.Fields("CREATED_DATETIME"))
lblFields(10).Caption = CNull(mrs.Fields("DECEASED_IND"))
lblFields(11).Caption = Format$(CNullToZero(mrs.Fields("TOT_FUNB_BENEFIT_AMT")), "Currency")
lblFields(12).Caption = Format$(0, "Currency")
lblFields(13).Caption = Format$(0, "Currency")

lblLabels(0).Caption = "DD Number:"
lblLabels(1).Caption = "Income Source:"
lblLabels(2).Caption = "Name:"
lblLabels(3).Caption = "Institution:"
lblLabels(4).Caption = "Account Number:"
lblLabels(5).Caption = "MRUN:"
lblLabels(6).Caption = "Debit/Credit:"
lblLabels(7).Caption = "ATP/PML:"
lblLabels(8).Caption = "FUNB As Of Date:"
lblLabels(9).Caption = "Created Date:"
lblLabels(10).Caption = "Deceased:"
lblLabels(11).Caption = "FUNB Amount:"
lblLabels(12).Caption = "Account Amount:"
lblLabels(13).Caption = "PF Amount:"



End Sub

Private Function CNull(ByVal v As Variant) As String
If IsNull(v) Then
    CNull = vbNullString
Else
    CNull = CStr(v)
End If

End Function
Private Function CNullToZero(ByVal v As Variant) As String
If IsNull(v) Then
    CNullToZero = 0
Else
    CNullToZero = CStr(v)
End If

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set mrs = Nothing

End Sub


Private Sub txtATPPML_LostFocus()

txtATPPML = UCase(txtATPPML)

End Sub

Private Sub ForceToValid()

Dim cmdValid As New ADODB.Command
    On Error GoTo ForceToValidErr

    'Set up all the stored procedures to be used
    Set cmdValid.ActiveConnection = gcnDDS
    If gStoredProcs("up_iu_ValidRecords").GetStoredProcCommand(cmdValid) = False Then
        Err.Raise 677666, , "Stored Procedure could not be created"
    End If

    'Add a record to the valid records table
    cmdValid.Parameters("work_file_record_id") = mrs!INVALID_RECORD_ID
    cmdValid.Parameters("valid_record_id") = Null
    cmdValid.Parameters("income_source_type_id") = mrs!INCOME_SOURCE_TYPE_ID
    cmdValid.Parameters("bai_file_id") = mrs!BAI_FILE_ID
    cmdValid.Parameters("institution_code") = lblFields(3)
    cmdValid.Parameters("affinity_acct_num") = lblFields(4)
    cmdValid.Parameters("affinity_visit_id") = msVisitId
    cmdValid.Parameters("medical_record_num") = lblFields(5)
    cmdValid.Parameters("dd_num") = mrs!DD_NUM
    cmdValid.Parameters("atp_pml_flag") = lblFields(7)
    cmdValid.Parameters("affinity_atp_rate_id") = 0
    cmdValid.Parameters("tot_funb_benefit_amt") = mrs!TOT_FUNB_BENEFIT_AMT
    cmdValid.Parameters("dr_cr_flag") = mrs!DR_CR_FLAG
    cmdValid.Parameters("as_of_datetime") = mrs!AS_OF_DATETIME
    cmdValid.Parameters("patient_name") = lblFields(2)
    cmdValid.Parameters("pf_distribution_amt") = CCur(lblFields(13))
    cmdValid.Parameters("pa_distribution_amt") = CCur(lblFields(12))
    cmdValid.Parameters("deceased_ind") = lblFields(10)
    cmdValid.Parameters("created_by") = gobjLoginInfo.UserId
    cmdValid.Parameters("tot_days_inhouse") = 0
    cmdValid.Parameters("spec_proc_cond_hash_tot") = 0
    cmdValid.Parameters("sent_for_posting_datetime") = Null
    cmdValid.Parameters("update_status") = "I"
    cmdValid.Parameters("posted_to_affinity") = Null
    cmdValid.Parameters("override") = "Y"
    cmdValid.Execute
    If cmdValid.Parameters("RETURN_VALUE") <> 0 Then
        MsgBox "Error forcing record", vbInformation
    End If
    
Xit:
    Set cmdValid = Nothing
    Exit Sub

ForceToValidErr:
    MsgBox Error, vbInformation
    Resume Xit


End Sub
