VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{8CD222DF-7752-11D3-9D1E-00105A19BCF2}#1.0#0"; "OAOTBar.ocx"
Begin VB.Form frmBalance 
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10485
   ControlBox      =   0   'False
   Icon            =   "frmBalance.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   10485
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   3345
      Left            =   105
      TabIndex        =   12
      Top             =   630
      Width           =   4230
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid sdgSummary 
         Height          =   3195
         Left            =   45
         TabIndex        =   13
         Top             =   120
         Width           =   4155
         ScrollBars      =   0
         _Version        =   196616
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         RecordSelectors =   0   'False
         GroupHeaders    =   0   'False
         GroupHeadLines  =   0
         HeadLines       =   0
         Col.Count       =   2
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         MaxSelectedRows =   0
         ForeColorEven   =   0
         BackColorOdd    =   8454143
         RowHeight       =   423
         ExtraHeight     =   79
         Columns.Count   =   2
         Columns(0).Width=   5292
         Columns(0).Caption=   "Desc"
         Columns(0).Name =   "Desc"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3519
         Columns(1).Caption=   "Amount"
         Columns(1).Name =   "Amount"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   7329
         _ExtentY        =   5636
         _StockProps     =   79
         Caption         =   "Auto Balance Summary"
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
   End
   Begin VB.PictureBox picBegBalance 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3315
      Left            =   4365
      ScaleHeight     =   3315
      ScaleWidth      =   4350
      TabIndex        =   4
      Top             =   780
      Visible         =   0   'False
      Width           =   4350
      Begin VB.TextBox txtBegBalance 
         Height          =   405
         Left            =   1845
         TabIndex        =   8
         Top             =   1815
         Width           =   1845
      End
      Begin VB.CommandButton cmdBalanceCancel 
         Caption         =   "&Cancel"
         Height          =   465
         Left            =   2145
         TabIndex        =   7
         Top             =   2505
         Width           =   1500
      End
      Begin VB.CommandButton cmdBalanceOK 
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   465
         Left            =   480
         TabIndex        =   6
         Top             =   2505
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808000&
         BorderWidth     =   5
         Height          =   3225
         Left            =   30
         Top             =   60
         Width           =   4290
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Beginning Balance:"
         Height          =   330
         Left            =   225
         TabIndex        =   9
         Top             =   1890
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   $"frmBalance.frx":000C
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         Left            =   195
         TabIndex        =   5
         Top             =   210
         Width           =   3915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pre-Edit File Records"
      Height          =   3330
      Left            =   4380
      TabIndex        =   1
      Top             =   660
      Width           =   5385
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid sdgInvalid 
         Height          =   2895
         Left            =   105
         TabIndex        =   11
         Top             =   330
         Width           =   5130
         _Version        =   196616
         DataMode        =   1
         RecordSelectors =   0   'False
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         ForeColorEven   =   0
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   5
         Columns(0).Width=   1852
         Columns(0).Caption=   "DD Number"
         Columns(0).Name =   "DD_NUM"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2752
         Columns(1).Caption=   "FUNB Income Type"
         Columns(1).Name =   "FUNB_INCOME_SRC_TYPE"
         Columns(1).CaptionAlignment=   0
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1693
         Columns(2).Caption=   "Amount"
         Columns(2).Name =   "TOT_FUNB_BENEFIT_AMT"
         Columns(2).Alignment=   1
         Columns(2).CaptionAlignment=   1
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   5
         Columns(2).FieldLen=   256
         Columns(3).Width=   794
         Columns(3).Caption=   "D/C"
         Columns(3).Name =   "DR_CR_FLAG"
         Columns(3).CaptionAlignment=   0
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   1535
         Columns(4).Caption=   "Deceased"
         Columns(4).Name =   "DECEASED_IND"
         Columns(4).CaptionAlignment=   0
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         _ExtentX        =   9049
         _ExtentY        =   5106
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
   End
   Begin OAOTitleBar.OutlookTitleBar OutlookTitle1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   979
      ForeColor       =   16777215
      Caption         =   "Auto Balance Direct Deposit"
   End
   Begin VB.CommandButton cmdFinish 
      Cancel          =   -1  'True
      Caption         =   "&Finish"
      Height          =   465
      Left            =   5865
      TabIndex        =   3
      Top             =   6600
      Width           =   1815
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid sdgSummaryRecs 
      Height          =   2295
      Left            =   150
      TabIndex        =   2
      Top             =   4155
      Width           =   9555
      _Version        =   196616
      DataMode        =   1
      GroupHeadLines  =   0
      UseGroups       =   -1  'True
      AllowUpdate     =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   2
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   2
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      ForeColorEven   =   0
      BackColorOdd    =   8454143
      Levels          =   2
      RowHeight       =   847
      Groups(0).Width =   15849
      Groups(0).Caption=   "My Group"
      Groups(0).Columns.Count=   9
      Groups(0).Columns(0).Width=   3069
      Groups(0).Columns(0).Caption=   "FUNB Date"
      Groups(0).Columns(0).Name=   "BAI_FILE_DATETIME"
      Groups(0).Columns(0).Alignment=   1
      Groups(0).Columns(0).CaptionAlignment=   1
      Groups(0).Columns(0).DataField=   "Column 0"
      Groups(0).Columns(0).DataType=   7
      Groups(0).Columns(0).NumberFormat=   "MM/dd/yyyy"
      Groups(0).Columns(0).FieldLen=   10
      Groups(0).Columns(1).Width=   4260
      Groups(0).Columns(1).Caption=   "Ledger Balance"
      Groups(0).Columns(1).Name=   "LEDGER_BALANCE"
      Groups(0).Columns(1).Alignment=   1
      Groups(0).Columns(1).CaptionAlignment=   1
      Groups(0).Columns(1).DataField=   "Column 1"
      Groups(0).Columns(1).DataType=   5
      Groups(0).Columns(1).FieldLen=   256
      Groups(0).Columns(2).Width=   4286
      Groups(0).Columns(2).Caption=   "Available Balance"
      Groups(0).Columns(2).Name=   "AVAILABLE_BALANCE"
      Groups(0).Columns(2).Alignment=   1
      Groups(0).Columns(2).CaptionAlignment=   1
      Groups(0).Columns(2).DataField=   "Column 2"
      Groups(0).Columns(2).DataType=   5
      Groups(0).Columns(2).FieldLen=   256
      Groups(0).Columns(3).Width=   4233
      Groups(0).Columns(3).Caption=   "Collected Balance"
      Groups(0).Columns(3).Name=   "COLLECTED_BALANCE"
      Groups(0).Columns(3).Alignment=   1
      Groups(0).Columns(3).CaptionAlignment=   1
      Groups(0).Columns(3).DataField=   "Column 3"
      Groups(0).Columns(3).DataType=   5
      Groups(0).Columns(3).FieldLen=   256
      Groups(0).Columns(4).Width=   3069
      Groups(0).Columns(4).Caption=   "As Of Date"
      Groups(0).Columns(4).Name=   "AS_OF_DATETIME"
      Groups(0).Columns(4).Alignment=   1
      Groups(0).Columns(4).DataField=   "Column 4"
      Groups(0).Columns(4).DataType=   8
      Groups(0).Columns(4).Level=   1
      Groups(0).Columns(4).FieldLen=   10
      Groups(0).Columns(5).Width=   2990
      Groups(0).Columns(5).Caption=   "FUNB Total Credits"
      Groups(0).Columns(5).Name=   "FUNB_TOTAL_CREDITS"
      Groups(0).Columns(5).Alignment=   1
      Groups(0).Columns(5).CaptionAlignment=   1
      Groups(0).Columns(5).DataField=   "Column 5"
      Groups(0).Columns(5).DataType=   5
      Groups(0).Columns(5).Level=   1
      Groups(0).Columns(5).FieldLen=   256
      Groups(0).Columns(6).Width=   2884
      Groups(0).Columns(6).Caption=   "FUNB Total Debits"
      Groups(0).Columns(6).Name=   "FUNB_TOTAL_DEBITS"
      Groups(0).Columns(6).Alignment=   1
      Groups(0).Columns(6).CaptionAlignment=   1
      Groups(0).Columns(6).DataField=   "Column 6"
      Groups(0).Columns(6).DataType=   5
      Groups(0).Columns(6).Level=   1
      Groups(0).Columns(6).FieldLen=   256
      Groups(0).Columns(7).Width=   3440
      Groups(0).Columns(7).Caption=   "Download Date"
      Groups(0).Columns(7).Name=   "CREATED_DATETIME"
      Groups(0).Columns(7).Alignment=   1
      Groups(0).Columns(7).CaptionAlignment=   1
      Groups(0).Columns(7).DataField=   "Column 7"
      Groups(0).Columns(7).DataType=   7
      Groups(0).Columns(7).Level=   1
      Groups(0).Columns(7).NumberFormat=   "MM/dd/yyyy"
      Groups(0).Columns(7).FieldLen=   256
      Groups(0).Columns(8).Width=   3466
      Groups(0).Columns(8).Caption=   "Downloaded By"
      Groups(0).Columns(8).Name=   "CREATED_BY1"
      Groups(0).Columns(8).CaptionAlignment=   0
      Groups(0).Columns(8).DataField=   "Column 8"
      Groups(0).Columns(8).DataType=   8
      Groups(0).Columns(8).Level=   1
      Groups(0).Columns(8).FieldLen=   256
      _ExtentX        =   16854
      _ExtentY        =   4048
      _StockProps     =   79
      Caption         =   "FUNB Summary Totals"
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   465
      Left            =   7785
      TabIndex        =   0
      Top             =   6600
      Width           =   1815
   End
End
Attribute VB_Name = "frmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************************************************
' * Description:
' *  This is the source code for the Auto Balance Direct Deposit
' *
' *
' *
' *
' *
' * Revisions:
' *  7/29/99    fml Added comments.
' *
' *
' ********************************************************************************


' Mod CONSTANTS
Private Const MODULE As String = "Auto Balance Direct Deposit"


' Mod ENUMS


' Mod TYPES


' Mod DECLARES


' Mod VARIABLES

Dim mdInvalidTotalSelected As Currency
Dim mdBegBal As Currency
Dim mdCreditsPosted As Currency
Dim mdCarryOverBal As Currency
Dim mdEndingBal As Currency
Dim mdInvalidTotal As Currency
Dim mdDeceasedTotal As Currency
Dim mdAdjEnd As Currency
Dim mdDifference As Currency
Dim mbBegBalanceNeeded As Boolean
Dim mbValidRecordsFound As Boolean
Dim mdLedgerBal As Currency
Dim mdAdjustments As Currency
Dim mrsSummaryRecs As ADODB.Recordset
Dim mrsInValid As ADODB.Recordset

Private Type udtSortedColumnFlag
    ColIndex As Integer
    Ascending As Boolean
End Type


Option Explicit




Private Sub cmdBalanceCancel_Click()
    
    txtBegBalance = vbNullString
    picBegBalance.Visible = False

End Sub


Private Sub cmdBalanceCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    txtBegBalance = vbNullString
    picBegBalance.Visible = False

End Sub

Private Sub cmdBalanceOK_Click()
    If IsNumeric(txtBegBalance) Then
        Hourglass True
        RefreshGrids
        Hourglass False
        picBegBalance.Visible = False
    Else
        MsgBox "Beginning Balance must be numeric.", vbInformation
        txtBegBalance.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdFinish_Click()

    FinishBalance

End Sub


Private Sub FinishBalance()
On Error GoTo FinishBalanceErr

Dim cmdBalance As New ADODB.Command
Dim bLocked As Boolean
Dim iX As Integer
    Hourglass True
    RefreshGrids
    If CheckForValidRecs = True Then
        Exit Sub
    End If
    Hourglass False

    If LockApplication = False Then
        MsgBox "Another user is loading, validating, posting or balancing Direct Deposit.  Try again later. ", vbInformation
        Exit Sub
    Else
        bLocked = True
    End If
    
    If MsgBox("Are you sure you want to complete the balance process?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        If mdDifference <> 0 Then
            If MsgBox("The calculated difference does not equal zero." & vbCrLf & "Are you sure you want to finish with the balance process?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                UnlockApplication
                Exit Sub
            End If
        End If
                
        Hourglass True
        Set cmdBalance.ActiveConnection = gcnDDS
        If gStoredProcs("up_i_Balance").GetStoredProcCommand(cmdBalance) = False Then
            MsgBox "Balance process could not finish" & "Error creating Insert Balance stored procedure.", vbCritical
        Else
            With cmdBalance
             For iX = 0 To cmdBalance.Parameters.Count - 1
                cmdBalance.Parameters(iX) = Null
             Next iX

            .Parameters("beginning_bal") = mdBegBal
            .Parameters("carryover_bal") = mdCarryOverBal
            .Parameters("tot_cr_dr_posted") = mdCreditsPosted
            .Parameters("ending_bal") = mdEndingBal
            .Parameters("tot_cr_dr_pre_edit_rpt") = mdInvalidTotal
            .Parameters("tot_cr_dr_deceased_except") = mdDeceasedTotal
            .Parameters("adj_ending_bal") = mdAdjEnd
            .Parameters("ledger_bal") = mdLedgerBal
            .Parameters("created_by") = gobjLoginInfo.UserId
            .Parameters("tot_cr_dr_adjustments") = mdAdjustments
            .Execute
            If .Parameters("RETURN_VALUE") <> 0 Then
                MsgBox "Balance process could not finish" & "Error creating Insert Balance stored procedure.", vbCritical
                UnlockApplication
                cmdFinish.Enabled = False
                Exit Sub
            End If
            
            End With
            cmdFinish.Enabled = False
            CreateBalanceReport
        End If
    End If
    UnlockApplication
    
Xit:
    On Error Resume Next
    Set cmdBalance = Nothing
    Hourglass False
    Exit Sub

FinishBalanceErr:
    UnlockApplication
    ShowUnexpectedError "Finish Balance", Err

End Sub


Private Sub CreateBalanceReport()
On Error GoTo CreateBalanceReportErr

miPrintMode = 0
ViewPrintReport 4, , Format(Now, "MM/dd/yyyy")

Xit:
    Hourglass False

Exit Sub
CreateBalanceReportErr:
Hourglass False
ShowError MODULE + "Automated Balancing Worksheet Report", Err
Resume Xit

End Sub


Private Sub Form_Activate()

CheckForValidRecs
If mbBegBalanceNeeded = True Then
    If txtBegBalance = vbNullString Then
        cmdFinish.Enabled = False
        picBegBalance.Visible = True
        txtBegBalance.SetFocus
    End If
End If

End Sub

Private Function CheckForValidRecs() As Boolean

Dim rsValid As New ADODB.Recordset
Dim sSql As String

'See If there are any records in the valid records table
sSql = "SELECT VALID_RECORD_ID FROM DD_VALID_REC"
rsValid.Open sSql, gcnDDS, adOpenForwardOnly
If Not rsValid.EOF Then
    MsgBox "You will not be able to complete the auto balance process." & vbCrLf & "There are valid records that still need to be posted.", vbInformation
    cmdFinish.Enabled = False
    CheckForValidRecs = True
    mbValidRecordsFound = True
Else
    CheckForValidRecs = False
    mbValidRecordsFound = False
End If

rsValid.Close
Set rsValid = Nothing

End Function
Private Sub Form_Load()
Dim sSql As String

Set OutlookTitle1.Picture = fMainForm.imlToolbarIcons.ListImages("Balance Direct Deposit").Picture

'Set the recordsource of the summary records for the previous week
'adcSummaryRecs.ConnectionString = gobjLoginInfo.ConnectString
'adcSummaryRecs.RecordSource = "SELECT BAI_FILE_DATETIME,SubString(LTrim(Str(FILE_ID_NUM)),5,2) + '/' + SubString(LTrim(Str(FILE_ID_NUM)),7,2) + '/' + SubString(LTrim(Str(FILE_ID_NUM)),1,4) As CONVERT_FILE_ID_NUM,AVAILABLE_BAL,COLLECTED_BAL,CREATED_BY,FUNB_TOTAL_CREDITS,FUNB_TOTAL_DEBITS,LEDGER_BAL,CONVERT(VARCHAR(10),CREATED_DATETIME,101) AS CREATED_DATETIME FROM DD_BAI_FILE_SUMMARY WHERE BAI_FILE_DATETIME >= '" & DateAdd("d", -10, Format(Now, "MM/dd/yyyy") & " 00:00:00") & "' ORDER BY FILE_ID_NUM DESC"
'adcSummaryRecs.Refresh

Set mrsSummaryRecs = New ADODB.Recordset
sSql = "SELECT BAI_FILE_DATETIME,LEDGER_BAL,AVAILABLE_BAL,COLLECTED_BAL,SubString(LTrim(Str(FILE_ID_NUM)),5,2) + '/' + SubString(LTrim(Str(FILE_ID_NUM)),7,2) + '/' + SubString(LTrim(Str(FILE_ID_NUM)),1,4) As CONVERT_FILE_ID_NUM,FUNB_TOTAL_CREDITS,FUNB_TOTAL_DEBITS,CONVERT(VARCHAR(10),CREATED_DATETIME,101) AS CREATED_DATETIME,CREATED_BY FROM DD_BAI_FILE_SUMMARY WHERE BAI_FILE_DATETIME >= '" & DateAdd("d", -10, Format(Now, "MM/dd/yyyy") & " 00:00:00") & "' ORDER BY FILE_ID_NUM DESC"
mrsSummaryRecs.Open sSql, gcnDDS
sdgSummaryRecs.Rebind

'adcInvalid.ConnectionString = gobjLoginInfo.ConnectString
'adcInvalid.RecordSource = "SELECT DD_NUM,FUNB_INCOME_SRC_TYPE,TOT_FUNB_BENEFIT_AMT,DR_CR_FLAG,DECEASED_IND FROM DD_INVALID_REC WHERE RECORD_STATUS = 'A' ORDER BY DD_NUM"
'adcInvalid.Refresh
Set mrsInValid = New ADODB.Recordset
sSql = "SELECT DD_NUM,FUNB_INCOME_SRC_TYPE,TOT_FUNB_BENEFIT_AMT,DR_CR_FLAG,DECEASED_IND FROM DD_INVALID_REC WHERE RECORD_STATUS = 'A' ORDER BY DD_NUM"
mrsInValid.Open sSql, gcnDDS
sdgInvalid.Rebind

RefreshGrids

End Sub

Private Sub RefreshGrids()

Dim rsBalance As New ADODB.Recordset
Dim rsInvalid As New ADODB.Recordset
Dim rsPosted As New ADODB.Recordset
Dim rsLedger As New ADODB.Recordset
Dim sSql As String
Dim iX As Integer
Dim dTotBenefit As Currency
Dim dtLastBalance As Date
Dim dTotHidden As Currency

mdInvalidTotalSelected = 0
mdBegBal = 0
mdCreditsPosted = 0
mdEndingBal = 0
mdInvalidTotal = 0
mdDeceasedTotal = 0
mdAdjEnd = 0
mdDifference = 0
mdLedgerBal = 0
mdAdjustments = 0

'Open the Balance table and move to the latest record
sSql = "SELECT * FROM DD_BALANCE ORDER BY BALANCE_ID DESC "
rsBalance.CursorLocation = adUseServer
rsBalance.Open sSql, gcnDDS, adOpenForwardOnly

If rsBalance.EOF Then
    mdCarryOverBal = 0
    dtLastBalance = DateAdd("yyyy", -1, Now)
Else
    mdCarryOverBal = rsBalance!TOT_CR_DR_PRE_EDIT_RPT + rsBalance!TOT_CR_DR_DECEASED_EXCEPT
    dtLastBalance = rsBalance!CREATED_DATETIME
End If

'Fill values for invalid List box
sSql = "SELECT INVALID_RECORD_ID,DD_NUM,FUNB_INCOME_SRC_TYPE,TOT_FUNB_BENEFIT_AMT,DR_CR_FLAG, DECEASED_IND FROM DD_INVALID_REC WHERE RECORD_STATUS = 'A'"
rsInvalid.Open sSql, gcnDDS, adOpenForwardOnly


'Sum all hidden records since last balance
sSql = "SELECT SUM(TOT_FUNB_BENEFIT_AMT) As 'TotBenefitSummed' FROM DD_INVALID_REC WHERE DR_CR_FLAG = 'DR' AND LAST_MOD_DATETIME >= '" & dtLastBalance & "' AND RECORD_STATUS = 'I'"
rsPosted.Open sSql, gcnDDS, adOpenForwardOnly
If Not rsPosted.EOF Then
    If IsNull(rsPosted!TotBenefitSummed) Then
        dTotHidden = 0
    Else
        dTotHidden = (-1 * rsPosted!TotBenefitSummed)
    End If
End If
rsPosted.Close

sSql = "SELECT SUM(TOT_FUNB_BENEFIT_AMT) As 'TotBenefitSummed' FROM DD_INVALID_REC WHERE DR_CR_FLAG = 'CR' AND LAST_MOD_DATETIME >= '" & dtLastBalance & "' AND RECORD_STATUS = 'I'"
rsPosted.Open sSql, gcnDDS, adOpenForwardOnly
If Not rsPosted.EOF Then
    If Not IsNull(rsPosted!TotBenefitSummed) Then
        dTotHidden = dTotHidden + rsPosted!TotBenefitSummed
    End If
End If
rsPosted.Close


'Sum the Debit total of all the benefits that have been posted to the history table since the balance record was created
sSql = "SELECT SUM(TOT_FUNB_BENEFIT_AMT) As 'TotBenefitSummed' FROM DD_POSTING_HISTORY WHERE DR_CR_FLAG = 'DR' AND POSTED_DATETIME > '" & dtLastBalance & "'"
rsPosted.Open sSql, gcnDDS, adOpenForwardOnly
If Not rsPosted.EOF Then
    If IsNull(rsPosted!TotBenefitSummed) Then
        dTotBenefit = 0
    Else
        dTotBenefit = (-1 * rsPosted!TotBenefitSummed)
    End If
End If
rsPosted.Close

'Sum the Credit total of all the benefits that have been posted to the history table since the balance record was created
sSql = "SELECT SUM(TOT_FUNB_BENEFIT_AMT) As 'TotBenefitSummed' FROM DD_POSTING_HISTORY WHERE DR_CR_FLAG = 'CR' AND POSTED_DATETIME > '" & dtLastBalance & "'"
rsPosted.Open sSql, gcnDDS, adOpenForwardOnly
If Not rsPosted.EOF Then
    If Not IsNull(rsPosted!TotBenefitSummed) Then
        dTotBenefit = dTotBenefit + rsPosted!TotBenefitSummed
    End If
End If
 
rsPosted.Close

'Get the ledger balance of the most recent BAI file
sSql = "SELECT LEDGER_BAL FROM DD_BAI_FILE_SUMMARY ORDER BY FILE_ID_NUM DESC"
rsLedger.CursorLocation = adUseServer
rsLedger.Open sSql, gcnDDS, adOpenForwardOnly

'Remove All rows from the invalid grid

'adcInvalid.Recordset.Requery

sdgSummary.RemoveAll

mdInvalidTotal = 0

With rsInvalid
    Do Until .EOF
        
        If !DECEASED_IND = "N" Then
            If !DR_CR_FLAG = "DR" Then
                mdInvalidTotal = mdInvalidTotal + (-1 * !TOT_FUNB_BENEFIT_AMT)
            Else
                mdInvalidTotal = mdInvalidTotal + !TOT_FUNB_BENEFIT_AMT
            End If
        Else
            If !DR_CR_FLAG = "DR" Then
                mdDeceasedTotal = mdDeceasedTotal + (-1 * !TOT_FUNB_BENEFIT_AMT)
            Else
                mdDeceasedTotal = mdDeceasedTotal + !TOT_FUNB_BENEFIT_AMT
            End If
        End If
        rsInvalid.MoveNext
    Loop
    
    If rsBalance.EOF Then
        If txtBegBalance = vbNullString Then
            mbBegBalanceNeeded = True
            cmdFinish.Enabled = False
            mdBegBal = 0
        Else
            If IsNumeric(txtBegBalance) Then
                mdBegBal = CDbl(txtBegBalance)
                If mbValidRecordsFound = False Then
                    cmdFinish.Enabled = True
                End If
            Else
                mdBegBal = 0
                mbBegBalanceNeeded = True
                cmdFinish.Enabled = False
            End If
        End If
    Else
        mbBegBalanceNeeded = False
        mdBegBal = rsBalance!ADJ_ENDING_BAL
    End If
    
    If rsLedger.EOF Then
        mdLedgerBal = 0
    Else
        mdLedgerBal = rsLedger!LEDGER_BAL
    End If
    
    sdgSummary.AddItem "Beginning Balance:" & vbTab & Format$(mdBegBal, "Currency")
    mdCreditsPosted = dTotBenefit
    sdgSummary.AddItem "- Carryover from previous balance:" & vbTab & Format$(mdCarryOverBal, "Currency")
    sdgSummary.AddItem "+ Total CR/DR Posted:" & vbTab & Format$(mdCreditsPosted, "Currency")
    sdgSummary.AddItem "+ Total CR/DR from Adjustments:" & vbTab & Format$(dTotHidden, "Currency")
    mdEndingBal = mdBegBal - mdCarryOverBal + mdCreditsPosted + dTotHidden
    sdgSummary.AddItem "= Ending Balance:" & vbTab & Format$(mdEndingBal, "Currency")
    sdgSummary.AddItem "+ Total CR/DR from Pre-Edit File:" & vbTab & Format$(mdInvalidTotal, "Currency")
    sdgSummary.AddItem "+ Total CR/DR from Deceased Patients:" & vbTab & Format$(mdDeceasedTotal, "Currency")
    mdAdjEnd = mdEndingBal + mdInvalidTotal + mdDeceasedTotal
    sdgSummary.AddItem "= Adjusted Ending Balance:" & vbTab & Format$(mdAdjEnd, "Currency")
    sdgSummary.AddItem "- FUNB Ledger Balance:" & vbTab & Format$(mdLedgerBal, "Currency")
    mdDifference = Round(mdAdjEnd - mdLedgerBal, 2)
    
    sdgSummary.AddItem "= Difference:" & vbTab & Format$(mdDifference, "Currency")
    mdInvalidTotalSelected = mdInvalidTotal + mdDeceasedTotal
End With

    'lblInvalidTotal.Caption = Format$(mdInvalidTotalSelected, "Currency")

mdAdjustments = dTotHidden

rsBalance.Close
Set rsBalance = Nothing
rsInvalid.Close
Set rsInvalid = Nothing
Set rsPosted = Nothing
rsLedger.Close
Set rsLedger = Nothing

End Sub




'Private Sub lsvInvalid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'
'    If ColumnHeader.Index = 3 Then
'        If lsvInvalid.SortKey = ColumnHeader.Index Then
'            If lsvInvalid.SortOrder = lvwAscending Then
'                lsvInvalid.SortOrder = lvwDescending
'            Else
'                lsvInvalid.SortOrder = lvwAscending
'            End If
'        Else
'            lsvInvalid.SortKey = ColumnHeader.Index
'            lsvInvalid.SortOrder = lvwAscending
'        End If
'    Else
'
'        If lsvInvalid.SortKey = ColumnHeader.Index - 1 Then
'            'Change the sort order
'            If lsvInvalid.SortOrder = lvwAscending Then
'                lsvInvalid.SortOrder = lvwDescending
'            Else
'                lsvInvalid.SortOrder = lvwAscending
'            End If
'        Else
'            lsvInvalid.SortKey = ColumnHeader.Index - 1
'            lsvInvalid.SortOrder = lvwAscending
'        End If
'    End If
'
'End Sub

'Private Sub lsvInvalid_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'If Item.Checked = False Then
'    mdInvalidTotalSelected = mdInvalidTotalSelected - CDbl(Item.ListSubItems(2))
'Else
'    mdInvalidTotalSelected = mdInvalidTotalSelected + CDbl(Item.ListSubItems(2))
'End If
'lblInvalidTotal = Format$(mdInvalidTotalSelected, "Currency")
'
'
'End Sub


Private Sub OutlookTitle1_IconClick()
    If cmdCancel.Enabled = True Then
        Unload Me
    End If
    
End Sub

Private Sub sdgInvalid_HeadClick(ByVal ColIndex As Integer)
'    Static PreEditColumnFlag As udtSortedColumnFlag
'    Dim sSql As String
'    Dim rsPreEdit As New ADODB.Recordset
'    Hourglass True
'    sSql = "SELECT DD_NUM,FUNB_INCOME_SRC_TYPE,TOT_FUNB_BENEFIT_AMT,DR_CR_FLAG,DECEASED_IND FROM DD_INVALID_REC WHERE RECORD_STATUS = 'A'"
'    sSql = sSql & " ORDER BY " & sdgInvalid.Columns(ColIndex).Name
'    If PreEditColumnFlag.ColIndex = ColIndex And PreEditColumnFlag.Ascending = True Then
'        sSql = sSql & " DESC"
'        PreEditColumnFlag.Ascending = False
'    Else
'        PreEditColumnFlag.Ascending = True
'    End If
'    PreEditColumnFlag.ColIndex = ColIndex
'    With sdgInvalid
'        .Redraw = False
'        rsPreEdit.Open sSql, gcnDDS, adOpenStatic
'        Set .DataSource = rsPreEdit
'        .Refresh
'        .Redraw = True
'    End With
'
'    Hourglass False
'
End Sub

Private Sub sdgInvalid_UnboundReadData(ByVal RowBuf As SSDataWidgets_B_OLEDB.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
On Error GoTo sdgInvalidReadErr

Dim iX, R, i, J As Integer
Dim ct As Integer

If mrsInValid Is Nothing Then
    Exit Sub
End If
    
If mrsInValid.RecordCount = 0 Then
    Exit Sub
End If

ct = mrsInValid.Fields.Count - 1

    If IsNull(StartLocation) Then
    If ReadPriorRows Then
        mrsInValid.MoveLast
    Else
        mrsInValid.MoveFirst
    End If

Else
    mrsInValid.Bookmark = StartLocation
    If ReadPriorRows Then
        mrsInValid.MovePrevious
    Else
        mrsInValid.MoveNext
    End If

End If
    
For i = 0 To RowBuf.RowCount
    If mrsInValid.BOF Or mrsInValid.EOF Then Exit For

    For J = 0 To ct
        RowBuf.Value(i, J) = mrsInValid(J)
    Next J
    
    RowBuf.Bookmark(i) = mrsInValid.Bookmark

If ReadPriorRows Then
        mrsInValid.MovePrevious
    Else
        mrsInValid.MoveNext
    End If

    R = R + 1

Next i

RowBuf.RowCount = R

Exit Sub

sdgInvalidReadErr:
End Sub


Private Sub sdgSummaryRecs_UnboundReadData(ByVal RowBuf As SSDataWidgets_B_OLEDB.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)

On Error GoTo sdgSummaryRecsReadErr

Dim iX, R, i, J As Integer
Dim ct As Integer

If mrsSummaryRecs Is Nothing Then
    Exit Sub
End If
    
If mrsSummaryRecs.RecordCount = 0 Then
    Exit Sub
End If

ct = mrsSummaryRecs.Fields.Count - 1

    If IsNull(StartLocation) Then
    If ReadPriorRows Then
        mrsSummaryRecs.MoveLast
    Else
        mrsSummaryRecs.MoveFirst
    End If

Else
    mrsSummaryRecs.Bookmark = StartLocation
    If ReadPriorRows Then
        mrsSummaryRecs.MovePrevious
    Else
        mrsSummaryRecs.MoveNext
    End If

End If
    
For i = 0 To RowBuf.RowCount - 1
    If mrsSummaryRecs.BOF Or mrsSummaryRecs.EOF Then Exit For

    For J = 0 To ct
        RowBuf.Value(i, J) = mrsSummaryRecs(J)
    Next J
    
    RowBuf.Bookmark(i) = mrsSummaryRecs.Bookmark

If ReadPriorRows Then
        mrsSummaryRecs.MovePrevious
    Else
        mrsSummaryRecs.MoveNext
    End If

    R = R + 1

Next i

RowBuf.RowCount = R

Exit Sub

sdgSummaryRecsReadErr:

End Sub

Private Sub txtBegBalance_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If IsNumeric(txtBegBalance) Then
        cmdBalanceOK.Enabled = True
    Else
        cmdBalanceOK.Enabled = False
    End If
    

End Sub


