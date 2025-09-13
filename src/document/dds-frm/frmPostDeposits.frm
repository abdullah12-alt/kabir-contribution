VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PicClp32.Ocx"
Object = "{8CD222DF-7752-11D3-9D1E-00105A19BCF2}#1.0#0"; "OAOTBar.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPostDeposits 
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9900
   ControlBox      =   0   'False
   Icon            =   "frmPostDeposits.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   9900
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5760
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin OAOTitleBar.OutlookTitleBar OutlookTitle1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   979
      ForeColor       =   16777215
      Caption         =   "Post Transactions"
   End
   Begin VB.Timer TimeoutTimer 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   5280
      Top             =   600
   End
   Begin MSComctlLib.ImageList imlPost 
      Left            =   4200
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPostDeposits.frx":000C
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPostDeposits.frx":045E
            Key             =   "NonFix"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPostDeposits.frx":08B0
            Key             =   "Error"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "&Post"
      Height          =   465
      Left            =   5790
      TabIndex        =   3
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   7725
      TabIndex        =   2
      Top             =   6600
      Width           =   1815
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   3180
      Top             =   7080
      _ExtentX        =   6033
      _ExtentY        =   3810
      _Version        =   393216
      Rows            =   2
      Cols            =   3
      Picture         =   "frmPostDeposits.frx":0D02
   End
   Begin VB.Timer AnimationTimer 
      Interval        =   200
      Left            =   4800
      Top             =   600
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   3870
      TabIndex        =   0
      Top             =   1080
      Width           =   5715
      Begin MSComctlLib.ListView lvwStatus 
         Height          =   2685
         Left            =   240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2010
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4736
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imlPost"
         SmallIcons      =   "imlPost"
         ColHdrIcons     =   "imlPost"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ProgressBar proStatus 
         Height          =   390
         Left            =   285
         TabIndex        =   7
         Top             =   2295
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblPerforming 
         Height          =   255
         Left            =   285
         TabIndex        =   6
         Top             =   1995
         Width           =   4455
      End
      Begin VB.Image Image1 
         Height          =   1395
         Left            =   240
         Top             =   435
         Width           =   1515
      End
      Begin VB.Label lblEditStatus 
         Caption         =   "Press the post button below to start the posting process."
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
         Left            =   1830
         TabIndex        =   1
         Top             =   555
         Width           =   3690
      End
   End
   Begin VB.Image imgImage 
      BorderStyle     =   1  'Fixed Single
      Height          =   6375
      Left            =   120
      Picture         =   "frmPostDeposits.frx":9194
      Stretch         =   -1  'True
      Top             =   750
      Width           =   3510
   End
End
Attribute VB_Name = "frmPostDeposits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' ********************************************************************************
' * Description:
' *
' *
' *
' *
' *
' *
' * Revisions:
' *  8/9/99 fml Added comments.
' *
' *
' ********************************************************************************


' Mod CONSTANTS
Private Const MODULE As String = "frmPostDeposits"
Private Const TIME_OUT_SEC As Integer = 60
' Mod ENUMS
Private Type udtPatientName
    LastName As String
    FirstName As String
    MiddleName As String
    Initial As String
End Type


' Mod TYPES


' Mod DECLARES


' Mod VARIABLES

Dim mbConnected As Boolean
Dim mbSent As Boolean
Dim mbDataArrived As Boolean

Public Toggle As Integer
Public iPic As Integer

Private sHublinkName As String
Private sHublinkPort As String
Private lLastAckTick As Long
'Flag to indicate that the PostDeposits is running or not
Private bPostDepositsRunning As Boolean
'Flag to indicate that the Socket is still reading or not
Private bSocketReadRunning As Boolean
Private oAckParser As New clsHL7AckParser
Private mcmdPFS As New ADODB.Command
Private mcmdValidToInvalid As New ADODB.Command
Private msTransGroupNum As String
Dim mdtStartDateTime As Date


Private Sub GetTransGroupNum(sTransGroupNum As String)
    Dim cmd As New ADODB.Command

On Error GoTo GetTransGroupNumErr

    Set cmd.ActiveConnection = gcnDDS
    gStoredProcs("up_i_GetTransGroupNum").GetStoredProcCommand cmd
    cmd.Execute
    sTransGroupNum = cmd.Parameters("sTransGroupNum_OUTPUT")


Xit:
    Set cmd = Nothing
    Exit Sub

GetTransGroupNumErr:
    ShowUnexpectedError MODULE + " GetTransGroupNum", Err
    Resume Xit


End Sub


Private Sub AnimationTimer_Timer()

    Static iPercent As Integer
    If Toggle = 1 Then
        iPic = iPic + 1
        If iPic = 6 Then
            iPic = 0
        End If
        Image1.Picture = PictureClip1.GraphicCell(iPic)
    End If
End Sub


Private Sub cmdCancel_Click()
    Toggle = 0
    Unload Me

End Sub

Private Sub cmdPost_Click()
    On Error Resume Next
    giProcess = VALIDATE_AND_POST
    cmdPost.Enabled = False
    PostDeposits
    
End Sub

Private Sub Form_Load()

On Error GoTo Form_LoadErr

    Set OutlookTitle1.Picture = fMainForm.imlToolbarIcons.ListImages("Post Transactions").Picture
    Image1.Picture = PictureClip1.GraphicCell(5)
    PurgeAckFiles
'3/6/2015 - AS will no longer use the socket due to security
'    InitSocket
    InitStatusView
    InitDBParams
    CleanUpValidRec ("Unknown")
Xit:
    Exit Sub

Form_LoadErr:
    ShowError MODULE + "Form_Load", Err
    Unload Me
    Resume Xit


End Sub

Public Sub PurgeAckFiles()
'********************************************************************************
'* Name: PurgeAckFiles
'*
'* Description: Delete Acknowledgement files that are over 1 month old
'* Parameters:
'* Created: 9/21/99 12:21:21 PM
'********************************************************************************
    'AS - 2/16/2014 - Removed all references to FileSystemObject
    'Dim oFileSys As FileSystemObject
    Dim oFileSys As FSO
    Dim oAckFiles As String
    Dim oSepFiles As Variant
    Dim oFile As Variant
    'Dim oAckFiles As Files
    'Dim oAckFile As File
    Dim dtNow As Date

On Error GoTo PurgeAckFilesErr

    dtNow = Now
    If DirExists(gsDataPath & "\ACK") Then
        Set oFileSys = New FSO
        oAckFiles = oFileSys.GetAllFilesInFolder(gsDataPath & "\ACK")
        oSepFiles = Split(oAckFiles, "|")
        
        For Each oFile In oSepFiles
            'If DateDiff("m", oAckFile.DateCreated, Now) >= 1 Then
                oFileSys.DeleteAFile CStr(oFile)
            'End If
        Next oFile
        
        Set oFileSys = Nothing
    End If
    

Xit:
    Exit Sub

PurgeAckFilesErr:
    If Err.Number = 76 Then
        Resume Xit
    Else
        ShowUnexpectedError MODULE + " PurgeAckFiles", Err
        Resume Xit
    End If

End Sub
Private Sub SendAllValidtoHistory()
'*************************************************
'* This routine is used only for testing purposes
'*************************************************
On Error GoTo SendAllValidToHistoryErr

Dim rsPATrans As New ADODB.Recordset
Dim cmdHistory As New ADODB.Command
Dim sSql As String
Dim ix As Integer
    
    Set cmdHistory.ActiveConnection = gcnDDS
    If gStoredProcs("up_id_Move_Valid_to_Posting").GetStoredProcCommand(cmdHistory) = False Then
        MsgBox Error
    End If
        
    sSql = "SELECT * FROM DD_VALID_REC"
    rsPATrans.Open sSql, gcnDDS, adOpenForwardOnly
    Do Until rsPATrans.EOF
        With cmdHistory
            For ix = 0 To .Parameters.Count - 1
                .Parameters(ix) = Null
            Next ix
            'Message Control ID is the HL7 equivalent of Valid Record ID
            .Parameters("dValidRecordID").value = rsPATrans!VALID_RECORD_ID
            .Parameters("sUserID").value = gobjLoginInfo.UserId
        End With
        cmdHistory.Execute
        rsPATrans.MoveNext
    Loop

Exit Sub

SendAllValidToHistoryErr:
    MsgBox Error
    Resume

End Sub
Private Sub CleanUpValidRec(sPAPostingStatus As String)
'********************************************************************************
'* Name: CleanUpValidRec
'*
'* Description: Move the partial posted records to invalid rec table
'* Parameters:
'* Created: 9/21/99 12:23:01 PM
'********************************************************************************
    Dim rs As New ADODB.Recordset
    Dim sSql As String


On Error GoTo CleanUpValidRecErr

    sSql = "SELECT * FROM DD_VALID_REC WHERE NOT(SENT_FOR_POSTING_DATETIME IS NULL)"
    rs.Open sSql, gcnDDS, adOpenForwardOnly, adLockOptimistic
    If Not (rs.EOF = True And rs.BOF = True) Then
        MsgBox "There are some partially posted records in valid record table from previous posting attempt. These records will be moved to Pre-Edit file.", vbExclamation
        Do Until rs.EOF
            'With mcmdValidToInvalid
            '    .Parameters("dValidRecordID").Value = rs.Fields("VALID_RECORD_ID")
            '    .Parameters("sUserID").Value = gobjLoginInfo.UserId
            '    .Parameters("sPAPostingStatus").Value = sPAPostingStatus
            '    .Parameters("dPAErrCode").Value = Null
            '    .Parameters("sPFPostingStatus").Value = "N/A"
            'End With
            'mcmdValidToInvalid.Execute
            rs.MoveNext
        Loop
    End If


Xit:
    Exit Sub

CleanUpValidRecErr:
    ShowUnexpectedError MODULE + "CleanUpValidRec", Err
    Resume Xit


End Sub
Private Sub InitDBParams()

On Error GoTo InitDBParamsErr

    Set mcmdPFS.ActiveConnection = gcnDDS
    gStoredProcs("up_id_Post_PFS_Trans").GetStoredProcCommand mcmdPFS
    Set mcmdValidToInvalid.ActiveConnection = gcnDDS
    gStoredProcs("up_id_Move_Valid_to_Invalid").GetStoredProcCommand mcmdValidToInvalid
        

Xit:
    Exit Sub

InitDBParamsErr:
    ShowUnexpectedError MODULE + " InitDBParams", Err
    Resume Xit


End Sub

Private Sub InitStatusView()
    Dim Item As ListItem
    With lvwStatus
        .ColumnHeaders.Add , "Entity Name", "Entity Name", 2000
        .ColumnHeaders.Add , "Count", "Count"
        .ColumnHeaders("Count").Alignment = lvwColumnRight
        .View = lvwReport
        Set Item = .ListItems.Add(, "Total Records", "Total Records", , "Fix")
        Set Item = .ListItems.Add(, "Total PA Trans", "Total PA Trans", , "Fix")
        Set Item = .ListItems.Add(, "Total PF Trans", "Total PF Trans", , "Fix")
        Set Item = .ListItems.Add(, "PA Sent", "PA Trans Sent", , "NonFix")
        Set Item = .ListItems.Add(, "PA Ack", "PA Trans Ack", , "NonFix")
        Set Item = .ListItems.Add(, "PF Post", "PF Trans Posted", , "NonFix")
        Set Item = .ListItems.Add(, "PA Error", "PA Posting Error", , "Error")
        Set Item = .ListItems.Add(, "PF Error", "PF Posting Error", , "Error")
        
    End With
End Sub


Public Sub ShowStatus(iPercent As Integer)
    
    proStatus.value = iPercent
    DoEvents
    
End Sub

Private Function SendToPFSBatch() As Boolean
    Dim rsPFTrans As New ADODB.Recordset
    'Dim cmdHistory As New ADODB.Command
    Dim sSql As String
    Dim dtDate As Date
    Dim lReturnValue As Long
    Dim sPFPostingStatus As String

On Error GoTo SendToPFSBatchErr

    lblEditStatus.Caption = "Posting transactions to Affinity." & vbCrLf & "Please wait until process is complete..."
    SendToPFSBatch = False
    
    'Set cmdHistory.ActiveConnection = gcnDDS
    'If gStoredProcs("up_i_Posting_History").GetStoredProcCommand(cmdHistory) = False Then
    '    Err.Raise INVALID_POSTING_HISTORY_STORED_PROC_FAILED
    'End If
        
    sSql = "SELECT DD_VALID_REC.VALID_RECORD_ID FROM DD_VALID_REC WHERE DD_VALID_REC.SENT_FOR_POSTING_DATETIME IS NULL AND PF_DISTRIBUTION_AMT > 0 AND PA_DISTRIBUTION_AMT = 0"
    rsPFTrans.Open sSql, gcnDDS, adOpenForwardOnly
    Do Until rsPFTrans.EOF
'        ShowStatus CInt(rsValid.AbsolutePosition / rsValid.RecordCount * 100)
        With mcmdPFS
            .Parameters("dValidRecordID").value = rsPFTrans.Fields("VALID_RECORD_ID")
            .Parameters("sTransGroupNum").value = msTransGroupNum
            .Parameters("sUserID").value = "lvmantooth"    'gobjLoginInfo.UserId
            .Parameters("sPAPostingStatus").value = "N/A"
        End With
        mcmdPFS.Execute
        lReturnValue = mcmdPFS.Parameters("RETURN_VALUE")
        'If successful update the status view
        If lReturnValue = 0 Then
            Debug.Print "Posting result = " & lReturnValue
            sPFPostingStatus = mcmdPFS.Parameters("sPFPostingStatus_OUTPUT")
            Select Case sPFPostingStatus
                Case "Posted"
                    lvwStatus.ListItems("PF Post").SubItems(1) = lvwStatus.ListItems("PF Post").SubItems(1) + 1
                Case "Post Fails"
                    lvwStatus.ListItems("PF Error").SubItems(1) = lvwStatus.ListItems("PF Error").SubItems(1) + 1
            End Select
        Else
            lvwStatus.ListItems("PF Error").SubItems(1) = lvwStatus.ListItems("PF Error").SubItems(1) + 1
        End If
        DoEvents
        rsPFTrans.MoveNext
    Loop
    rsPFTrans.Close
    Set rsPFTrans = Nothing
    
    lblEditStatus.Caption = "Posting of PFS only transactions is complete. Waiting for the PA Acknowledgements ..."
    
    

Xit:
    Exit Function

SendToPFSBatchErr:
    ShowUnexpectedError MODULE + "SendToPFSBatch", Err
    Resume Xit


End Function

Private Function SendToAffinity() As Boolean
    Dim oP03Msg As New clsHL7P03Message
    Dim sP03Msg As String
    Dim rsPATrans As New ADODB.Recordset
    Dim rsConfig As New ADODB.Recordset
    'Dim cmdHistory As New ADODB.Command
    Dim sSql As String
    Dim dtDate As Date
    Dim cmd As New ADODB.Command
    Dim lWait As Long
    Dim sMsgRcvd As String
    Dim sMsgControlID As String
    Dim iAck As AcknowledgementCode
    Dim sErrCode As String
    Dim lReturnValue As Long
    Dim sPFPostingStatus As String
    Dim ix As Integer
    Dim sTestProd As String
    Dim bSkipAffinityPost As Boolean
    Dim sErrMsg As String
    Dim tf As clsTextFile
    Dim sFileName As String
    Dim s7ZName As String
    Dim s7ZSearch As String
    
On Error GoTo SendToAffinityErr

    sTestProd = ReadIniFile(App.Path & "\" & App.EXEName & ".ini", "Startup", "TestOrProduction")
    If sTestProd = "T" Or sTestProd = "P" Then
        'Skip
    Else
        sTestProd = "T"
    End If
    
    sSql = "SELECT FT1_INSURANCE_CODE, PATCODE_ENTERING_AREA FROM DD_CONFIG_INFO"
    rsConfig.Open sSql, gcnDDS, adOpenForwardOnly
    If rsConfig.EOF Then
        SendToAffinity = False
        GoTo Xit
    End If
    
    mbSent = False
    mbDataArrived = False
    Set cmd.ActiveConnection = gcnDDS
    gStoredProcs("up_u_Sent_for_Posting").GetStoredProcCommand cmd
    
    lblEditStatus.Caption = "Posting transactions to Affinity." & vbCrLf & "Please wait until process is complete..."
    SendToAffinity = False
    
    'Set cmdHistory.ActiveConnection = gcnDDS
    'If gStoredProcs("up_i_Posting_History").GetStoredProcCommand(cmdHistory) = False Then
    '    Err.Raise INVALID_POSTING_HISTORY_STORED_PROC_FAILED
    'End If
        
    sSql = "SELECT DD_VALID_REC.*, DD_INCOME_SOURCE_TYPE.PA_PMT_CODE FROM DD_VALID_REC, DD_INCOME_SOURCE_TYPE WHERE DD_VALID_REC.INCOME_SOURCE_TYPE_ID = DD_INCOME_SOURCE_TYPE.INCOME_SOURCE_TYPE_ID AND PA_DISTRIBUTION_AMT > 0"
    rsPATrans.Open sSql, gcnDDS, adOpenForwardOnly
    dtDate = GetDateTime
    Do Until rsPATrans.EOF
'        ShowStatus CInt(rsValid.AbsolutePosition / rsValid.RecordCount * 100)
        If IsNull(rsPATrans!SENT_FOR_POSTING_DATETIME) = False And rsPATrans.Fields!PA_DISTRIBUTION_AMT > 0 Then
            sErrMsg = "The following record was previously posted to Affinity but due to an error not to Personal Funds.  Should I post only to Personal Funds?" & vbCrLf
            sErrMsg = sErrMsg & "DDNumber: " & rsPATrans!DD_NUM & vbCrLf
            sErrMsg = sErrMsg & "MRUN: " & rsPATrans!MEDICAL_RECORD_NUM & vbCrLf
            sErrMsg = sErrMsg & "Name: " & rsPATrans!PATIENT_NAME & vbCrLf
            sErrMsg = sErrMsg & "PA Amount: " & rsPATrans!PA_DISTRIBUTION_AMT
            sErrMsg = sErrMsg & "PF Amount: " & rsPATrans!PF_DISTRIBUTION_AMT
            bSkipAffinityPost = CBool(MsgBox(sErrMsg, vbQuestion + vbYesNo + vbDefaultButton1) - vbNo)
            sMsgControlID = rsPATrans!VALID_RECORD_ID
        Else
            bSkipAffinityPost = False
        End If
        
        mbSent = False
        mbDataArrived = False
        sMsgControlID = rsPATrans.Fields("VALID_RECORD_ID")
        With oP03Msg
            .SendingApplication = "DDS"
            .SendingFacility = ""
            .ReceivingApplication = ""
            .ReceivingFacility = rsPATrans.Fields("INSTITUTION_CODE")
            .MessageControlID = rsPATrans.Fields("VALID_RECORD_ID")
            .RecordedDateTime = dtDate
            .PatientName = .GetName(rsPATrans.Fields("PATIENT_NAME"))
            .PatientIDInternal = Format$(rsPATrans.Fields("MEDICAL_RECORD_NUM"), "0000000")
            .PatientAccountNumber = Format$(rsPATrans.Fields("AFFINITY_ACCT_NUM"), "00000000")
            'T for training, P for production, D for debugging
            .ProcessingID = sTestProd
            .TransactionDate = dtDate
            .TransactionType = "PY"
            .TransactionCode = rsPATrans.Fields("PA_PMT_CODE")
            .TransactionQuantity = "1"
            .TransactionAmountExtended = rsPATrans.Fields("PA_DISTRIBUTION_AMT")
            .DepartmentCode = Trim$(rsConfig!PATCODE_ENTERING_AREA)
            .InsurancePlanID = Trim$(rsConfig!FT1_INSURANCE_CODE)
        End With
        sP03Msg = oP03Msg.GetString
        'Send the Hl7 Message to Hublink and wait until the message is sent completely
        If bSkipAffinityPost = True Then
            mbSent = True
        Else
            'Create a file to send
            Set tf = New clsTextFile
            sFileName = gsDataPath & "\ToPost\" & rsPATrans.Fields("INSTITUTION_CODE") & "_" & rsPATrans.Fields("VALID_RECORD_ID") & ".txt"
            If tf.OpenFile(sFileName, OUTPUT_NEW) = False Then
                PostFinished
                CleanUpValidRec ("Time Out")
                SendToAffinity = False
            Else
                mbSent = True
                tf.WriteLine sP03Msg
                tf.CloseFile
                
                'If SecureSend("uload", sFileName, "/sin/dds") = False Then
                '    MsgBox "Message could not be sent to Affinity", vbCritical
                '    PostFinished
                '    CleanUpValidRec ("Time Out")
                '    SendToAffinity = False
                '    Exit Function
                'End If
                
            End If
             '3/6/2015 - Will no longer send the transactions directly to Affinity.  The files will be created and send to Affinity Server first
'            lWait = 0
'            Winsock1.SendData sP03Msg
'            Do Until mbSent = True
'                 lWait = lWait + 1
'                DoEvents
'                If lWait > 9000000 Then
'                    MsgBox "Message could not be sent to Hublink", vbCritical
'                    Winsock1.Close
'                    PostFinished
'                    CleanUpValidRec ("Time Out")
'                    SendToAffinity = False
'                    Exit Function
'                End If
'            Loop
        End If
        
        'Delete the valid record from the table
        cmd.Parameters("dValidRecordID") = rsPATrans.Fields("VALID_RECORD_ID")
        cmd.Execute
        lvwStatus.ListItems("PA Sent").SubItems(1) = lvwStatus.ListItems("PA Sent").SubItems(1) + 1
        DoEvents
        rsPATrans.MoveNext
    
        If bSkipAffinityPost = True Then
            mbDataArrived = True
        Else
            mbDataArrived = True
'3/9/2015 - We will no longer be sending transactions directly to cloverleaf
'            lWait = 0
'            Do Until mbDataArrived = True
'                lWait = lWait + 1
'                If lWait > 9000000 Then
'                    MsgBox "Acknowledgement for Hublink was not received", vbCritical
'                    Winsock1.Close
'                    PostFinished
'                    CleanUpValidRec ("Time Out")
'                    SendToAffinity = False
'                    Exit Function
'                End If
'                DoEvents
'            Loop
        End If
        On Error Resume Next
        'If rsPATrans.Fields("VALID_RECORD_ID") Mod 5 = 0 Then
        '    Sleep 1
        'End If
        On Error GoTo SendToAffinityErr
        
        If mbDataArrived = True Then
            If bSkipAffinityPost = True Then
                lvwStatus.ListItems("PA Ack").SubItems(1) = lvwStatus.ListItems("PA Ack").SubItems(1) + 1
                With mcmdPFS
                    For ix = 0 To .Parameters.Count - 1
                        .Parameters(ix) = Null
                    Next ix
                    'Message Control ID is the HL7 equivalent of Valid Record ID
                    .Parameters("dValidRecordID").value = CDbl(sMsgControlID)
                    .Parameters("sTransGroupNum").value = msTransGroupNum
                    .Parameters("sUserID").value = "lvmantooth"   'gobjLoginInfo.UserId
                    .Parameters("sPAPostingStatus").value = "Posted"
                End With
                mcmdPFS.Execute
                lReturnValue = mcmdPFS.Parameters("RETURN_VALUE")
                'If successful update the status view
                If lReturnValue = 0 Then
                    Debug.Print "Posting result = " & lReturnValue
                    sPFPostingStatus = mcmdPFS.Parameters("sPFPostingStatus_OUTPUT")
                    Select Case sPFPostingStatus
                        Case "Posted"
                            lvwStatus.ListItems("PF Post").SubItems(1) = lvwStatus.ListItems("PF Post").SubItems(1) + 1
                        Case "Post Fails"
                            lvwStatus.ListItems("PF Error").SubItems(1) = lvwStatus.ListItems("PF Error").SubItems(1) + 1
                        Case "N/A"
                            'No PF posting for this record, do nothing for now
                    End Select
                Else
                'Not succesful
                    lvwStatus.ListItems("PF Error").SubItems(1) = lvwStatus.ListItems("PF Error").SubItems(1) + 1
                End If
            Else
                '3/10/2015 - Eliminating the acknowledgement
                'Winsock1.GetData sMsgRcvd, vbString
                'oAckParser.Message = sMsgRcvd
                'If oAckParser.MessageValid(sMsgControlID, iAck, sErrCode) Then
                If 1 = 1 Then
                iAck = APPLICATION_ACCEPT
                    If iAck = APPLICATION_ACCEPT Then
                    'PA Posting successful
                        lvwStatus.ListItems("PA Ack").SubItems(1) = lvwStatus.ListItems("PA Ack").SubItems(1) + 1
                        With mcmdPFS
                            For ix = 0 To .Parameters.Count - 1
                                .Parameters(ix) = Null
                            Next ix
                            'Message Control ID is the HL7 equivalent of Valid Record ID
                            .Parameters("dValidRecordID").value = CDbl(sMsgControlID)
                            .Parameters("sTransGroupNum").value = msTransGroupNum
                            .Parameters("sUserID").value = "lvmantooth"    'gobjLoginInfo.UserId
                            .Parameters("sPAPostingStatus").value = "Posted"
                        End With
                        mcmdPFS.Execute
                        lReturnValue = mcmdPFS.Parameters("RETURN_VALUE")
                        'If successful update the status view
                        If lReturnValue = 0 Then
                            Debug.Print "Posting result = " & lReturnValue
                            sPFPostingStatus = mcmdPFS.Parameters("sPFPostingStatus_OUTPUT")
                            Select Case sPFPostingStatus
                                Case "Posted"
                                    lvwStatus.ListItems("PF Post").SubItems(1) = lvwStatus.ListItems("PF Post").SubItems(1) + 1
                                Case "Post Fails"
                                    lvwStatus.ListItems("PF Error").SubItems(1) = lvwStatus.ListItems("PF Error").SubItems(1) + 1
                                Case "N/A"
                                    'No PF posting for this record, do nothing for now
                            End Select
                        Else
                        'Not succesful
                            lvwStatus.ListItems("PF Error").SubItems(1) = lvwStatus.ListItems("PF Error").SubItems(1) + 1
                        End If
                    ElseIf iAck = APPLICATION_ERROR Then
                    'PA Posting fails move the record to invalid record table
                        lvwStatus.ListItems("PA Error").SubItems(1) = lvwStatus.ListItems("PA Error").SubItems(1) + 1
                        With mcmdValidToInvalid
                            .Parameters("dValidRecordID").value = CDbl(sMsgControlID)
                            .Parameters("sUserID").value = "lvmantooth"      'gobjLoginInfo.UserId
                            .Parameters("sPAPostingStatus").value = "Post Fails"
                            .Parameters("dPAErrCode").value = CInt(sErrCode)
                            .Parameters("sPFPostingStatus").value = "N/A"
                        End With
                        mcmdValidToInvalid.Execute
                    Else
                    'PA Posting rejected move the record to invalid record table
                        lvwStatus.ListItems("PA Error").SubItems(1) = lvwStatus.ListItems("PA Error").SubItems(1) + 1
                        With mcmdValidToInvalid
                            .Parameters("dValidRecordID").value = CDbl(sMsgControlID)
                            .Parameters("sUserID").value = "lvmantooth"    'gobjLoginInfo.UserId
                            .Parameters("sPAPostingStatus").value = "Rejected"
                            .Parameters("dPAErrCode").value = Null
                            .Parameters("sPFPostingStatus").value = "N/A"
                        End With
                        mcmdValidToInvalid.Execute
                    End If
                End If
                
                'Check to see if we receive all acknowledgements
                If lvwStatus.ListItems("PA Ack").SubItems(1) = lvwStatus.ListItems("Total PA Trans").SubItems(1) Then
                    'PA posting finished
                End If
            End If
        End If
    Loop
    
    lblEditStatus.Caption = "Sending Affinity transactions..."
    
    s7ZSearch = gsDataPath & "\ToPost\*.txt"
    s7ZName = gsDataPath & "\ToPost\MSG" & Format(Now(), "MMddyyyy") & ".7z"
    If ZipFiles(s7ZName, s7ZSearch) = False Then
        MsgBox "Error Zipping Files please contact Direct Deposit Support"
    Else
        If SecureSend("uload", gsDataPath & "\ToPost\*.7z", "/sin/dds") = False Then
            MsgBox "Messages could not be sent to Affinity", vbCritical
        End If
    End If
    
    lblEditStatus.Caption = "Sending of Affinity transactions is complete..."
    
    SendToAffinity = True
    

Xit:
On Error Resume Next
    rsPATrans.Close
    Set rsPATrans = Nothing
    rsConfig.Close
    Set rsConfig = Nothing
    Set oP03Msg = Nothing
    Set cmd = Nothing
    
    Exit Function

SendToAffinityErr:

    ShowUnexpectedError MODULE + " SendToAffinity", Err
    Resume Xit

End Function

Private Function InitSocket() As Boolean
On Error GoTo ErrorHandle
    Dim lWait As Long
    
    Winsock1.RemoteHost = ReadIniFile(App.Path & "\" & App.EXEName & ".ini", "Startup", "HublinkName")
    Winsock1.RemotePort = ReadIniFile(App.Path & "\" & App.EXEName & ".ini", "Startup", "HublinkPort")
    Winsock1.Connect
    Do Until mbConnected = True
        lWait = lWait + 1
        If lWait > 9000000 Then
            MsgBox "Connection to Hublink could not be established.", vbCritical
            Winsock1.Close
            cmdPost.Enabled = False
            InitSocket = False
            Exit Function
        End If
        DoEvents
    Loop
    
    InitSocket = True
    
    Exit Function
ErrorHandle:
    MsgBox Error, vbCritical
    InitSocket = False
    Exit Function

End Function


Private Sub Form_Unload(Cancel As Integer)
    Dim lRet As Long
    On Error Resume Next
    '3/6/2015 - AS No longer sending transaction directly
    'Winsock1.Close
    Set mcmdPFS = Nothing
    Set mcmdValidToInvalid = Nothing
End Sub

Private Sub OutlookTitle1_IconClick()
    If cmdCancel.Enabled = True Then
        Unload Me
    End If

End Sub


Public Sub PostDeposits()
    Dim lTotalRecords As Long
    Dim lTotalPATrans As Long
    Dim lTotalPFTrans As Long
    Dim rs As New ADODB.Recordset
    Dim sSql As String
    Dim fs As New FSO
    Dim sFileList As String
    Dim vSplit As Variant
    Dim ix As Long
On Error GoTo PostDepositsErr

    mdtStartDateTime = GetDateTime
    
    'AS - Create the To Post Folder if file does not exist
    If fs.FolderExists(gsDataPath & "\ToPost") = False Then
        fs.CreateAFolder gsDataPath & "\ToPost"
    End If
    'AS - 3/6/2015 Delete all messages in folder
    sFileList = fs.GetAllFilesInFolder(gsDataPath & "\ToPost")
    vSplit = Split(sFileList, "|")
    For ix = 1 To UBound(vSplit)
        If vSplit(ix) <> "" Then
             fs.DeleteAFile CStr(vSplit(ix))
        End If
    Next ix
    'AS - 3/6/2015 End of Changes
    
    'Find the total number of records that haven't been sent to Affinity
    sSql = "SELECT COUNT(VALID_RECORD_ID) AS 'Total Records' FROM DD_VALID_REC"
    rs.Open sSql, gcnDDS, adOpenStatic, adLockReadOnly
    lTotalRecords = rs.Fields("Total Records")
    If rs.State = adStateOpen Then
        rs.Close
    End If
    'No records, notify user
    If lTotalRecords = 0 Then
        Beep
        MsgBox "There are no records that need to be posted.", vbExclamation
        GoTo Xit
    End If
    
    If giProcess = VALIDATE_AND_POST Then
        lvwStatus.Visible = False
        'ValidateTransactionsNew
    End If
    
    lblPerforming = ""
    lvwStatus.Visible = True
   
    'Find the total number of records that haven't been sent to Affinity
    sSql = "SELECT COUNT(VALID_RECORD_ID) AS 'Total Records' FROM DD_VALID_REC"
    rs.Open sSql, gcnDDS, adOpenStatic, adLockReadOnly
    lTotalRecords = rs.Fields("Total Records")
    If rs.State = adStateOpen Then
        rs.Close
    End If
    'No records, notify user
    If lTotalRecords = 0 Then
        Beep
        MsgBox "There are no records that need to be posted.", vbExclamation
        GoTo Xit
    End If
    
    'Find the total number of PA Trans
    sSql = "SELECT COUNT(VALID_RECORD_ID) AS 'Total PA Trans' FROM DD_VALID_REC WHERE PA_DISTRIBUTION_AMT > 0"
    rs.Open sSql, gcnDDS, adOpenStatic, adLockReadOnly
    lTotalPATrans = rs.Fields("Total PA Trans")
    If rs.State = adStateOpen Then
        rs.Close
    End If
    'If there are PA trans, we are reading ACK from socket. Set the flag now.
    If lTotalPATrans > 0 Then
        bSocketReadRunning = True
    End If
    
    'Find the total number of PF Trans
    sSql = "SELECT COUNT(VALID_RECORD_ID) AS 'Total PF Trans' FROM DD_VALID_REC WHERE PF_DISTRIBUTION_AMT > 0"
    rs.Open sSql, gcnDDS, adOpenStatic, adLockReadOnly
    lTotalPFTrans = rs.Fields("Total PF Trans")
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    'Update the status view
    With lvwStatus
        .ListItems("Total Records").SubItems(1) = lTotalRecords
        .ListItems("Total PA Trans").SubItems(1) = lTotalPATrans
        .ListItems("Total PF Trans").SubItems(1) = lTotalPFTrans
        .ListItems("PA Sent").SubItems(1) = 0
        .ListItems("PA Ack").SubItems(1) = 0
        .ListItems("PF Post").SubItems(1) = 0
        .ListItems("PA Error").SubItems(1) = 0
        .ListItems("PF Error").SubItems(1) = 0
    End With
    
    cmdPost.Enabled = False
    cmdCancel.Enabled = False
    
    bPostDepositsRunning = True
    
    iPic = 0
    Toggle = 1
    lblEditStatus.Caption = "Preparing to post transactions." & vbCrLf & "Please wait until the post process is complete..."
    
    GetTransGroupNum msTransGroupNum
    
    If lTotalPATrans > 0 Then
        SendToAffinity
    End If
    If lTotalPFTrans > 0 Then
        SendToPFSBatch
    End If
    
    bPostDepositsRunning = False
    
    PostFinished
    
    PrintReports

    Winsock1.Close
    
    'If MsgBox("Would you like to verify the payments in Affinity?", vbQuestion + vbYesNo) = vbYes Then
    '    VerifyAffinity Format$(Now, "mm/dd/yyyy"), 0
    'End If


Xit:
    On Error Resume Next
    'Release the reference to the recordset
    Close
    Set rs = Nothing
    Winsock1.Close
    
    Exit Sub

PostDepositsErr:
    'ShowUnexpectedError MODULE + "PostDeposits", Err
    'Resume Xit
Resume

End Sub

Private Sub PostFinished()
    lblEditStatus.Caption = "Posting is complete"
    Toggle = 0
    cmdCancel.Enabled = True
End Sub
'Private Function HL7NameFormat(sName As String) As String
'
'    Dim vName As Variant
'    Dim vFirstMiddleName As Variant
'    sName = Trim(sName)
'    vName = Split(sName, ",")
'    'The name we got from affinity is not in standard format
'    'We'll treat the string before comma(,) as last name
'    'and the string after comma(,) as first name
'    HL7NameFormat.LastName = CStr(vName(0))
'    HL7NameFormat.FirstName = CStr(vName(1))
''    vFirstMiddleName = Split(vName(1), " ")
''    HL7NameFormat.FirstName = vFirstMiddleName(0)
''    Dim sMiddleName As String
''    sMiddleName = CStr(vFirstMiddleName(1))
''    If Len(sMiddleName) = 1 Then
''        HL7NameFormat.Initial = sMiddleName
''    Else
''        HL7NameFormat.MiddleName = sMiddleName
''    End If
'End Function

Private Sub TimeoutTimer_Timer()
    Dim lCurrentTick As Long
    Dim dTimeElapsed As Double
    dTimeElapsed = (GetTickCount - lLastAckTick) / 1000
    Debug.Print dTimeElapsed
    If dTimeElapsed > TIME_OUT_SEC And bPostDepositsRunning = False Then
        MsgBox "Posting process has been timed out."
        TimeoutTimer.Enabled = False
        PostFinished
        CleanUpValidRec ("Time Out")
    End If
End Sub
Private Sub PrintReports()

On Error GoTo PrintReportsErr

'Report Name: DD Detail Exception Report
ViewPrintReport 3, "", Format(Now, "MM/dd/yyyy")
'------------------------------------------------------------
Xit:
    Exit Sub
PrintReportsErr:
    Hourglass False
    MsgBox Error, vbInformation
    Resume Xit
End Sub

Public Function SecureSend(ByVal sUsername As String, ByVal sInFileName As String, sRemotePath) As Boolean
    Dim sLine As String
    Dim sRet As String
    Dim sText As String
    ChDrive "C:"
    ChDir "C:\putty"
    
    sLine = "pscp -l " & sUsername & " -pw " & "u?T2@4s9M" & " """ & sInFileName & """ " & sUsername & "@hes001.dhr.state.nc.us:" & sRemotePath
    fMainForm.msDosOutput = ""
    fMainForm.objDOS.CommandLine = sLine
    sRet = fMainForm.objDOS.ExecuteCommand
    DoEvents
    If (sRet <> "" And InStr(1, sRet, "100%") > 0) Then
        'MsgBox "File has been securely copied."
        SecureSend = True
    Else
       Exit Function
    End If
End Function




Private Sub Winsock1_Connect()
    mbConnected = True
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Debug.Print "Winsock Connection Request " & requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Debug.Print "Winsock Data Arrival Bytes received " & bytesTotal
    mbDataArrived = True

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    MsgBox Description, vbCritical, Source
    CancelDisplay = True
    Winsock1.Close
    cmdPost.Enabled = False
    
End Sub

Private Sub Winsock1_SendComplete()
    
    mbSent = True
    
End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Debug.Print "Bytes Sent " & bytesSent & "Bytes Remaining " & bytesRemaining
End Sub
