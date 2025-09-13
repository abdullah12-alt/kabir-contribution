VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PicClp32.Ocx"
Object = "{8CD222DF-7752-11D3-9D1E-00105A19BCF2}#1.0#0"; "OAOTBar.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLoadFUNB 
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10155
   ControlBox      =   0   'False
   Icon            =   "frmLoadFUNB.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9585
   ScaleWidth      =   10155
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aquire BAI Detail File"
      Height          =   435
      Left            =   4200
      TabIndex        =   6
      Top             =   6345
      Visible         =   0   'False
      Width           =   1680
   End
   Begin OAOTitleBar.OutlookTitleBar OutlookTitle1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   979
      ForeColor       =   16777215
      Caption         =   "Load FUNB Transaction File"
   End
   Begin VB.Timer AnimationTimer 
      Interval        =   200
      Left            =   3570
      Top             =   2625
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   3945
      Top             =   7050
      _ExtentX        =   7197
      _ExtentY        =   2699
      _Version        =   393216
      Cols            =   2
      Picture         =   "frmLoadFUNB.frx":000C
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   3780
      TabIndex        =   2
      Top             =   2025
      Width           =   5895
      Begin MSComctlLib.ProgressBar proStatus 
         Height          =   390
         Left            =   405
         TabIndex        =   4
         Top             =   2655
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image Image1 
         Height          =   1785
         Left            =   105
         Top             =   450
         Width           =   1995
      End
      Begin VB.Label lblEditStatus 
         Caption         =   $"frmLoadFUNB.frx":1457E
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
         Left            =   2385
         TabIndex        =   3
         Top             =   450
         Width           =   3360
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   7770
      TabIndex        =   1
      Top             =   6345
      Width           =   1605
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load and Validate"
      Height          =   435
      Left            =   6000
      TabIndex        =   0
      Top             =   6345
      Width           =   1680
   End
   Begin VB.Image imgImage 
      BorderStyle     =   1  'Fixed Single
      Height          =   6345
      Left            =   165
      Picture         =   "frmLoadFUNB.frx":1461C
      Top             =   795
      Width           =   3450
   End
End
Attribute VB_Name = "frmLoadFUNB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iPic As Integer
Dim Toggle As Integer

'FUNB Record Types
Const FILE_HDR_REC = "01"
Const GROUP_HDR_REC = "02"
Const ACCOUNT_HDR_REC = "03"
Const TRANS_DTL_REC = "16"
Const ACCOUNT_TRLR_REC = "49"
Const CONTINUE_REC = "88"
Const GROUP_TRLR_REC = "98"
Const FILE_TRLR_REC = "99"
Const MODULE As String = "Load FUNB Transaction File - "

Private Enum LoadErrors
    NO_HEADER_RECORD = 2500
    SENDER_ID_RCVD_INVALID = 2501
    RECEIVER_ID_RCVD_INVALID = 2502
    UNEXPECTED_FILE_HDR = 2503
    UNEXPECTED_GROUP_HDR = 2504
    UNEXPECTED_ACCOUNT_HDR = 2505
    NO_GROUP_REC_RECEIVED = 2506
    UNEXPECTED_TRANS_DTL = 2507
    UNEXPECTED_ACCOUNT_TRLR = 2508
    UNEXPECTED_GROUP_TRLR = 2509
    UNEXPECTED_FILE_TRLR_REC = 2510
    BAI_FILE_STORED_PROC_FAILED = 2511
    FUNB_SUMMARY_STORED_PROC_FAILED = 2512
    ACCOUNT_CREDITS_DONT_MATCH = 2513
    ACCOUNT_DEBITS_DONT_MATCH = 2514
    FILE_PROCESSED_ALREADY = 6014
    NO_VALID_BAI_FILES_RECEIVED = 6015
    COULD_NOT_GET_CONFIG_INFO = 6016
    ERROR_CREATING_WORK_RECORD = 6017
    WORK_FILE_STORED_PROC_FAILED = 6018
    NO_FILE_TRAILER_RECEIVED = 6019
    DEBITS_ON_DETAIL_NOT_MATCH = 6020
    CREDITS_ON_DETAIL_NOT_MATCH = 6021
End Enum


Private Sub AnimationTimer_Timer()
    If Toggle = 1 Then
        iPic = iPic + 1
        If iPic = 2 Then
            iPic = 0
        End If
        Image1.Picture = PictureClip1.GraphicCell(iPic)
    End If

End Sub

Private Sub cmdCancel_Click()
Toggle = 0
Unload Me
End Sub

Private Sub cmdLoad_Click()
        
    
    giProcess = LOAD_AND_VALIDATE
    PerformLoad

End Sub



Private Sub cmdLoadPost_Click()
    
    Hourglass True
    giProcess = LOAD_VALIDATE_AND_POST
    PerformLoad

End Sub


Private Sub Command1_Click()

frmBrowser.Show

End Sub

Private Sub Command2_Click()
'2/14/2014 - AS Changed dictionary to collection
Dim colDetailRecs As New Collection
Dim oDet As New clsBAIDetail
Dim sEffectiveDate As String
Dim rsWork As New ADODB.Recordset
Dim cmdUpdateDD As New ADODB.Command
Dim ix As Long
Dim oDetail As clsBAIDetail
Dim sSql As String
Dim iGood As Integer
Dim bInvalid As Boolean
Dim sNewDDNum As String

If ReadNewDetailFile(App.Path & "\" & "WF030111.txt", colDetailRecs) = True Then

    For ix = 0 To colDetailRecs.Count - 1
        Set oDet = Nothing
        Set oDet = colDetailRecs.Item(ix)
        Debug.Print oDet.AccountID & " | " & oDet.IncomeSource & " | " & oDet.ACHRef & " | " & oDet.DDNumber
    Next ix
    
    Set cmdUpdateDD.ActiveConnection = gcnDDS
    If gStoredProcs("up_u_UpdateDDNum").GetStoredProcCommand(cmdUpdateDD) = False Then
        Err.Raise 2345, , "SSI Update Stored Procedure failed"
    End If
    
    'Lock the application
    If LockApplication = True Then
        'Err.Raise 3423534, , "Could not lock the database.  Some may be using database. Try again later."
        sSql = "SELECT INVALID_RECORD_ID,DD_NUM,FUNB_INCOME_SRC_TYPE, TOT_FUNB_BENEFIT_AMT, DR_CR_FLAG, AS_OF_DATETIME From DD_INVALID_REC"
        sSql = sSql & " WHERE RECORD_STATUS = 'A' AND FUNB_INCOME_SRC_TYPE = 'SUPP SEC'"
        
        rsWork.Open sSql, gcnDDS, adOpenStatic
        If rsWork.EOF Then
            Err.Raise 2323, , "No records to validate for SSI.", vbInformation
        End If
        Do Until rsWork.EOF
            If Left(rsWork!DD_NUM, 4) = "0000" Then
                'Check the records to see if we have a match from the file provided
                iGood = 0
                For ix = 0 To colDetailRecs.Count - 1
                    Set oDetail = Nothing
                    Set oDetail = colDetailRecs.Item(ix)
                   
                   'Keep going thru all the records it is possible to have two nunbers with the same ending 4 digits
                   bInvalid = False
                   If oDetail.AccountID <> rsWork!DD_NUM Then
                      bInvalid = True
                   End If
                   If oDetail.CRAmt > 0 And (rsWork!DR_CR_FLAG = "DR" Or rsWork!TOT_FUNB_BENEFIT_AMT <> oDetail.CRAmt) Then
                      bInvalid = True
                   End If
                   If oDetail.DRAmt > 0 And (rsWork!DR_CR_FLAG = "CR" Or rsWork!TOT_FUNB_BENEFIT_AMT <> oDetail.DRAmt) Then
                      bInvalid = True
                   End If
                   If CDate(oDetail.CheckwriteDate) <> CDate(rsWork!AS_OF_DATETIME) Then
                      bInvalid = True
                   End If
                   If bInvalid = False Then
                      iGood = iGood + 1
                      If iGood > 1 Then
                         MsgBox "Multiple entries in file for DD Number " & rsWork!DD_NUM & vbCrLf & "Resolve manually."
                         'set a number that will not update the record
                         iGood = 100
                      Else
                         If Len(oDetail.DDNumber) > 7 Then
                            sNewDDNum = oDetail.DDNumber
                         Else
                            iGood = 100
                         End If
                      End If
                   End If
                Next ix
                If iGood = 1 Then
                   cmdUpdateDD.Parameters("new_dd_num") = sNewDDNum
                   cmdUpdateDD.Parameters("invalid_record_id") = rsWork!INVALID_RECORD_ID
                   cmdUpdateDD.Parameters("user_id") = gobjLoginInfo.UserId
                   cmdUpdateDD.Execute
                End If
            End If
            rsWork.MoveNext
        Loop
        rsWork.Close
        UnlockApplication
    End If
End If

Xit:
Set rsWork = Nothing
Set cmdUpdateDD = Nothing
For ix = colDetailRecs.Count - 1 To 0
    colDetailRecs.Remove ix
Next ix
Set colDetailRecs = Nothing
Exit Sub

UpdateSSIDDNumErr:
MsgBox Error
Resume Xit


End Sub


Private Sub Command3_Click()
    ZipFiles gsDataPath & "\ToPost\MSG" & Format(Now(), "MMddyyyy") & ".7z", gsDataPath & "\ToPost\*.txt"

End Sub

Private Sub Form_Load()
Set OutlookTitle1.Picture = fMainForm.imlToolbarIcons.ListImages("Load FUNB File").Picture
'OptLoadOptions(0).Value = True
Image1.Picture = PictureClip1.GraphicCell(0)
End Sub


Public Sub ShowStatus(iPercent As Integer)
    
    proStatus.value = iPercent
    DoEvents
    
End Sub

Private Function CheckBAIFile(ByVal sFileName As String, ByRef sEffectiveDate As String) As Boolean
'4/4/2011 - AS Modified function to check for Supplemental social security
On Error GoTo CheckBaiFileErr

'AS - 2/16/2014 Removed reference to FileSystemObject
'Dim fso As New FileSystemObject
'Dim ts As TextStream
Dim ts As clsTextFile
Dim bRet As Boolean
Dim bFinished As Boolean
Dim i88Count As Integer
Dim sLine As String
Dim sRecType As String
Dim vSplit As Variant
Dim bMiscDebit As Boolean
If FileExists(sFileName) = True Then
    Set ts = New clsTextFile
    bRet = ts.OpenFile(sFileName, INPUT_TYPE)
    bRet = ts.ReadLine(sLine)
    Do Until EOF(ts.miFile)
        bFinished = False
        i88Count = 0
        If Left$(sLine, 2) = "02" Then 'Get the date for this file
            vSplit = Split(sLine, ",")
            sEffectiveDate = Mid$(vSplit(4), 3, 4) & Mid$(vSplit(4), 1, 2)
        End If
        If Left$(sLine, 2) = "16" Then '  The next three lines should be 88 records
            If Mid$(sLine, 4, 1) = "6" Then
                bMiscDebit = True
            Else
                bMiscDebit = False
            End If
            Do Until bFinished = True
                bRet = ts.ReadLine(sLine)
                sRecType = Left$(sLine, 2)
                Select Case sRecType
                Case "88"
                    If InStr(sLine, "SUPP SEC") > 0 Then
                       CheckBAIFile = True
                    End If
                    i88Count = i88Count + 1
                Case Else
                    bFinished = True
                End Select
            Loop
            If bMiscDebit = False Then ' This is a miscelaneous and wont get a number
                If i88Count = 2 Then
                    CheckBAIFile = True
                    Exit Do
                End If
            End If
        End If
        If Left$(sLine, 2) <> "16" Then
            bRet = ts.ReadLine(sLine)
        End If
    Loop
Else
    MsgBox sFileName & " does not exist.", vbCritical
End If

Xit:
On Error Resume Next
ts.CloseFile
Set ts = Nothing
'Set fso = Nothing
Exit Function

CheckBaiFileErr:
MsgBox Error, vbCritical
Resume Xit

End Function

Private Function ReadNewDetailFile(ByVal sFileName As String, ByRef rsDetailRecs As ADODB.Recordset) As Boolean
'AS 3/5/2012 - Change
On Error GoTo ReadNewDetailFileErr

'3/5/2012 - AS Connection will be used to create recordset from csv file
Dim cnn As New ADODB.Connection
'AS - 2/16/2014 - Removed all references to FileSystemObject
'Dim fso As New FileSystemObject
Dim FSO As New FSO
Dim sTempName As String
sTempName = "TempBaiDetail" & gobjLoginInfo.UserId & ".csv"
If FileExists(sFileName) = True Then
    If FileExists(gsDataPath & "\" & sTempName) = True Then
        FSO.DeleteAFile gsDataPath & "\" & sTempName
    End If
    FSO.CopyAFile sFileName, gsDataPath & "\" & sTempName
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & gsDataPath & ";" & _
            "Extended Properties=""Text;HDR=Yes;"""
   rsDetailRecs.Open "Select * from " & sTempName, cnn, adOpenStatic, adLockReadOnly, adCmdText
    ReadNewDetailFile = True
Else
    MsgBox sFileName & " does not exist", vbCritical
End If

'    Set ts = fso.OpenTextFile(sFileName, ForReading, False)
'    Do Until ts.AtEndOfStream
'        sLine = ts.ReadLine
'        Set vFields = Split(sLine, ",")
'        Set oDet = Nothing
'        Set oDet = New clsBAIDetail
'        oDet.DRAmt = CCur(collMatches(0).SubMatches(0))
'        oDet.CRAmt = CCur(collMatches(0).SubMatches(1))
'                    oDet.CheckwriteDate = sCheckwriteDate
'                Else
'                    Err.Raise 34234, , "Could not determine debit and credit amounts " & vbCrLf & sLine
'                End If
'                bFoundId = False
'                iCnt = 0
'                Do Until iCnt > 10 Or bFoundId = True
'                   sLine = ts.ReadLine
'                   Set collMatches = Regexp(sLine, "ID:[ ]+([A-Za-z0-9]+)[ ]+([A-Za-z0-9]+)")
'                   If collMatches.Count > 0 Then
'                        oDet.AccountID = collMatches(0).SubMatches(0)
'                        oDet.IncomeSource = collMatches(0).SubMatches(1)
'                        bFoundId = True
'                   Else
'                        Set collMatches = Regexp(sLine, "TRANSACTION REF #   : ([0-9]+)")
'                        If collMatches.Count > 0 Then
'                             oDet.ACHRef = CCur(collMatches(0).SubMatches(0))
'                        End If
'                   End If
'                   iCnt = iCnt + 1
'                Loop
'                If bFoundId = False Then
'                   ' Err.Raise 34235, , "Could not find id number. Detail File Layout changed please call for support" & vbCrLf
'                End If
'                iCnt = 0
'                bFoundAddenda = False
'                If oDet.IncomeSource = "SSI" Then
'
'                Do Until iCnt > 10 Or bFoundAddenda = True
'                   sLine = ts.ReadLine
'                   Set collMatches = Regexp(sLine, "N1\*BE\*([A-Za-z +)\*34\*([0-9]+)")
'                   If collMatches.Count > 0 Then
'                        vSplit = Split(collMatches(0).SubMatches(0), "*")
'                        If UBound(vSplit) > 1 Then
'                            oDet.PatientName = vSplit(0)
'                            oDet.DDNumber = vSplit(2)
'                            bFoundAddenda = True
'                        End If
'                   End If
'                   If Left$(sLine, 10) = "----------" Then
'                       iCnt = 11
'                   End If
'                   iCnt = iCnt + 1
'                Loop
'                '8/5/2011 - AS - Added code to look at the account number
'                If Left$(oDet.AccountID, 4) <> "0000" Then
'                   oDet.DDNumber = oDet.AccountID
'                   bFoundAddenda = True
'                End If
'                If bFoundAddenda = False Then
'                    Err.Raise 34236, , "Could not find DD Number from addenda." & vbCrLf
'                Else
'                    colDetailRecs.Add "D" & oDet.ACHRef, oDet
'                End If
'                End If
'            End If
'        Loop
'    End If
'End If
Xit:
cnn.Close
Set cnn = Nothing
Set FSO = Nothing
Exit Function
ReadNewDetailFileErr:
If Err.Number = 34236 Then
    Resume Next
Else
    MsgBox "Please Contact Dirm for ReadNewDetailFile.  " & vbCrLf & Error, vbCritical
   'End
    
End If
    
Resume
End Function

Private Sub ReadDetailFile(ByVal sFileName As String, ByRef colDetailRecs As Collection)
'AS - 2/14/2014 - Removed references to FileSystemObject
'Dim fso As New FileSystemObject
'Dim ts As TextStream
Dim ts As clsTextFile
Dim sLine As String
Dim ix As Integer
Dim sCheckwriteDate As String
Dim bFoundDate As Boolean
Dim oDet As clsBAIDetail
Dim sChar As String
Dim sDDNum As String
Dim bRet As Boolean

If FileExists(sFileName) = True Then
    Set ts = New clsTextFile
    bRet = ts.OpenFile(sFileName, INPUT_TYPE)
    Do Until EOF(ts.miFile)
        bRet = ts.ReadLine(sLine)
        If bFoundDate = False Then
            ix = InStr(1, sLine, "EFT ADVICE REPORT FOR")
            If ix > 0 Then
                sCheckwriteDate = Mid$(sLine, ix + 22, 8)
                bFoundDate = True
                Exit Do
            End If
        End If
    Loop
    If bFoundDate = True Then
        Do Until EOF(ts.miFile)
            bRet = ts.ReadLine(sLine)
            ix = InStr(1, sLine, "TYPE           DR AMT")
            If ix > 0 Then
                Set oDet = Nothing
                Set oDet = New clsBAIDetail
                'We have a new entry Read the next five lines
                bRet = ts.ReadLine(sLine)
                oDet.DRAmt = CCur(Trim(Mid$(sLine, 14, 14)))
                oDet.CRAmt = CCur(Trim(Mid$(sLine, 28, 17)))
                oDet.IncomeSource = Trim(Mid$(sLine, 45, 15))
                oDet.ACHRef = CLng(Mid$(sLine, 69, 8))
                'Advance three lines
                bRet = ts.ReadLine(sLine)
                bRet = ts.ReadLine(sLine)
                bRet = ts.ReadLine(sLine)
                oDet.OrigCompanyID = Trim(Mid$(sLine, 22, 18))
                'Advance two lines and get the DDNumber
                bRet = ts.ReadLine(sLine)
                bRet = ts.ReadLine(sLine)
                bRet = ts.ReadLine(sLine)
                sDDNum = Trim(Mid$(sLine, 5, 15))
                Select Case Right$(sDDNum, 3)
                Case "SSI", "SSA", "CSF"
                    sDDNum = Trim$(Left$(sDDNum, Len(sDDNum) - 3))
                End Select
                'Remove any - or space
                oDet.AccountID = ""
                For ix = 1 To Len(sDDNum)
                    sChar = Mid$(sDDNum, ix, 1)
                    Select Case sChar
                    Case " ", "-"
                        'Do Nothing
                    Case Else
                        oDet.AccountID = oDet.AccountID & sChar
                    End Select
                Next ix
                bRet = ts.ReadLine(sLine)
                bRet = ts.ReadLine(sLine)
                bRet = ts.ReadLine(sLine)
                If Left$(sLine, 7) = "ADDENDA" Then
                    bRet = ts.ReadLine(sLine)
                    oDet.Note = sLine
                End If
                colDetailRecs.Add oDet, "D" & oDet.ACHRef
            End If
        Loop
    End If
End If

ts.CloseFile
Set oDet = Nothing
Set ts = Nothing
'Set fso = Nothing

End Sub

Public Function Regexp(ByVal sSearchString As String, ByVal sPattern As String) As MatchCollection

Dim oRegExp As Regexp
Dim collMatches As MatchCollection
Set oRegExp = New Regexp
oRegExp.IgnoreCase = True
oRegExp.Global = True
oRegExp.Pattern = sPattern
Set collMatches = oRegExp.Execute(sSearchString)
Set Regexp = collMatches

Set oRegExp = Nothing
End Function

Private Function LoadFUNBFile(ByRef rsDetailRecs As ADODB.Recordset, ByVal sFilePath As String, ByVal sFileName As String, Optional ByVal sDetailName As String) As Boolean

'********************************************************************************
'* Name: LoadFUNBFile
'*
'* Description: Brings data from the FUNB File and loads it into Work file
'* Parameters:
'* 3/7/2012 - Modified logic to accept funb file wil multiple 02 records and to read in csv detail file
'* 3/11/2014 - Modified logic to add misc debits when misc debits are excluded from detail file
'* 10/16/2014 - Modified logic to add misc credits when misc credits are excluded from detail file
'* 12/04/2015 - Modified logic to utilize new Wells Fargo detail file.
'               Changes to fields Entry Class Description, ID, Individual Name changed to Entry Description,Recipeint ID, Recipient Name
'* Created: 6/7/99 1:09:08 PM
'********************************************************************************

'This Function loads data into our Transaction invoice header file
'and transaction invoice detail file from our input file
On Error GoTo LoadFUNBFileErr

Dim cmdSummary As New ADODB.Command
Dim rsConfig As New ADODB.Recordset
Dim cmdBAIFile As New ADODB.Command
Dim cmdWork As New ADODB.Command

Dim dbFUNB As DAO.Database
Dim tdfLinked As DAO.TableDef
'Dim rsFUNB As DAO.Recordset
Dim rsWork As DAO.Recordset

Dim ix As Integer
Dim iy As Integer
Dim iDtlAccum As Integer
Dim sSql As String
Dim sLastRecRead As String
Dim bIngroup As Boolean
Dim bInAccount As Boolean
Dim bInFile As Boolean
Dim sAsOfDate As String
Dim sAsOfTime As String
Dim sCustAcctNum As String
Dim dTotalDebits As Double
Dim dTotalCredits As Double
Dim dCheckDebits As Double
Dim dCheckCredits As Double
Dim dOpeningLedger As Double
Dim dClosingLedger As Double
Dim dOpeningAvailable As Double
Dim dClosingAvailable As Double
Dim sIncomeSourceType As String
Dim sNote As String
Dim sFullNote As String
Dim sBaiFileId(1 To 50) As String
Dim iBaiCount As Integer
Dim iFieldIndex As Integer
Dim iPercentComplete As Integer
Dim sBAIFileDate As String
Dim rsIncomeType As New ADODB.Recordset
Dim i88Count As Integer
Dim sDDNum As String
Dim dTotFUNBAmt As Double
Dim sDrCrFlag As String
Dim dErrNoStore As Double
Dim oFile As New clsTextFile
Dim rsFUNB As ADODB.Recordset
Dim sTransRef As String

Dim sLine As String
Dim vSplit As Variant
Dim iLast As Integer
Dim iLast88 As Integer
Dim iGood As Integer
Dim oDetail As clsBAIDetail
Dim sTransCode As String
Dim bInvalid As Boolean
Dim sNewDDNum As String
Dim sASOfDateFull As String
Dim sTempDate As String
Dim sDateField As String
Dim dMiscDebits As Double
Dim dMiscCredits As Double
Dim bMiscDebitsMissing As Boolean
Dim bMiscCreditsMissing As Boolean

Dim testcounter As Integer

sSql = "SELECT FUNB_INCOME_SRC_TYPE,START_POS,LENGTH,DET_START_POS FROM DD_INCOME_SOURCE_TYPE WHERE RECORD_STATUS = 'A'"
rsIncomeType.CursorLocation = adUseServer
rsIncomeType.Open sSql, gcnDDS, adOpenStatic
If rsIncomeType.EOF Then
    MsgBox "Error opening Income Source Table"
    LoadFUNBFile = False
    Exit Function
End If

lblEditStatus.Caption = "Loading FUNB File.  Please wait until process is complete..."

On Error GoTo LoadFUNBFileErr

Set rsFUNB = New ADODB.Recordset
For ix = 1 To 75
    rsFUNB.Fields.Append "Col" & ix, adVariant
Next ix
sLine = "Start"
rsFUNB.Open
oFile.OpenFile sFilePath & "\" & sFileName, INPUT_TYPE
' Read in the BAI File and put it into a recordset.  We will use this to find totals
Do Until sLine = ""
    oFile.ReadLine sLine
    If sLine <> "" Then
        vSplit = Split(sLine, ",")
        rsFUNB.AddNew
        If vSplit(0) = "03" Then
            iLast = UBound(vSplit)
            For iy = 0 To iLast
                If iy = iLast Then
                    rsFUNB.Fields(iy) = Replace(vSplit(iy), "/", "")
                Else
                    rsFUNB.Fields(iy) = vSplit(iy)
                End If
            Next iy
            'Read the First Continuation Line
            oFile.ReadLine sLine
            vSplit = Split(sLine, ",")
            If vSplit(0) = "88" Then
                iLast88 = UBound(vSplit)
                For iy = 1 To iLast88
                    If iy = iLast88 Then
                        rsFUNB.Fields(iy + iLast) = Replace(vSplit(iy), "/", "")
                    Else
                        rsFUNB.Fields(iy + iLast) = vSplit(iy)
                    End If
                Next iy
                iLast = iLast + iLast88
            Else
                Err.Raise 2505
            End If
            'Read the second continuation line
            oFile.ReadLine sLine
            vSplit = Split(sLine, ",")
            If vSplit(0) = "88" Then
                iLast88 = UBound(vSplit)
                For iy = 1 To iLast88
                    If iy = iLast88 Then
                        rsFUNB.Fields(iy + iLast) = Replace(vSplit(iy), "/", "")
                    Else
                        rsFUNB.Fields(iy + iLast) = vSplit(iy)
                    End If
                Next iy
                iLast = iLast + iLast88
                
            Else
                'There was not two continue lines update the 03 record
                rsFUNB.Update
                rsFUNB.AddNew
                iLast = UBound(vSplit)
                For iy = 0 To iLast
                    If iy = iLast Then
                        If Right$(vSplit(iy), 1) = "/" Then
                            rsFUNB.Fields(iy) = Replace(vSplit(iy), "/", "")
                            rsFUNB.Fields(iy + 1) = "/"
                        Else
                            rsFUNB.Fields(iy) = Replace(vSplit(iy), "/", "")
                        End If
                    Else
                        rsFUNB.Fields(iy) = vSplit(iy)
                    End If
                Next iy
            End If
        Else
            iLast = UBound(vSplit)
            For iy = 0 To iLast
                If iy = iLast Then
                    If Right$(vSplit(iy), 1) = "/" Then
                        rsFUNB.Fields(iy) = Replace(vSplit(iy), "/", "")
                        rsFUNB.Fields(iy + 1) = "/"
                    Else
                        rsFUNB.Fields(iy) = Replace(vSplit(iy), "/", "")
                    End If
                Else
                    rsFUNB.Fields(iy) = vSplit(iy)
                End If
            Next iy
        End If
        rsFUNB.Update
    End If
Loop
' We now have the FUNB File Loaded into a recordet.  Close the file
oFile.CloseFile

sSql = "SELECT SENDER_ID,RECEIVER_ID FROM DD_CONFIG_INFO"
rsConfig.Open sSql, gcnDDS, adOpenForwardOnly, adLockReadOnly, adCmdText
If rsConfig.EOF Then
    Err.Raise COULD_NOT_GET_CONFIG_INFO
End If

Set cmdBAIFile.ActiveConnection = gcnDDS
If gStoredProcs("up_iud_BAI_File_Summary").GetStoredProcCommand(cmdBAIFile) = False Then
    Err.Raise BAI_FILE_STORED_PROC_FAILED
End If

Set cmdWork.ActiveConnection = gcnDDS
If gStoredProcs("up_iud_DDWorkFile").GetStoredProcCommand(cmdWork) = False Then
    Err.Raise WORK_FILE_STORED_PROC_FAILED
End If

ShowStatus 10


testcounter = 1
With rsFUNB
    .MoveLast
    .MoveFirst

    Do Until .EOF ' Loop until end of file.
    
        If testcounter = 1257 Then
        testcounter = 1257
        End If
        
        testcounter = testcounter + 1
        iPercentComplete = CInt((rsFUNB.AbsolutePosition / rsFUNB.RecordCount / 100 * 90) + 10)
        ShowStatus iPercentComplete

        Select Case !Col1
        Case FILE_HDR_REC
            If bInFile = False Then
                'Verify that this is a valid BAI Text file
                If !Col1 <> FILE_HDR_REC Then
                    'This is No Header record.  BAI File always starts with a Record Type of 01"
                    Err.Raise NO_HEADER_RECORD
                End If
                iBaiCount = iBaiCount + 1

                sLastRecRead = FILE_HDR_REC
                'Verify the correct sender and receiver id
                If !Col2 <> rsConfig!SENDER_ID Then
                    Err.Raise SENDER_ID_RCVD_INVALID
                End If

                If !Col3 <> rsConfig!RECEIVER_ID Then
                    Err.Raise RECEIVER_ID_RCVD_INVALID
                End If
                If Len(!Col4) = 6 Then
                    sBAIFileDate = Convert_yymmdd_date(!Col4)
                Else
                    sBAIFileDate = Convert_yyyymmdd_date(!Col4)
                End If
                On Error GoTo LoadFUNBFileErr
                bInFile = True
            Else
                Err.Raise UNEXPECTED_FILE_HDR
            End If
            sLastRecRead = FILE_HDR_REC
        Case GROUP_HDR_REC
            If bIngroup = True Then
                Err.Raise UNEXPECTED_GROUP_HDR
            Else
                bIngroup = True
                dMiscDebits = 0
                dMiscCredits = 0
                bMiscDebitsMissing = False
                bMiscCreditsMissing = False
                sAsOfDate = ConvertNull(!Col5)
                sAsOfTime = ConvertNull(!Col6)
                sASOfDateFull = Convert_yymmdd_date(sAsOfDate)
                'Check to see if we have processed this file before
                cmdBAIFile.Parameters("bai_file_id") = Null
                cmdBAIFile.Parameters("bai_file_datetime") = sBAIFileDate
                If IsNull(!Col5) Then
                    'Do Nothing
                    cmdBAIFile.Parameters("file_id_num") = ""
                Else
                    If Len(!Col5) = 6 Then
                        'Convert a date of yymmdd to yyyymmdd
                        cmdBAIFile.Parameters("file_id_num") = Format$(Convert_yymmdd_date(!Col5), "yyyymmdd")
                    Else
                        cmdBAIFile.Parameters("file_id_num") = Format$(Convert_yyyymmdd_date(!Col5), "yyyymmdd")
                    End If
                End If
                
                cmdBAIFile.Parameters("available_bal") = 0
                cmdBAIFile.Parameters("collected_bal") = 0
                cmdBAIFile.Parameters("created_by") = gobjLoginInfo.UserId
                cmdBAIFile.Parameters("funb_total_credits") = 0
                cmdBAIFile.Parameters("funb_total_debits") = 0
                cmdBAIFile.Parameters("ledger_bal") = 0
                cmdBAIFile.Parameters("update_status") = "I"
                cmdBAIFile.Execute
                If cmdBAIFile.Parameters("RETURN_VALUE") = FILE_PROCESSED_ALREADY Then
                    Err.Raise FILE_PROCESSED_ALREADY
                End If
                sBaiFileId(iBaiCount) = cmdBAIFile.Parameters("bai_file_id_OUTPUT")

            End If
            sLastRecRead = GROUP_HDR_REC
        Case ACCOUNT_HDR_REC
            If bInAccount = True Then
                Err.Raise UNEXPECTED_ACCOUNT_HDR
            Else
                If bIngroup = False Then
                    Err.Raise NO_GROUP_REC_RECEIVED
                Else
                    bInAccount = True
                    dCheckDebits = 0
                    dCheckCredits = 0
                    sCustAcctNum = !Col2
                    'Check the type values
                    For iFieldIndex = 3 To 71 Step 4
                        If Not IsNull(.Fields(iFieldIndex)) And Not IsNull(.Fields(iFieldIndex + 1)) Then
                            Select Case .Fields(iFieldIndex)
                            Case "100" ' Total Credits
                                dTotalCredits = .Fields(iFieldIndex + 1) / 100
                            Case "400" ' Total Debits
                                dTotalDebits = .Fields(iFieldIndex + 1) / 100
                            Case "010"
                                dOpeningLedger = .Fields(iFieldIndex + 1) / 100
                            Case "015"
                                dClosingLedger = .Fields(iFieldIndex + 1) / 100
                            Case "040"
                                dOpeningAvailable = .Fields(iFieldIndex + 1) / 100
                            Case "045"
                                dClosingAvailable = .Fields(iFieldIndex + 1) / 100
                            End Select
                        End If
                    Next iFieldIndex
                    sLastRecRead = ACCOUNT_HDR_REC
                End If
            End If
        Case TRANS_DTL_REC
            sLastRecRead = TRANS_DTL_REC
            If CInt(.Fields(1)) >= 200 And CInt(.Fields(1)) < 300 Then
                'We have a miscellaneous debit amount.  THey do not appear in the detail section.  Total it up
                dMiscDebits = dMiscDebits + Round(.Fields(2) / 100, 2)
            End If
            If CInt(.Fields(1)) >= 300 And CInt(.Fields(1)) < 400 Then
                'We have a miscellaneous debit amount.  THey do not appear in the detail section.  Total it up
                dMiscCredits = dMiscCredits + Round(.Fields(2) / 100, 2)
            End If
            If CInt(.Fields(1)) >= 400 And CInt(.Fields(1)) < 500 Then
                'We have a miscellaneous debit amount.  THey do not appear in the detail section.  Total it up
                dMiscDebits = dMiscDebits + Round(.Fields(2) / 100, 2)
            End If
            If CInt(.Fields(1)) >= 500 And CInt(.Fields(1)) < 600 Then
                'We have a miscellaneous debit amount.  THey do not appear in the detail section.  Total it up
                dMiscCredits = dMiscCredits + Round(.Fields(2) / 100, 2)
            End If
            If CInt(.Fields(1)) >= 600 And CInt(.Fields(1)) < 700 Then
                'We have a miscellaneous debit amount.  THey do not appear in the detail section.  Total it up
                dMiscDebits = dMiscDebits + Round(.Fields(2) / 100, 2)
            End If
        Case CONTINUE_REC
            Select Case sLastRecRead
            Case ACCOUNT_HDR_REC
                'Check the account type values
                For iFieldIndex = 1 To 21 Step 4
                    If Not IsNull(.Fields(iFieldIndex)) And Not IsNull(.Fields(iFieldIndex + 1)) Then
                        Select Case .Fields(iFieldIndex)
                        Case "010"
                            dOpeningLedger = .Fields(iFieldIndex + 1) / 100
                        Case "015"
                            dClosingLedger = .Fields(iFieldIndex + 1) / 100
                        Case "040"
                            dOpeningAvailable = .Fields(iFieldIndex + 1) / 100
                        Case "045"
                            dClosingAvailable = .Fields(iFieldIndex + 1) / 100
                        Case "100"
                            dTotalCredits = .Fields(iFieldIndex + 1) / 100
                        Case "400"
                            dTotalDebits = .Fields(iFieldIndex + 1) / 100
                        Case Else
                            'Ignore
                        End Select
                    End If
                Next iFieldIndex
            End Select
        Case ACCOUNT_TRLR_REC
            If bIngroup = False Or bInAccount = False Then
                'Err.Raise UNEXPECTED_ACCOUNT_TRLR
            End If
            sLastRecRead = ACCOUNT_TRLR_REC
            bInAccount = False
        Case GROUP_TRLR_REC
            If bIngroup = False Then
                Err.Raise UNEXPECTED_GROUP_TRLR
            Else
                bIngroup = False
            End If
            rsDetailRecs.Filter = ""
            'Check the details to see if the debits and credits match
            dCheckDebits = 0
            dCheckCredits = 0
            If rsDetailRecs.BOF And rsDetailRecs.EOF Then
               'Skip
            Else
                If dTotalDebits = 0 And dTotalCredits = 0 Then
                    'Skip
                Else
                    rsDetailRecs.MoveFirst
                    sDateField = "[" & rsDetailRecs.Fields(0).Name & "]"
                    If InStr(1, rsDetailRecs.Fields(0), "/") > 3 Then
                        'The Date format is like 2012/01/01
                        sTempDate = Format(sASOfDateFull, "yyyy/mm/dd")
                        rsDetailRecs.Filter = sDateField & " = '" & sTempDate & "'"
                    Else
                        rsDetailRecs.Filter = sDateField & " = #" & sASOfDateFull & "#"
                    End If
                    
                    Do Until rsDetailRecs.EOF
                        dCheckDebits = dCheckDebits + rsDetailRecs![Debit Amount]
                        dCheckCredits = dCheckCredits + rsDetailRecs![Credit Amount]
                        rsDetailRecs.MoveNext
                    Loop
                End If
            End If
            If CCur(dCheckDebits) <> CCur(dTotalDebits) Then
                If CCur(dCheckDebits) <> CCur(dTotalDebits) - CCur(dMiscDebits) Then
                    Err.Raise ACCOUNT_CREDITS_DONT_MATCH
                Else
                    bMiscDebitsMissing = True
                End If
            End If
            If CCur(dCheckCredits) <> CCur(dTotalCredits) Then
                If CCur(dCheckCredits) <> CCur(dTotalCredits) - CCur(dMiscCredits) Then
                    Err.Raise ACCOUNT_CREDITS_DONT_MATCH
                Else
                    bMiscCreditsMissing = True
                End If
            End If
            'Insert our records into the work file
            If dTotalDebits = 0 And dTotalCredits = 0 Then
                    'Skip
               'Skip there are no details
            Else
                'Advance to the first record
                If rsDetailRecs.RecordCount <> 0 Then
                    rsDetailRecs.MoveFirst
                End If
                
                Dim testcounter2 As Integer
                testcounter2 = 1
                Do Until rsDetailRecs.EOF
                testcounter2 = testcounter2 + 1
                    For ix = 0 To cmdWork.Parameters.Count - 1
                        cmdWork.Parameters(ix) = Null
                    Next ix
                    cmdWork.Parameters("bai_file_id") = CDbl(sBaiFileId(iBaiCount))
                    If rsDetailRecs![Credit Amount] > 0 Then
                       cmdWork.Parameters("tot_funb_benefit_amt") = rsDetailRecs![Credit Amount]
                       cmdWork.Parameters("dr_cr_flag") = "CR"
                    Else
                       cmdWork.Parameters("tot_funb_benefit_amt") = rsDetailRecs![Debit Amount]
                       cmdWork.Parameters("dr_cr_flag") = "DR"
                    End If
                    cmdWork.Parameters("as_of_datetime") = sASOfDateFull
                    'AS - 12/4/2015 - Field name has changed from Entry Class Description to Entry Description
                    'sIncomeSourceType = rsDetailRecs![Entry Class Description]
                    sIncomeSourceType = rsDetailRecs![Entry Description]
                    If Left$(sIncomeSourceType, 2) = "XX" Then
                        sIncomeSourceType = Right$(sIncomeSourceType, Len(sIncomeSourceType) - 2)
                    End If
                    If sIncomeSourceType = "VA BENEF" Then
                        sIncomeSourceType = "VA BENEFIT"
                    End If
    
                    'Find The income Source
                    rsIncomeType.MoveFirst
                    rsIncomeType.Find "FUNB_INCOME_SRC_TYPE = '" & sIncomeSourceType & "'", , adSearchForward
                    'AS - 12/4/2015 - Field Id changed to Recipient ID
                    'vSplit = Split(rsDetailRecs!Id, " ")
                    vSplit = Split(rsDetailRecs.Fields("Recipient ID"), " ")
                    sDDNum = vSplit(0)
                    'AS - To Do We need some code here to extract part of the DD Number for Civil Service
                    If sIncomeSourceType = "CIV SERV" Then
                        If UBound(vSplit) > 0 Then
                            sDDNum = vSplit(1)
                        End If
                    End If
                    cmdWork.Parameters("income_source_type") = sIncomeSourceType
                    'sFullNote = rsDetailRecs![Id] & vbCrLf & rsDetailRecs![First Addenda] & vbCrLf & rsDetailRecs![Individual Name]
                    sFullNote = rsDetailRecs.Fields("Recipient ID") & vbCrLf & rsDetailRecs![First Addenda] & vbCrLf & rsDetailRecs![Recipient Name]
                    cmdWork.Parameters("comment") = sFullNote
                    cmdWork.Parameters("dd_num") = Left$(sDDNum, 11)
                    cmdWork.Parameters("created_by") = gobjLoginInfo.UserId
                    cmdWork.Parameters("shared_dd_num_ind") = "N"
                    cmdWork.Parameters("deceased_ind") = "N"
                    cmdWork.Parameters("record_status") = "A"
                    cmdWork.Parameters("validated") = "N"
                    cmdWork.Parameters("update_status") = "I"
                    cmdWork.Execute
                    If cmdWork.Parameters("RETURN_VALUE") <> 0 Then
                        Err.Raise ERROR_CREATING_WORK_RECORD
                    End If
                    sFullNote = vbNullString
                    rsDetailRecs.MoveNext
                Loop
                
                If bMiscDebitsMissing = True Then 'Add a miscellaneous Debit
                    For ix = 0 To cmdWork.Parameters.Count - 1
                        cmdWork.Parameters(ix) = Null
                    Next ix
                    cmdWork.Parameters("bai_file_id") = CDbl(sBaiFileId(iBaiCount))
                    cmdWork.Parameters("tot_funb_benefit_amt") = dMiscDebits
                    cmdWork.Parameters("dr_cr_flag") = "DR"
                    cmdWork.Parameters("as_of_datetime") = sASOfDateFull
                    sDDNum = ""
                    cmdWork.Parameters("income_source_type") = "BANK DEBIT"
                    sFullNote = "Total Misc Debits"
                    cmdWork.Parameters("comment") = sFullNote
                    cmdWork.Parameters("dd_num") = Left$(sDDNum, 11)
                    cmdWork.Parameters("created_by") = gobjLoginInfo.UserId
                    cmdWork.Parameters("shared_dd_num_ind") = "N"
                    cmdWork.Parameters("deceased_ind") = "N"
                    cmdWork.Parameters("record_status") = "A"
                    cmdWork.Parameters("validated") = "N"
                    cmdWork.Parameters("update_status") = "I"
                    cmdWork.Execute
                    If cmdWork.Parameters("RETURN_VALUE") <> 0 Then
                        Err.Raise ERROR_CREATING_WORK_RECORD
                    End If
                End If
            
                If bMiscCreditsMissing = True Then 'Add a miscellaneous Credit
                    For ix = 0 To cmdWork.Parameters.Count - 1
                        cmdWork.Parameters(ix) = Null
                    Next ix
                    cmdWork.Parameters("bai_file_id") = CDbl(sBaiFileId(iBaiCount))
                    cmdWork.Parameters("tot_funb_benefit_amt") = dMiscCredits
                    cmdWork.Parameters("dr_cr_flag") = "CR"
                    cmdWork.Parameters("as_of_datetime") = sASOfDateFull
                    sDDNum = ""
                    cmdWork.Parameters("income_source_type") = "BANK CREDT"
                    sFullNote = "Total Misc Credits"
                    cmdWork.Parameters("comment") = sFullNote
                    cmdWork.Parameters("dd_num") = Left$(sDDNum, 11)
                    cmdWork.Parameters("created_by") = gobjLoginInfo.UserId
                    cmdWork.Parameters("shared_dd_num_ind") = "N"
                    cmdWork.Parameters("deceased_ind") = "N"
                    cmdWork.Parameters("record_status") = "A"
                    cmdWork.Parameters("validated") = "N"
                    cmdWork.Parameters("update_status") = "I"
                    cmdWork.Execute
                    If cmdWork.Parameters("RETURN_VALUE") <> 0 Then
                        Err.Raise ERROR_CREATING_WORK_RECORD
                    End If
                End If
            
            End If
            'Rewrite the Summary Record
            cmdBAIFile.Parameters("bai_file_id") = sBaiFileId(iBaiCount)
            cmdBAIFile.Parameters("available_bal") = dClosingAvailable
            cmdBAIFile.Parameters("collected_bal") = 0
            cmdBAIFile.Parameters("funb_total_credits") = dTotalCredits
            cmdBAIFile.Parameters("funb_total_debits") = dTotalDebits
            cmdBAIFile.Parameters("ledger_bal") = dClosingLedger
            cmdBAIFile.Parameters("update_status") = "U"
            'Sybase 12 will not accept default values put in null
            'cmdBAIFile.Parameters("bai_file_datetime") = Null
            cmdBAIFile.Parameters("created_by") = Null
            cmdBAIFile.Execute

            sLastRecRead = GROUP_TRLR_REC
        Case FILE_TRLR_REC
            If bInFile = False Or bIngroup = True Or bInAccount = True Then
                Err.Raise UNEXPECTED_FILE_TRLR_REC
            End If
            sLastRecRead = FILE_TRLR_REC
            bInFile = False
            dCheckCredits = 0
            dCheckDebits = 0
        End Select

        rsFUNB.MoveNext
        If rsFUNB.EOF Then
          Exit Do
        End If
    Loop

    If bInFile = True Then
        Err.Raise NO_FILE_TRAILER_RECEIVED
    End If
    
    If iBaiCount = 0 Then
        Err.Raise NO_VALID_BAI_FILES_RECEIVED
    End If

End With


cmdLoad.Enabled = True
Command1.Enabled = True
cmdCancel.Enabled = True

LoadFUNBFile = True

Xit:
    
    On Error Resume Next
    
    Set oDetail = Nothing
    Set cmdSummary = Nothing
    Set rsConfig = Nothing
    Set dbFUNB = Nothing
    Set rsFUNB = Nothing
    Set rsWork = Nothing
    Set cmdBAIFile = Nothing
    Set rsIncomeType = Nothing
    Set cmdWork = Nothing
    Set rsDetailRecs = Nothing
    ShowStatus 100
    Image1.Picture = PictureClip1.GraphicCell(0)
    Toggle = 0
    cmdCancel.Enabled = True
    Hourglass False
    
Exit Function

LoadFUNBFileErr:
    dErrNoStore = Err
'Resume
    For ix = 0 To cmdWork.Parameters.Count - 1
        cmdWork.Parameters(ix) = Null
    Next ix

    'Delete the Bai File summary record
    If iBaiCount > 0 Then
        For ix = 1 To iBaiCount
            If sBaiFileId(ix) = vbNullString Then
                Exit For
            Else
                With cmdWork
                cmdWork.Parameters("bai_file_id") = sBaiFileId(ix)
                cmdWork.Parameters("update_status") = "D"
                cmdWork.Execute
                If cmdWork.Parameters("RETURN_VALUE") <> 0 Then
                    MsgBox "Fatal Error deleting records from BAI File.  Please call your Direct Deposit Support person for assistance.", vbCritical
                End If
                End With
                
                With cmdBAIFile
                .Parameters("bai_file_id") = sBaiFileId(ix)
                .Parameters("update_status") = "D"
                .Execute
                End With
            End If
        Next ix
    End If
    Select Case dErrNoStore
    Case FILE_PROCESSED_ALREADY
        lblEditStatus.Caption = vbCrLf & vbCrLf & "FUNB File was previously processed..."
    Case NO_HEADER_RECORD
        lblEditStatus.Caption = "Invalid FUNB file." & vbCrLf & "No header record was found."
    Case SENDER_ID_RCVD_INVALID
        lblEditStatus.Caption = "Invalid FUNB file." & vbCrLf & "Sender Id does not match DDS Configuration value."
    Case RECEIVER_ID_RCVD_INVALID
        lblEditStatus.Caption = "Invalid FUNB file." & vbCrLf & "Receiver Id does not match DDS Configuration value."
    Case UNEXPECTED_FILE_HDR
        lblEditStatus.Caption = "Invalid FUNB file." & vbCrLf & "Unexpected File Header record. Line " & rsFUNB.AbsolutePosition
    Case UNEXPECTED_GROUP_HDR
        lblEditStatus.Caption = "Invalid FUNB file." & vbCrLf & "Unexpected Group Header record. Line " & rsFUNB.AbsolutePosition
    Case UNEXPECTED_ACCOUNT_HDR
        lblEditStatus.Caption = "Invalid FUNB file." & vbCrLf & "Unexpected Account Header record. Line " & rsFUNB.AbsolutePosition
    Case NO_GROUP_REC_RECEIVED
        lblEditStatus.Caption = "Invalid FUNB file." & vbCrLf & "No group header received. Line " & rsFUNB.AbsolutePosition
    Case UNEXPECTED_TRANS_DTL
        lblEditStatus.Caption = "Invalid FUNB file." & vbCrLf & "Unexpected transaction detail record. Line " & rsFUNB.AbsolutePosition
    Case UNEXPECTED_ACCOUNT_TRLR
        lblEditStatus.Caption = "Invalid FUNB file." & vbCrLf & "Unexpected Account Trailer record. Line " & rsFUNB.AbsolutePosition
    Case UNEXPECTED_GROUP_TRLR
        lblEditStatus.Caption = "Invalid FUNB file." & vbCrLf & "Unexpected Group Trailer record. Line " & rsFUNB.AbsolutePosition
    Case UNEXPECTED_FILE_TRLR_REC
        lblEditStatus.Caption = "Invalid FUNB file." & vbCrLf & "Unexpected File Trailer record. Line " & rsFUNB.AbsolutePosition
    Case ACCOUNT_CREDITS_DONT_MATCH
        lblEditStatus.Caption = "Invalid FUNB file." & vbCrLf & "Total credits in account summary do not equal the detail total credits. Line " & rsFUNB.AbsolutePosition
    Case ACCOUNT_DEBITS_DONT_MATCH
        lblEditStatus.Caption = "Invalid FUNB file." & vbCrLf & "Total debits in account summary do not equal the detail total debits. Line " & rsFUNB.AbsolutePosition
    Case BAI_FILE_STORED_PROC_FAILED, FUNB_SUMMARY_STORED_PROC_FAILED
        lblEditStatus.Caption = "Stored procedure could not be created."
    Case NO_VALID_BAI_FILES_RECEIVED
        lblEditStatus.Caption = "No valid FUNB Header records received"
    Case COULD_NOT_GET_CONFIG_INFO
        lblEditStatus.Caption = "Could not access the configuration information."
    Case ERROR_CREATING_WORK_RECORD
        lblEditStatus.Caption = "Error creating detail record."
    Case WORK_FILE_STORED_PROC_FAILED
        lblEditStatus.Caption = "Work File Stored procedure could not be created."
    Case NO_FILE_TRAILER_RECEIVED
        lblEditStatus.Caption = "FUNB File did not end properly.  This indicates an incomplete file."
    Case Else
        MsgBox Error
        lblEditStatus.Caption = CStr(Error)
        Resume Xit
    End Select
    cmdCancel.Enabled = True
    LoadFUNBFile = False
    Resume Xit

End Function

    
Private Sub PerformLoad()

On Error GoTo PerformLoadErr

    '2/16/2014 - Replaced FileSystemObject with FSO Object
    'Dim fso As New FileSystemObject
    Dim FSO As New FSO
    
    Dim sEffectiveDate As String
    Dim sPath As String
    Dim sFileName As String
    Dim rsWork As New ADODB.Recordset
    Dim bFinished As Boolean
    Dim bDetailNeeded As Boolean
    Dim sDetailName As String
    '2/14/2014 - Converted Dictionary to Collection
    Dim colDetailRecs As New Collection
    Dim cnDetails As New ADODB.Connection
    Dim rsDetailRecs As New ADODB.Recordset
    Dim sTempName As String
    Dim sFUNBFile As String
    
    rsWork.Open "SELECT * FROM DD_WORK_FILE", gcnDDS, adOpenForwardOnly
    If Not rsWork.EOF Then
        MsgBox "The previous validation process ended with errors." & vbCrLf & "You must run Validate Transaction before loading another file.", vbInformation
        Set rsWork = Nothing
        Exit Sub
    End If

    Set rsWork = Nothing
    
    'Lock the Application so no other users can load at the same time
    If LockApplication = False Then
        lblEditStatus.Caption = "Another user is currently loading or validating Direct Deposit Data." & vbCrLf & "Try loading data at a later time."
        Hourglass False
        cmdLoad.Enabled = True
        Command1.Enabled = True
        Exit Sub
    End If
    
    Do Until bFinished = True
        'Provide a dialog box to have the user select the bai2 file
        fMainForm.dlgCommonDialog.DialogTitle = "Select FUNB File"
        fMainForm.dlgCommonDialog.InitDir = GetSetting(App.EXEName, "Settings", "LoadInitDir", "")
        fMainForm.dlgCommonDialog.CancelError = True
        fMainForm.dlgCommonDialog.FileName = ""
        fMainForm.dlgCommonDialog.Filter = "BAI Files|*.bai2|BAI Text Files|*.txt"
        On Error Resume Next
        fMainForm.dlgCommonDialog.ShowOpen
        If Err > 0 Then
            Hourglass False
            cmdLoad.Enabled = True
            Command1.Enabled = True
            UnlockApplication
            Exit Sub
        Else
            'Provide a dialog box for the user to select the detail file
            sFUNBFile = fMainForm.dlgCommonDialog.FileName
            If sFUNBFile <> vbNullString Then
                sDetailName = vbNullString
                Hourglass True
                SeparatePathAndFileName sFUNBFile, sPath, sFileName
                fMainForm.dlgCommonDialog.DialogTitle = "Select Detail File Name"
                fMainForm.dlgCommonDialog.InitDir = GetSetting(App.EXEName, "Settings", "LoadInitDir", "")
                fMainForm.dlgCommonDialog.CancelError = True
                fMainForm.dlgCommonDialog.FileName = ""
                fMainForm.dlgCommonDialog.Filter = "BAI CSV Files|*.csv|"
                On Error Resume Next
                fMainForm.dlgCommonDialog.ShowOpen
                If Err > 0 Then
                    Hourglass False
                    cmdLoad.Enabled = True
                    Command1.Enabled = True
                    UnlockApplication
                    Exit Sub
                Else
                    sDetailName = fMainForm.dlgCommonDialog.FileName
                End If
                On Error GoTo PerformLoadErr
                'Check to see if we have a detail file
                If FileExists(sDetailName) = False Then
                    Hourglass False
                    cmdLoad.Enabled = True
                    Command1.Enabled = True
                    UnlockApplication
                    Exit Sub
                Else
                    'The user has selected a bai file and a detail file
                    sTempName = "TempBaiDetail" & gobjLoginInfo.UserId & ".csv"
                    If FileExists(gsDataPath & "\" & sTempName) = True Then
                        FSO.DeleteAFile gsDataPath & "\" & sTempName
                    End If
                    FSO.CopyAFile sDetailName, gsDataPath & "\" & sTempName
                    'The detail file is in the form of a csv file. Create a connection to allow program to read the detail file as a recordset
                    cnDetails.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                    "Data Source=" & gsDataPath & ";" & _
                    "Extended Properties=""Text;HDR=Yes;"""
                    'Open the detail recordset and if there are no detail generate an error
                    rsDetailRecs.Open "Select * from " & sTempName, cnDetails, adOpenStatic, adLockReadOnly, adCmdText
                    If rsDetailRecs.RecordCount = 0 Then
                        Hourglass False
                        cmdLoad.Enabled = True
                        Command1.Enabled = True
                        UnlockApplication
                        MsgBox "Detail File Produced no valid entries"
                        Exit Sub
                    End If
                End If
            Else
                Hourglass False
                cmdLoad.Enabled = True
                Command1.Enabled = True
                UnlockApplication
            End If
            iPic = 0
            Toggle = 1
            lblEditStatus.Caption = "Loading transactions." & vbCrLf & "Please wait until the load process is complete..."
            cmdLoad.Enabled = False
            Command1.Enabled = False
            cmdCancel.Enabled = False
            'We will now load the bai file into the dds database
            If LoadFUNBFile(rsDetailRecs, sPath, sFileName, sDetailName) = False Then
                cmdLoad.Enabled = True
                Command1.Enabled = True
                UnlockApplication
                Exit Sub
            End If
'        Else
'            Hourglass False
'            cmdLoad.Enabled = True
'            Command1.Enabled = True
'            UnlockApplication
'            Exit Sub
        End If
        
        'If MsgBox("Would you like to load another file before validation", vbYesNo + vbQuestion) = vbNo Then
            bFinished = True
        'End If
    Loop

    Select Case giProcess
    Case LOAD_AND_VALIDATE, LOAD_VALIDATE_AND_POST
        ValidateRecords
        MoveZeroToHidden
        UnlockApplication
    End Select
    

Xit:
Set FSO = Nothing

    Exit Sub

PerformLoadErr:
    ShowUnexpectedError MODULE + "PerformLoad", Err
    Resume


End Sub

Private Sub MoveZeroToHidden()
On Error Resume Next
    Dim cmdTemp As New ADODB.Command

    cmdTemp.ActiveConnection = gcnDDS
    If gStoredProcs("up_u_MoveZeroToHidden").GetStoredProcCommand(cmdTemp) = False Then
        'do Nothing
    Else
        cmdTemp.Parameters("user_id") = gobjLoginInfo.UserId
        cmdTemp.Execute
    End If
    Set cmdTemp = Nothing

End Sub



Private Sub OutlookTitle1_IconClick()
    If cmdCancel.Enabled = True Then
        Unload Me
    End If

End Sub
