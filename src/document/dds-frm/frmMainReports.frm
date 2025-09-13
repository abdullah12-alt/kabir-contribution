VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{BB3B26D0-99DF-11D2-9C22-00105A19BCF2}#8.0#0"; "DatePicker.ocx"
Object = "{8CD222DF-7752-11D3-9D1E-00105A19BCF2}#1.0#0"; "OAOTBar.ocx"
Begin VB.Form frmMainReports 
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   9930
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   6090
      TabIndex        =   4
      Top             =   6495
      Width           =   1095
   End
   Begin VB.CommandButton cmdEmail 
      Caption         =   "&Email"
      Height          =   375
      Left            =   4875
      TabIndex        =   3
      Top             =   6495
      Width           =   1095
   End
   Begin OAOTitleBar.OutlookTitleBar OutlookTitle1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   979
      ForeColor       =   16777215
      Caption         =   "Reports"
   End
   Begin VB.Frame fraInquiryCriteria 
      Caption         =   "Inquiry Criteria"
      Height          =   2415
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Width           =   9255
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sdcInstitution 
         DataSource      =   "adcInstitution"
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   2415
         DataFieldList   =   "Column 2"
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
         ForeColorEven   =   0
         BackColorOdd    =   8454143
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "INSTITUTION_ID"
         Columns(0).Name =   "INSTITUTION_ID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2619
         Columns(1).Caption=   "Institution Code"
         Columns(1).Name =   "INSTITUTION_CODE"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Caption=   "Institution Name"
         Columns(2).Name =   "INSTITUTION_NAME"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin DatePicker.DateSelector dtStart 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   1320
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   -2147483643
         CalendarForeColor=   -2147483630
         CalendarTitleBackColor=   -2147483633
         CalendarTitleForeColor=   -2147483630
         CalendarTrailingForeColor=   -2147483631
         MaxDate         =   2958465
         MinDate         =   -36522
      End
      Begin VB.Label lblInquiryStart 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date:"
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label lblInstitution 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Institution:"
         Height          =   345
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   6495
      Width           =   1095
   End
   Begin VB.CommandButton cmdReportPreview 
      Caption         =   "Pre&view"
      Default         =   -1  'True
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   6495
      Width           =   1095
   End
   Begin VB.Frame Frame11 
      Caption         =   "Choose a Report"
      Height          =   3255
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Width           =   9255
      Begin VB.TextBox lblReportDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   1440
         Left            =   255
         MultiLine       =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1635
         Width           =   5385
      End
      Begin VB.ListBox lstMainReports 
         Height          =   1035
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   5415
      End
      Begin VB.Image imgReport 
         Height          =   1575
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   1455
         Left            =   7320
         Picture         =   "frmMainReports.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmMainReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
' * Form Name:frmMainReports
' * Form File Name: MainReports.frm
' * Start Date: 5/5/1999
' * End Date:
' * Description:
' * --------------------------------
' * The REPORT MAIN Screen is used to display reports for the HEART-DDS Project.
' *  This screen serves 4 primary reports
' *  1) Deceased Exception Report
' *  2) Pre-Edit Report
' *  3) Direct Deposits Detail and Summary Report
' *  4) Automated Balancing Worksheet Report"
'
'
' Mod CONSTANTS
'
'
Option Explicit
Private Const MODULE As String = "Main Report Screen - "
Const CReport1 = "Deceased Exception Report"
Const CReport2 = "Pre-Edit Report"
Const CReport3 = "Direct Deposits Detail and Summary Report"
Const CReport4 = "Automated Balancing Worksheet Report"
Const CReport5 = "Affinity Post Verification Report"


'********************************************************************************
'* Name:  cmdClose()
'*
'* Description: close the Report Sceen and return to the Main Screen
'* Parameters:
'* Created:  5/5/99
'*
'********************************************************************************
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEmail_Click()
  
    
'********************************************************************************
'* Name:  SendEmail()
'* Description:  Creates a rich text file for every applicable institution
'*  then it will send the file via email to the proper person
'* Created: 5/22/2000
'********************************************************************************

On Error GoTo cmdEmailErr

Dim rsInst As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim iRptNo As Integer
Dim sSql As String

    cmdEmail.Enabled = False
    cmdPrint.Enabled = False
    cmdReportPreview.Enabled = False
    Hourglass True
    Select Case lstMainReports.Text
    Case CReport1
        iRptNo = 1
    Case CReport3
        iRptNo = 3
    Case Else
        GoTo Xit
    End Select
    

    If sdcInstitution.Text = "All Institutions" Then
        'Find which institutions need a report
        sSql = "SELECT INSTITUTION_CODE, INSTITUTION_NAME,DD_SEND_REPORT_TO"
        sSql = sSql & " FROM PF_INSTITUTION"
        sSql = sSql & " WHERE RECORD_STATUS = 'A'"
        rsInst.Open sSql, gcnPFS, adOpenStatic
        
        Select Case iRptNo
        Case 1
            sSql = "SELECT INSTITUTION_CODE FROM DD_INVALID_REC"
            sSql = sSql & " WHERE DECEASED_IND = 'Y' AND RECORD_STATUS = 'A'"
            sSql = sSql & " GROUP BY INSTITUTION_CODE"
        Case 3
            sSql = "SELECT INSTITUTION_CODE FROM DD_POSTING_HISTORY"
            sSql = sSql & " WHERE POSTED_DATETIME >= '" & dtStart.Text & " 00:00:00'"
            sSql = sSql & " AND POSTED_DATETIME <= '" & dtStart.Text & " 23:59:59'"
            sSql = sSql & " GROUP BY INSTITUTION_CODE"
        End Select
        
        rsTemp.Open sSql, gcnDDS, adOpenForwardOnly
        Do Until rsTemp.EOF
            rsInst.MoveFirst
            rsInst.Find "INSTITUTION_CODE = '" & rsTemp!INSTITUTION_CODE & "'"
            If Not rsInst.EOF Then
                If IsNull(rsInst!DD_SEND_REPORT_TO) Then
                    MsgBox "The report for " & rsInst!INSTITUTION_NAME & " is not set to be sent via email", vbInformation + vbOKOnly
                Else
                    SendEmailReport iRptNo, rsInst!INSTITUTION_CODE, rsInst!DD_SEND_REPORT_TO
                End If
            End If
            rsTemp.MoveNext
        Loop
        
    Else
        'Find the institution need a report
        sSql = "SELECT INSTITUTION_CODE , INSTITUTION_NAME, DD_SEND_REPORT_TO"
        sSql = sSql & " FROM PF_INSTITUTION"
        sSql = sSql & " WHERE INSTITUTION_CODE = '" & sdcInstitution.Columns(1).Text & "' AND RECORD_STATUS = 'A'"
        rsInst.Open sSql, gcnPFS, adOpenStatic
        If Not rsInst.EOF Then
            If IsNull(rsInst!DD_SEND_REPORT_TO) Then
                MsgBox "The report for " & rsInst!INSTITUTION_NAME & " is not set to be sent via email", vbInformation + vbOKOnly
            Else
                SendEmailReport iRptNo, rsInst!INSTITUTION_CODE, rsInst!DD_SEND_REPORT_TO
            End If
        End If
    End If
    MsgBox "Finished emailing report.", vbInformation
Exit Sub

Xit:
    Set rsInst = Nothing
    Set rsTemp = Nothing
    cmdEmail.Enabled = True
    cmdReportPreview.Enabled = True
    cmdPrint.Enabled = True

Exit Sub

cmdEmailErr:
If Err = 32050 Then
    Resume Next
Else
    MsgBox Error
    Resume Xit
End If

End Sub



Private Sub SendEmailReport(ByVal iRptNo As Integer, sInstCode As String, ByVal sSendTo)

On Error GoTo SendEmailReportErr

Dim crRpt As CRAXDRT.Report
Dim sCriteria As String
Dim rs As New ADODB.Recordset
Dim oEmailMsg As New clsEmailMessage
Dim bRet As Boolean
Dim sWordDoc As String
Dim sSql As String

Hourglass True
oEmailMsg.From = ReadIniFile(App.Path & "\dds.ini", "Startup", "SMTPuser")
'
'Get the report
Select Case iRptNo
Case 1
    Set crRpt = gcrApp.OpenReport(gsDataPath & "\deceased.rpt", crOpenReportByTempCopy)
    oEmailMsg.Subject = "Direct Deposit Deceased Report"
    oEmailMsg.Message = "Attached you will find the Direct Deposit Deceased Report"
    crRpt.ExportOptions.DiskFileName = App.Path & "\deceased" & sInstCode & ".rtf"
    sWordDoc = App.Path & "\deceased" & sInstCode & ".doc"
    'Create the criteria
    sSql = "SELECT DD_INVALID_REC.*,PF_INSTITUTION.INSTITUTION_NAME As INSTITUTION_NAME"
    sSql = sSql & " FROM dds.dbo.DD_INVALID_REC DD_INVALID_REC,pfs.dbo.PF_INSTITUTION PF_INSTITUTION"
    sSql = sSql & " WHERE DD_INVALID_REC.INSTITUTION_CODE = PF_INSTITUTION.INSTITUTION_CODE"
    sSql = sSql & " AND DD_INVALID_REC.DECEASED_IND = 'Y'"
    sSql = sSql & " AND DD_INVALID_REC.RECORD_STATUS = 'A'"
    If sInstCode <> "" Then
        sSql = sSql & " AND DD_INVALID_REC.INSTITUTION_CODE = '" & sInstCode & "'"
    End If
    rs.Open sSql, gcnDDS, adOpenStatic
    
    crRpt.Database.SetDataSource rs
    crRpt.ReadRecords
Case 3
    Set crRpt = gcrApp.OpenReport(gsDataPath & "\DDDetSum.rpt", crOpenReportByTempCopy)
    oEmailMsg.Subject = "Direct Deposit Detail and Summary Report"
    oEmailMsg.Message = "Attached you will find the Direct Deposit Detail and Summary Report"
    crRpt.ExportOptions.DiskFileName = App.Path & "\DDDetSum" & sInstCode & ".rtf"
    sWordDoc = App.Path & "\DDDetSum" & sInstCode & ".doc"
    'Create the criteria
    sSql = "SELECT DD_POSTING_HISTORY.INSTITUTION_CODE, DD_POSTING_HISTORY.AFFINITY_ACCT_NUM, DD_POSTING_HISTORY.MEDICAL_RECORD_NUM, DD_POSTING_HISTORY.DD_NUM, DD_POSTING_HISTORY.TOT_FUNB_BENEFIT_AMT, DD_POSTING_HISTORY.DR_CR_FLAG, DD_POSTING_HISTORY.AS_OF_DATETIME, DD_POSTING_HISTORY.PA_DISTRIBUTION_AMT, DD_POSTING_HISTORY.PF_DISTRIBUTION_AMT, DD_POSTING_HISTORY.PATIENT_NAME, DD_POSTING_HISTORY.POSTED_DATETIME, DD_POSTING_HISTORY.TOT_DAYS_INHOUSE, DD_POSTING_HISTORY.SPEC_PROC_COND_HASH_TOT, DD_POSTING_HISTORY.ATP_PML_FLAG,DD_INCOME_SOURCE_TYPE.INCOME_SRC_TYPE_DESCR"
    sSql = sSql & " FROM DD_POSTING_HISTORY,DD_INCOME_SOURCE_TYPE"
    sSql = sSql & " WHERE DD_POSTING_HISTORY.INCOME_SOURCE_TYPE_ID = DD_INCOME_SOURCE_TYPE.INCOME_SOURCE_TYPE_ID"
    sSql = sSql & " AND DD_POSTING_HISTORY.INSTITUTION_CODE = '" & sInstCode & "'"
    sSql = sSql & " AND DD_POSTING_HISTORY.POSTED_DATETIME >= '" & dtStart.Text & " 00:00:00'"
    sSql = sSql & " AND DD_POSTING_HISTORY.POSTED_DATETIME <= '" & dtStart.Text & " 23:59:59'"
    sSql = sSql & Chr$(13) + Chr$(10) & "ORDER BY DD_POSTING_HISTORY.INSTITUTION_CODE Asc"
    rs.Open sSql, gcnDDS, adOpenStatic
    
    crRpt.ParameterFields.Item(1).AddCurrentValue "Posted on " & dtStart.Text
    
    crRpt.Database.SetDataSource rs
    'If rs.EOF Then
    '    MsgBox "Nothing was posted for this date"
    '    GoTo Xit
    'End If
    
    crRpt.ReadRecords
    
End Select


'Set up the email parameters
crRpt.ExportOptions.DestinationType = crEDTDiskFile
'crRpt.EMailToList = sSendTo
crRpt.ExportOptions.FormatType = crEFTRichText

'Trigger the event at run time
Sleep 1
crRpt.Export False
oEmailMsg.ToRecipient = sSendTo
bRet = ConvertRTFToWordDoc(crRpt.ExportOptions.DiskFileName, True, sWordDoc)
If bRet = True Then
    If oEmailMsg.AddEmailAttachment(sWordDoc) = False Then
        Err.Raise 834543, , "Error attaching file to email"
    End If
    fMainForm.SendEmail oEmailMsg
End If

'Reset the Mouse Pointer to the default
Xit:

Hourglass False
Exit Sub

SendEmailReportErr:
    ShowError MODULE + "Preview Report", Err
    Resume Xit

End Sub

'Private Sub EmailFile(ByVal sPrintFileName As String, ByVal sRecipEmail As String, ByVal sSubject As String, ByVal sText As String)
'
'Dim sPathName As String
'Dim sFileName As String
'Dim iStart As Integer
'Dim iEnd As Integer
'Dim sPartAddr As String
''Open up a registry object
'
'On Error GoTo EMailFileErr
'
'    With MAPIMessages1
'
'   .SessionID = MAPISession1.SessionID
'   .Compose
'   .MsgIndex = -1
'   'address the message
'   'Separate any commas or semicolon in the RecipEmail
'
'    iStart = 1
'    iEnd = InStr(1, sRecipEmail, ",")
'    If iEnd = 0 Then
'        .RecipType = 1
'        .RecipAddress = sRecipEmail
'        'cc: addresses
'    Else
'        Do Until iEnd = 0
'            sPartAddr = Mid$(sRecipEmail, iStart, iEnd - iStart)
'            .RecipIndex = .RecipCount
'            .RecipType = 1
'            .RecipAddress = sPartAddr
'            'cc: addresses
'            iStart = iEnd + 1
'            iEnd = InStr(iStart, sRecipEmail, ",")
'            If iEnd = 0 Then
'                sPartAddr = Mid$(sRecipEmail, iStart, Len(sRecipEmail) - iStart + 1)
'                .RecipIndex = .RecipCount
'                .RecipType = 1
'                .RecipAddress = sPartAddr
'            End If
'        Loop
'    End If
'   .MsgSubject = sSubject
'   .MsgNoteText = sText
'   'define your attatchments
'    'Attachment # 1
'   .AttachmentIndex = .AttachmentCount
'   .AttachmentType = mapData
'    SeparatePathAndFileName sPrintFileName, sPathName, sFileName
'   .AttachmentName = sFileName
'   .AttachmentPathName = sPrintFileName
'   .AttachmentPosition = 1
'
'   'send the message
'   .send True
'    Sleep 5
'
'End With
'
'SendEmail = True
'
'Exit Sub
'EMailFileErr:
'    ShowError MODULE + "Preview Report", Err
'
'End Sub

Private Sub cmdPrint_Click()
'Error handler
On Error GoTo cmdPrintErr
Dim IntRet As Integer
miPrintMode = 1
    Select Case lstMainReports.Text
       Case CReport1
            'call a procedure called Report1 for
            ' *  1) Deceased Exception Report
            ViewPrintReport 1, sdcInstitution.Columns(1).Text
            
       Case CReport2
            'call a procedure called Report2 for
            ' *  2) Pre-Edit Report
            ViewPrintReport 2
            
            'fNewRpt.Show
       Case CReport3
            ' *  3) Direct Deposits Detail and Summary Report
            'Check that there is a date
            If dtStart.Text = vbNullString Then
                MsgBox "Date is required.  Please enter a valid date.", vbExclamation
                Hourglass False
                dtStart.SetFocus
                Exit Sub
            End If
            ViewPrintReport 3, sdcInstitution.Columns(1).Text, dtStart.Text
       
       Case CReport4
            'call a procedure called Report4 for
            ' *  4) Automated Balancing Worksheet Report"
            
            'Check that there is a date
            If dtStart.Text = vbNullString Then
                MsgBox "Date is required.  Please enter a valid date.", vbExclamation
                Hourglass False
                dtStart.SetFocus
                Exit Sub
            End If
            
            ViewPrintReport 4, , dtStart.Text
         
       Case CReport5
            ''call a procedure called Report5 for
            '' *  5) Affinity Post Verification Report"
            'Call REPORT5
        
       Case Else
            'should never happen since command button is not enabled yet.
            Beep
            IntRet = MsgBox("Please select a report before you click Preview.", vbInformation, "Report Preview Message")
    
    End Select

Xit:
      Exit Sub
cmdPrintErr:
    ShowError MODULE + "cmdPrintErr", Err
    Resume Xit

End Sub

'********************************************************************************
'* Name: cmdReportPreview_Click()
'*
'* Description: Sets the choices  for the Report when Preview Command Button is clicked
'* Parameters:
'* Created: 6/5/99
'********************************************************************************
Private Sub cmdReportPreview_Click()
'Error handler
On Error GoTo ReportPreviewErr
Dim IntRet As Integer
miPrintMode = 0
    Select Case lstMainReports.Text
       Case CReport1
            'call a procedure called Report1 for
            ' *  1) Deceased Exception Report
            ViewPrintReport 1, sdcInstitution.Columns(1).Text
            
       Case CReport2
            'call a procedure called Report2 for
            ' *  2) Pre-Edit Report
            ViewPrintReport 2
            
            'fNewRpt.Show
       Case CReport3
            ' *  3) Direct Deposits Detail and Summary Report
            'Check that there is a date
            If dtStart.Text = vbNullString Then
                MsgBox "Date is required.  Please enter a valid date.", vbExclamation
                Hourglass False
                dtStart.SetFocus
                Exit Sub
            End If
            ViewPrintReport 3, sdcInstitution.Columns(1).Text, dtStart.Text
       
       Case CReport4
            'call a procedure called Report4 for
            ' *  4) Automated Balancing Worksheet Report"
            
            'Check that there is a date
            If dtStart.Text = vbNullString Then
                MsgBox "Date is required.  Please enter a valid date.", vbExclamation
                Hourglass False
                dtStart.SetFocus
                Exit Sub
            End If
            
            ViewPrintReport 4, , dtStart.Text
         
       Case CReport5
            ''call a procedure called Report5 for
            '' *  5) Affinity Post Verification Report"
            'Call REPORT5
        
       Case Else
            'should never happen since command button is not enabled yet.
            Beep
            IntRet = MsgBox("Please select a report before you click Preview.", vbInformation, "Report Preview Message")
    
    End Select

Xit:
      Exit Sub
ReportPreviewErr:
    ShowError MODULE + "Preview Report", Err
    Resume Xit
End Sub


Private Sub Command1_Click()
Dim printDlg As PrinterDlg
Set printDlg = New PrinterDlg
' Set the starting information for the dialog box based on the current
' printer settings.
printDlg.PrinterName = Printer.DeviceName
printDlg.DriverName = Printer.DriverName
printDlg.Port = Printer.Port

' Set the default PaperBin so that a valid value is returned even
' in the Cancel case.
printDlg.PaperBin = Printer.PaperBin

' Set the flags for the PrinterDlg object using the same flags as in the
' common dialog control. The structure starts with VBPrinterConstants.
printDlg.Flags = VBPrinterConstants.cdlPDNoSelection _
                 Or VBPrinterConstants.cdlPDReturnDC
Printer.TrackDefault = False
'                 Or VBPrinterConstants.cdlPDNoPageNums _

printDlg.Min = 1
printDlg.Max = 3
printDlg.FromPage = 1
printDlg.ToPage = 3


' When CancelError is set to True the ShowPrinterDlg will return error
' 32755. You can handle the error to know when the Cancel button was
' clicked. Enable this by uncommenting the lines prefixed with "'**".
'**printDlg.CancelError = True

' Add error handling for Cancel.
'**On Error GoTo Cancel
If Not printDlg.ShowPrinter(Me.hwnd) Then
    Debug.Print "Cancel Selected"
    Exit Sub
End If

'Turn off Error Handling for Cancel.
'**On Error GoTo 0
Dim NewPrinterName As String
Dim objPrinter As Printer
Dim strsetting As String

' Locate the printer that the user selected in the Printers collection.
NewPrinterName = UCase$(printDlg.PrinterName)
If Printer.DeviceName <> NewPrinterName Then
    For Each objPrinter In Printers
       If UCase$(objPrinter.DeviceName) = NewPrinterName Then
            Set Printer = objPrinter
       End If
    Next
End If

' Copy user input from the dialog box to the properties of the selected printer.
Printer.Copies = printDlg.Copies
Printer.Orientation = printDlg.Orientation
Printer.ColorMode = printDlg.ColorMode
Printer.Duplex = printDlg.Duplex
Printer.PaperBin = printDlg.PaperBin
Printer.PaperSize = printDlg.PaperSize
Printer.PrintQuality = printDlg.PrintQuality

' Display the results in the immediate (Debug) window.
' NOTE: Supported values for PaperBin and Size are printer specific. Some
' common defaults are defined in the Win32 SDK in MSDN and in Visual Basic.
' Print quality is the number of dots per inch.
With Printer
    Debug.Print .DeviceName
    If .Orientation = 1 Then
        strsetting = "Portrait. "
    Else
        strsetting = "Landscape. "
    End If
    Debug.Print "Copies = " & .Copies, "Orientation = " & _
       strsetting
    If .ColorMode = 1 Then
        strsetting = "Black and White. "
    Else
        strsetting = "Color. "
    End If
    Debug.Print "ColorMode = " & strsetting
    If .Duplex = 1 Then
        strsetting = "None. "
    ElseIf .Duplex = 2 Then
        strsetting = "Horizontal/Long Edge. "
    ElseIf .Duplex = 3 Then
        strsetting = "Vertical/Short Edge. "
    Else
        strsetting = "Unknown. "
    End If
    Debug.Print "Duplex = " & strsetting
    Debug.Print "PaperBin = " & .PaperBin
    Debug.Print "PaperSize = " & .PaperSize
    Debug.Print "PrintQuality = " & .PrintQuality
    If (printDlg.Flags And VBPrinterConstants.cdlPDPrintToFile) = _
       VBPrinterConstants.cdlPDPrintToFile Then
         Debug.Print "Print to File Selected"
    Else
         Debug.Print "Print to File Not Selected"
    End If
    Debug.Print "hDC = " & printDlg.hDC
End With
Exit Sub
'**Cancel:
'**If Err.Number = 32755 Then
'**    Debug.Print "Cancel Selected"
'**Else
'**    Debug.Print "A nonCancel Error Occured - "; Err.Number
'**End If

End Sub

Private Sub Command2_Click()
Dim bRet As Boolean
Dim sDocName As String
    bRet = ConvertRTFToWordDoc("H:\dds\deceased1.rtf", True, sDocName)
    MsgBox bRet
    
End Sub

Private Function ConvertRTFToWordDoc(ByVal sRTFFileName As String, ByVal bProtect As Boolean, ByRef sWordDocName As String)
    
On Error GoTo ConvertRTFToWordDocErr

    Dim wordApp As Word.Application
    Set wordApp = CreateObject("Word.Application")
    Dim wordDoc As Word.Document
    
    If LCase(Right$(sRTFFileName, 4)) = ".rtf" Then
        sWordDocName = Left$(sRTFFileName, Len(sRTFFileName) - 4) & ".doc"
    Else
        Err.Raise 14234, "Error Converting RTF File"
        
    End If
    
    If Dir(sWordDocName) <> "" Then
        Kill sWordDocName
    End If
    
    Set wordDoc = wordApp.Documents.Open(sRTFFileName)
    If bProtect = True Then
        wordDoc.SaveAs sWordDocName, wdFormatDocument, , "hearts1"
    Else
        wordDoc.SaveAs sWordDocName, wdFormatDocument
    End If
    wordDoc.Close
    ConvertRTFToWordDoc = True
    
Xit:
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Exit Function

ConvertRTFToWordDocErr:
    MsgBox Error, vbCritical

End Function

Private Sub dtStart_Change()
    DetermineShowButtons
End Sub

Private Sub dtStart_CloseUp(Cancel As Boolean)

    DetermineShowButtons
    
End Sub

''********************************************************************************
''* Name: dtStart_DropDown()
''* Description: To turn the Preview command on since we have a single date drop
''* Parameters:
''* Created: ----  7/11/99
''********************************************************************************
''
''
'Private Sub dtStart_DropDown(Cancel As Boolean)
' cmdReportPreview.Enabled = True
'If lstMainReports.Text = CReport3 Then
'    cmdEmail.Enabled = True
'End If
'End Sub

'********************************************************************************
'* Name: dtStart_lostFocus()
'* Description: Date Validation Process if lost focus on the DateEnd object
'* Parameters:
'* Created: ----  6/16/99
'********************************************************************************
Private Sub dtStart_lostFocus()
On Error GoTo dtStart_LostFocusError
If Len(dtStart.Text) > 0 Then
        If IsDate(dtStart.Text) And Year(dtStart.Text) > 1799 Then
            dtStart.Text = Format(dtStart.Text, "mm/dd/yyyy")
        Else
            MsgBox "The date entered is an invalid date.  Please enter a valid start date.", vbExclamation
            dtStart.Text = Format(Now, "mm/dd/yyyy")
            dtStart.SetFocus
        End If
End If
 cmdReportPreview.Enabled = True
 If lstMainReports.Text = CReport3 Then
    cmdEmail.Enabled = True
 End If
Xit:
    Exit Sub
dtStart_LostFocusError:
    ShowUnexpectedError MODULE + "dtStart_LostFocus", Err
    Resume Xit
End Sub

'********************************************************************************
'* Name: dtStart_Validate(Cancel As Boolean)
'*
'* Description: Date Validation Process for Date Start Object
'* Parameters:  Cancel
'* Created: ----  6/16/99
'********************************************************************************
Private Sub dtStart_Validate(Cancel As Boolean)
    If dtStart.Text <> vbNullString Then
        If Not IsDate(dtStart.Text) Then
            MsgBox "Date entered is not valid.  Please enter date using format MM/DD/YYYY or select it from the drop down calendar.", vbExclamation
            Cancel = True
            dtStart.SetFocus
        ElseIf CDate(dtStart.Text) > Now Then
            MsgBox "Date entered may not be greater than today's date.", vbExclamation
            Cancel = True
            dtStart.Text = Format(Now, "mm/dd/yyyy")
            dtStart.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()

fMainForm.SetMainToolbar True

End Sub

Private Sub Form_Deactivate()

fMainForm.SetMainToolbar False

End Sub

Private Sub Form_Load()

' Error handler
On Error GoTo ReportFormLoadError

Set OutlookTitle1.Picture = fMainForm.imlToolbarIcons.ListImages("Reports").Picture

'Load all ITEMS for the list of REPORTS to the Main Menu LIST Box
'
Call HideAllFrames
Call FilllstMainReports
Call RefreshsdcInstitution

Exit Sub

ReportFormLoadError:
  If (Err.Number <> 0) Then
    MsgBox (Err.Description)
  End If

End Sub

Private Sub lstMainReports_Click()

Select Case lstMainReports.Text
    Case CReport1
        Call ReportDescription1
        ' *  1) Deceased Exception Report
        Call HideAllFrames
        Call ShowFrameReport1
        
    Case CReport2
        Call ReportDescription2
        ' *  2) Pre-Edit Report
        Call HideAllFrames
        Call ShowFrameReport2
        
    Case CReport3
        Call ReportDescription3
        ' *  3) Direct Deposits Detail and Summary Report
        Call HideAllFrames
        Call ShowFrameReport3
        
    Case CReport4
        Call ReportDescription4
        ' *  4) Automated Balancing Worksheet Report
        Call HideAllFrames
        Call ShowFrameReport4
        
    Case CReport5
        Call ReportDescription5
        ' *  5) Affinity Post Verification Report
        Call HideAllFrames
        Call ShowFrameReport5
        
   Case Else
        ' Should never happen
        MsgBox "No valid report was selected", vbOKOnly, "Report Menu"
End Select
   
DetermineShowButtons
If sdcInstitution.Visible Then
    sdcInstitution.SetFocus
    lstMainReports.SetFocus
End If

End Sub

Private Sub DetermineShowButtons()
    cmdEmail.Enabled = False
    cmdReportPreview.Enabled = False
    Select Case lstMainReports.Text
    Case CReport1
        If sdcInstitution.IsItemInList Then
            cmdEmail.Enabled = True
            cmdReportPreview.Enabled = True
        End If
    Case CReport2
        cmdEmail.Enabled = False
        cmdReportPreview.Enabled = True
    Case CReport3
        If sdcInstitution.IsItemInList Then
            If IsDate(dtStart.Text) Then
                cmdEmail.Enabled = True
                cmdReportPreview.Enabled = True
            End If
        End If
    Case CReport4, CReport5
        cmdEmail.Enabled = False
        If IsDate(dtStart.Text) Then
            cmdReportPreview.Enabled = True
        End If
    Case Else
        cmdEmail.Enabled = False
        cmdReportPreview.Enabled = False
    End Select

End Sub
'********************************************************************************
'* Name: FilllstMainReports()
'*
'* Description: Fills the Report listbox by Reports Names
'* Parameters:
'* Created: 5/5/1999
'* Modified:
'********************************************************************************
Public Sub FilllstMainReports()

On Error GoTo FillLstMainReportErr

With lstMainReports
        .AddItem CReport1
        .AddItem CReport2
        .AddItem CReport3
        .AddItem CReport4
'        .AddItem CReport5
End With

Xit:
Exit Sub

FillLstMainReportErr:
ShowError MODULE + "Report", Err
Resume Xit

End Sub
'This function is to Check if it is Null value for the field
Public Function ConvertNull(ByVal vValue As Variant) As Variant
    If IsNull(vValue) Then
        ConvertNull = vbNullString
    Else
        ConvertNull = vValue
    End If
End Function
Public Sub ReportDescription1()
'Report1 = "Deceased Exception Report"
lblReportDescription = " "
lblReportDescription = "This report is generated as a result of the pre-edit process and " & vbNewLine & _
                               "identifies accounts that are discharged due to death. Report reflects" & vbNewLine & _
                               "cumulative records to date that have not been resolved.  This report" & vbNewLine & _
                               "is sorted by institution, by income source, by FUNB 'as of' date, " & vbNewLine & _
                               "and then by wage earner claim number."
'Clear the Image
imgReport.Picture = LoadPicture()
'Load new Image
Call ImageReport1

End Sub
Public Sub ReportDescription2()
'Report2 = "Pre-Edit Report"
lblReportDescription = " "
lblReportDescription = "This report is generated as a result of the pre-edit process and " & vbNewLine & _
                               "identifies accounts that did not pass the validation process. Report" & vbNewLine & _
                               "reflects cumulative records to date that have not been resolved.  This" & vbNewLine & _
                               "report is sorted by FUNB 'as of' date."
'Clear the Image
imgReport.Picture = LoadPicture()
'Load new Image
Call ImageReport2

End Sub

Public Sub ReportDescription3()
'Report3 = "Direct Deposits Detail and Summary Report"
lblReportDescription = " "
lblReportDescription = "This report is generated after the posting process is complete and " & vbNewLine & _
                               "identifies those accounts that resulted in an actual posting to the" & vbNewLine & _
                               "personal fund and/or patient account on the date selected. " & vbNewLine & _
                               "This report is sorted by institution, ATP or PML, income source, " & vbNewLine & _
                               "client name, and FUNB 'as of' date."
'Clear the Image
imgReport.Picture = LoadPicture()
'Load new Image
Call ImageReport3

End Sub

Public Sub ReportDescription4()
'Report4 = "Automated Balancing Worksheet Report"
lblReportDescription = " "
lblReportDescription = "This report provides the Central Billing Office with information" & vbNewLine & _
                               "for a point in time that documents that all areas are in balance," & vbNewLine & _
                               "or information documenting that all areas are NOT in balance so" & vbNewLine & _
                               "that the appropriate follow-up can occur."
'Clear the Image
imgReport.Picture = LoadPicture()
'Load new Image
Call ImageReport4

End Sub


Public Sub ReportDescription5()
'Report5 = "Affinity Post Verification Report"
lblReportDescription = " "
lblReportDescription = "This report provides the Central Billing Office with a list of transactions posted from Direct Deposit that currently are not transactions in Affinity.  If payments were not created in Affinity, steps should be taken to determine what happpened."
'Clear the Image
imgReport.Picture = LoadPicture()
'Load new Image
Call ImageReport5

End Sub

Private Sub ImageReport1()
'Report1 = "Deceased Exception Report"
Dim i As Integer
Randomize
i = Int(5 * Rnd()) + 1
'Load new Image
Select Case i
    Case 1
        imgReport.Picture = LoadPicture()
         Set imgReport.Picture = fMainForm.ImageList1.ListImages("rotateskull").Picture
    Case 2
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Death").Picture
    Case 3
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Devil").Picture
    Case 4
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Skull").Picture
    Case 5
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Eye").Picture
End Select

End Sub
Private Sub ImageReport2()
'Report2 = "Pre-Edit Report"
Dim i As Integer
Randomize
i = Int(2 * Rnd()) + 1
'Load new Image
Select Case i
    Case 1
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Notes").Picture
       
    Case 2
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Yes2b").Picture

    Case Else
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Comments1").Picture

End Select
End Sub
Private Sub ImageReport3()
'Report3 = "Direct Deposits Detail and Summary Report"
Dim i As Integer
Randomize
i = Int(7 * Rnd()) + 1
'Load new Image
Select Case i
    Case 1
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("monBag").Picture
    Case 2
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Money").Picture
    Case 3
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Signat").Picture
    Case 4
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("CkCard").Picture
    Case 5
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Cashbook").Picture
    Case 6
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Moneybut").Picture
    Case 7
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Safe").Picture
        
    Case Else
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Direct-Deposit").Picture

End Select
End Sub
Private Sub ImageReport4()
'Report4 = "Automated Balancing Worksheet Report"
Dim i As Integer
Randomize
i = Int(3 * Rnd()) + 1
'Load new Image
Select Case i
                      
    Case 1
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Signboo").Picture

    Case 2
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Write").Picture

    Case 3
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Yes").Picture

    Case Else
        imgReport.Picture = LoadPicture()
        Set imgReport.Picture = fMainForm.ImageList1.ListImages("Money").Picture

End Select
End Sub

Private Sub ImageReport5()
On Error Resume Next
'Report5 = "Affinity Post Verification Report"
    imgReport.Picture = LoadPicture()
    Set imgReport.Picture = fMainForm.ImageList1.ListImages("Yes").Picture

End Sub

Public Sub HideAllFrames()
fraInquiryCriteria.Visible = False
lblInstitution.Visible = False
sdcInstitution.Visible = False
cmdReportPreview.Enabled = False
cmdEmail.Enabled = False
'Move the label, dtPicker to the original position
lblInquiryStart(0).Top = 1320
dtStart.Top = 1320

End Sub
Private Sub ShowFrameReport1()
On Error Resume Next
'Make Frames of Options visible
'* 1) "Deceased Exception Report

fraInquiryCriteria.Visible = True
lblInstitution.Visible = True
sdcInstitution.Visible = True
lblInquiryStart(0).Visible = False
dtStart.Visible = False

End Sub
Private Sub ShowFrameReport2()
On Error Resume Next

'No Criteria
cmdReportPreview.Enabled = True


End Sub

Private Sub ShowFrameReport3()
On Error Resume Next
'Make frame of options visible
'* 'Report3 = "Direct Deposits Detail and Summary Report"
'
fraInquiryCriteria.Visible = True
lblInstitution.Visible = True
sdcInstitution.Visible = True
lblInquiryStart(0).Visible = True
dtStart.Visible = True
lblInquiryStart(0).Caption = "Date Posted"
'Move the label, dtPicker to top
lblInquiryStart(0).Top = 1320
dtStart.Top = 1320
End Sub
Private Sub ShowFrameReport4()
On Error Resume Next
'Make frame of options visible
'* "Report4- Automated Balancing Worksheet Report"
'
fraInquiryCriteria.Visible = True
lblInstitution.Visible = False
sdcInstitution.Visible = False
lblInquiryStart(0).Visible = True
dtStart.Visible = True
'Move the label, dtPicker to top
lblInquiryStart(0).Top = 1320
lblInquiryStart(0).Caption = "As of Date"
dtStart.Top = 1320
End Sub
Private Sub ShowFrameReport5()
On Error Resume Next
'Make frame of options visible
'* "Report5- Affinty Post Verification Report"
'
fraInquiryCriteria.Visible = True
lblInstitution.Visible = False
sdcInstitution.Visible = False
lblInquiryStart(0).Visible = True
dtStart.Visible = True
'Move the label, dtPicker to top
lblInquiryStart(0).Top = 1320
lblInquiryStart(0).Caption = "As of Date"
dtStart.Top = 1320
End Sub

Private Sub RefreshsdcInstitution()
Dim rs1 As New ADODB.Recordset
Dim sSql As String, sLine As String, sLine1 As String

'Fill list box with active Institutions
sSql = "SELECT INSTITUTION_ID AS 'Institution ID' , INSTITUTION_CODE AS 'Institution Code'," _
     & "INSTITUTION_NAME AS 'Institution Name' FROM PF_INSTITUTION " _
     & "WHERE RECORD_STATUS = 'A' " & _
     "ORDER BY PF_INSTITUTION.INSTITUTION_NAME"

With rs1
.Open sSql, gcnPFS, adOpenForwardOnly, adLockReadOnly, adCmdText
.MoveFirst
    Do Until .EOF
       On Error Resume Next
       sLine = .Fields("Institution ID") & vbTab & .Fields("Institution Code") & vbTab & .Fields("Institution Name")
       
       If Err.Number = 0 Then
            sdcInstitution.AddItem sLine
       Else
            'sdcInstitution.AddItem ""
       End If
       .MoveNext
    Loop
' add "All Institutions" as first item in the List
sLine1 = " " & vbTab & " " & vbTab & "All Institutions"
sdcInstitution.AddItem sLine1, 0
End With

rs1.Close
Set rs1 = Nothing

End Sub



Private Sub OutlookTitle1_IconClick()
    If cmdClose.Enabled = True Then
        Unload Me
    End If

End Sub

Private Sub sdcInstitution_click()

    DetermineShowButtons
  
End Sub
Private Sub sdcInstitutions_Change()

    DetermineShowButtons

End Sub
'
'Private Sub REPORT1()
''********************************************************************************
''* Name:  REPORT1()
''* Description: This is a procedure for report # 1 "Deceased Exception Report"
''*              to launch crystal report # 1 with its criteria By Institution
''* Parameters:
''* Created:
''********************************************************************************
'
'On Error GoTo Report1Err
'
''Report1 Name: Deceased Exception Report
''------------------------------------------------------------
''
''This is a Crystal Report Coding
'Dim crReport1 As Crystal.CrystalReport
'Dim IntRet As Integer
'Dim sCriteria As String
'
'Hourglass True
'
''set to the crystal report object
'Set crReport1 = fMainForm.crReport
'crReport1.Reset
'
''** This is commented out so that users can see all Crystal options
''Call ReportButtons
'
''Open a connection to SQL Server
'crReport1.LogonInfo(0) = gobjLoginInfo.ConnectString
'crReport1.LogonInfo(1) = gobjLoginInfo.PFSConnectString
''
''
''Call the Report file Name
'crReport1.ReportFileName = gsDataPath & "\deceased.rpt"
' 'Check the Criteria
' '
'Select Case sdcInstitution.Text
'   Case "All Institutions"
'        sCriteria = ""
'
'        'turn the group tree on
'        crReport1.WindowShowGroupTree = True
'
'Case Else
'        sCriteria = sCriteria & "{DD_INVALID_REC.INSTITUTION_CODE} = '" & sdcInstitution.Columns(1).Text & "'"
'
'        'turn the group tree off
'        crReport1.WindowShowGroupTree = False
'
'End Select
'
'crReport1.SelectionFormula = sCriteria
'
''Send the report either to be sent to the printer or to screen
'crReport1.Destination = miPrintMode
''
''Maximize the report window
'crReport1.WindowState = crptMaximized
'
''Show a Window Title for the Report
'crReport1.WindowTitle = "Deceased Exception Report"
'
''Trigger the event at run time
'crReport1.Action = 1
'
''Reset the Mouse Pointer to the default
'Hourglass False
'
'Xit:
'    Exit Sub
'Report1Err:
'    Hourglass False
'    ShowError MODULE + "Deceased Exception Report", Err
'    Resume Xit
'End Sub
'
'Private Sub REPORT2()
''********************************************************************************
''* Name:  REPORT2()
''* Description: This is a procedure for report # 2 " Pre-Edit Report" to
''*              launch crystal report # 2 with NO criteria for the Pre-Edit Report
''* Parameters:
''* Created: 7/15/1999
''********************************************************************************
'
'On Error GoTo Report2Err
'
''Report2 Name: Pre-Edit Report
''------------------------------------------------------------
''
''This is a Crystal Report Coding
'
'Dim crReport2 As Crystal.CrystalReport
'Dim IntRet As Integer
'Dim sCriteria As String
'
'Hourglass True
'
''set to the crystal report Object
'Set crReport2 = fMainForm.crReport
'crReport2.Reset
'
''** This is commented out so that users can see all Crystal options
''Call ReportButtons
'
''Open a connection to SQL Server
'crReport2.LogonInfo(0) = gobjLoginInfo.ConnectString
'crReport2.LogonInfo(1) = gobjLoginInfo.PFSConnectString
'
''Call the Report file Name
'crReport2.ReportFileName = gsDataPath & "\PreEdit.rpt"
'
''No selection Criteria per DDS REQUIREMENTS
'crReport2.Destination = miPrintMode
''
''Maximize the report window
'crReport2.WindowState = crptMaximized
'
''Show a Window Title for the Report
'crReport2.WindowTitle = "Pre-Edit Report"
'
''Trigger the event at run time
'crReport2.Action = 1
'
''Reset the Mouse Pointer to the default
'Hourglass False
'
'Xit:
'    Exit Sub
'Report2Err:
'    Hourglass False
'    ShowError MODULE + "Pre-edit Report", Err
'    Resume Xit
'End Sub
' '********************************************************************************
''* Name:  REPORT3()
''* Description: This is a procedure for report # 3 Direct Deposits Detail and Summary Report
''*               to launch crystal report # 3 with its criteria: Single Date
''* Parameters:
''* Created:
''********************************************************************************
'Private Sub REPORT3()
''Report Name: "Direct Deposits Detail and Summary Report"
'
''Error Handler !!
'On Error GoTo Report3Err
'
'Dim crReport3 As Crystal.CrystalReport
'Dim IntRet As Integer
'Dim sCriteria As String
'
'Hourglass True
'
''set to the crystal report
'Set crReport3 = fMainForm.crReport
'crReport3.Reset
'
''** This is commented out so that users can see all Crystal options
'' Call ReportButtons
'
''Open a connection to SQL Server
'crReport3.Connect = gobjLoginInfo.ConnectString
'
''call the Reportfile Name called Direct Deposits Detail and Summary Report
'crReport3.ReportFileName = gsDataPath & "\DDDetSum.rpt"
'
''Check that there is a date
'If dtStart.Text = vbNullString Then
'    MsgBox "Date is required.  Please enter a valid date.", vbExclamation
'    Hourglass False
'    dtStart.SetFocus
'    Exit Sub
'End If
'
'sCriteria = "SELECT DD_POSTING_HISTORY.INSTITUTION_CODE, DD_POSTING_HISTORY.AFFINITY_ACCT_NUM, DD_POSTING_HISTORY.MEDICAL_RECORD_NUM, DD_POSTING_HISTORY.DD_NUM, DD_POSTING_HISTORY.TOT_FUNB_BENEFIT_AMT, DD_POSTING_HISTORY.DR_CR_FLAG, DD_POSTING_HISTORY.AS_OF_DATETIME, DD_POSTING_HISTORY.PA_DISTRIBUTION_AMT, DD_POSTING_HISTORY.PF_DISTRIBUTION_AMT, DD_POSTING_HISTORY.PATIENT_NAME, DD_POSTING_HISTORY.POSTED_DATETIME, DD_POSTING_HISTORY.TOT_DAYS_INHOUSE, DD_POSTING_HISTORY.SPEC_PROC_COND_HASH_TOT, DD_POSTING_HISTORY.ATP_PML_FLAG,DD_INCOME_SOURCE_TYPE.INCOME_SRC_TYPE_DESCR"
'sCriteria = sCriteria & " FROM DD_POSTING_HISTORY,DD_INCOME_SOURCE_TYPE"
'sCriteria = sCriteria & " WHERE DD_POSTING_HISTORY.INCOME_SOURCE_TYPE_ID = DD_INCOME_SOURCE_TYPE.INCOME_SOURCE_TYPE_ID"
'
'Select Case sdcInstitution.Text
'    Case "All Institutions"
'        sCriteria = sCriteria & " AND DD_POSTING_HISTORY.POSTED_DATETIME >= '" & dtStart.Text & " 00:00:00'"
'        sCriteria = sCriteria & " AND DD_POSTING_HISTORY.POSTED_DATETIME <= '" & dtStart.Text & " 23:59:59'"
'        crReport3.WindowShowGroupTree = True
'    Case Else
'        sCriteria = sCriteria & " AND DD_POSTING_HISTORY.INSTITUTION_CODE = '" & sdcInstitution.Columns(1).Text & "'"
'        sCriteria = sCriteria & " AND DD_POSTING_HISTORY.POSTED_DATETIME >= '" & dtStart.Text & " 00:00:00'"
'        sCriteria = sCriteria & " AND DD_POSTING_HISTORY.POSTED_DATETIME <= '" & dtStart.Text & " 23:59:59'"
'        crReport3.WindowShowGroupTree = False
'End Select
'
'sCriteria = sCriteria & Chr$(13) + Chr$(10) & "ORDER BY DD_POSTING_HISTORY.INSTITUTION_CODE Asc"
'
'
''crReport3.SelectionFormula = sCriteria
'crReport3.SQLQuery = sCriteria
'
''Send the report to be sent to the Screen
'crReport3.Destination = miPrintMode
'
''Maximize the report window
'crReport3.WindowState = crptMaximized
'
'' Show a Window Title for the Report
'crReport3.WindowTitle = "Direct Deposits Detail and Summary Report"
''Passing a parameter "Posted Date"
'crReport3.ParameterFields(0) = "FromDate; " & "Posted on " & dtStart.Text & ";True"
'
''Change the Mouse Pointer to Hourglass cursor
'frmMainReports.MousePointer = vbHourglass
'
''Trigger the event at run time
'crReport3.Action = 1
'
''Reset the Mouse Pointer to the default
'frmMainReports.MousePointer = vbDefault
'Xit:
'    Hourglass False
'
'Exit Sub
'Report3Err:
'    Hourglass False
'    ShowError MODULE + "Direct Deposits Detail and Summary Report", Err
'Resume Xit
'End Sub
' '********************************************************************************
''* Name:  REPORT4()
''* Description: This is a procedure for report # 4 Automated Balancing Worksheet Report
''*               to launch crystal report # 4 with its criteria: Single Date
''* Parameters:
''* Created:
''********************************************************************************
'Private Sub REPORT4()
''Report Name: "Automated Balancing Worksheet Report"
''testing a Crystal Report
''---------------------------------
''
''Error Handler !!
'On Error GoTo Report4Err
'
'Dim crReport4 As Crystal.CrystalReport
'Dim IntRet As Integer
'Dim sCriteria As String
'
'Hourglass True
'
''set to the crystal report
'Set crReport4 = fMainForm.crReport
'crReport4.Reset
'
''** This is commented out so that users can see all Crystal options
''Call ReportButtons
'
''Open a connection to SQL Server
'crReport4.Connect = gobjLoginInfo.ConnectString
'
''call the Reportfile Name called Direct Deposits Detail and Summary Report
'crReport4.ReportFileName = gsDataPath & "\BalWkSht.RPT"
'
''Check that there is a date
'If dtStart.Text = vbNullString Then
'    MsgBox "Date is required.  Please enter a valid date.", vbExclamation
'    Hourglass False
'    dtStart.SetFocus
'    Exit Sub
'End If
'
'sCriteria = "{DD_BALANCE.CREATED_DATETIME} >= Date(Datetime(" & Year(dtStart.Text) & "," & Month(dtStart.Text) & "," & Day(dtStart.Text) & ",00,00,00))"
'sCriteria = sCriteria & " AND {DD_BALANCE.CREATED_DATETIME} <=Date(Datetime(" & Year(dtStart.Text) & "," & Month(dtStart.Text) & "," & Day(dtStart.Text) & ",23,59,59))"
'crReport4.WindowShowGroupTree = False
'
'crReport4.SelectionFormula = sCriteria
'
''Send the report to be sent to the Screen
'crReport4.Destination = miPrintMode
''
''Maximize the report window
'crReport4.WindowState = crptMaximized '
'
'' Show a Window Title for the Report
'crReport4.WindowTitle = "Automated Balancing Worksheet Report"
'
''Passing a parameter As of Date
'crReport4.ParameterFields(0) = "FromDate; " & "AS OF " & dtStart.Text & ";True"
'
''Change the Mouse Pointer to Hourglass cursor
'frmMainReports.MousePointer = vbHourglass
'
''Trigger the event at run time
'crReport4.Action = 1
'
''Reset the Mouse Pointer to the default
'frmMainReports.MousePointer = vbDefault
'Xit:
'    Hourglass False
'
'Exit Sub
'Report4Err:
'Hourglass False
'ShowError MODULE + "Automated Balancing Worksheet Report", Err
'Resume Xit
'End Sub
'
'Private Sub REPORT5()
''Report Name: "Automated Balancing Worksheet Report"
''testing a Crystal Report
''---------------------------------
''
''Error Handler !!
'On Error GoTo Report5Err
'
'Dim IntRet As Integer
'
'Hourglass True
'
''Check that there is a date
'If dtStart.Text = vbNullString Then
'    MsgBox "Date is required.  Please enter a valid date.", vbExclamation
'    Hourglass False
'    dtStart.SetFocus
'    Exit Sub
'End If
'
'If Not IsDate(dtStart.Text) Then
'    MsgBox "Valid date is required.  Please enter a valid date.", vbExclamation
'    Hourglass False
'    dtStart.SetFocus
'    Exit Sub
'End If
'
'VerifyAffinity dtStart.Text, miPrintMode
'
''Reset the Mouse Pointer to the default
'Xit:
'    Hourglass False
'
'Exit Sub
'Report5Err:
'Hourglass False
'ShowError MODULE + "Automated Balancing Worksheet Report", Err
'Resume Xit
'End Sub
 
 '********************************************************************************
'* Name:  ReportButtons()
'* Description: Set all the Options for the Crystal Report button
'* Parameters:
'* Created:7/8/1999
'********************************************************************************
'Private Sub ReportButtons()
'On Error Resume Next
''This is a procedure to turn on and off for the crystal report buttons
' fMainForm.crReport.WindowControlBox = False
' fMainForm.crReport.WindowMinButton = False
' fMainForm.crReport.WindowMaxButton = False
' fMainForm.crReport.WindowShowCancelBtn = True
' fMainForm.crReport.WindowShowCloseBtn = True
' fMainForm.crReport.WindowShowExportBtn = False
' fMainForm.crReport.WindowShowGroupTree = False
' fMainForm.crReport.WindowShowNavigationCtls = True
' fMainForm.crReport.WindowAllowDrillDown = False
' fMainForm.crReport.WindowShowZoomCtl = True
' fMainForm.crReport.WindowShowRefreshBtn = False
' fMainForm.crReport.WindowShowPrintBtn = True
' fMainForm.crReport.WindowShowPrintSetupBtn = True
' fMainForm.crReport.WindowShowProgressCtls = False
'
'End Sub

