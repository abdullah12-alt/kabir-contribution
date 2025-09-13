VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{BB3B26D0-99DF-11D2-9C22-00105A19BCF2}#8.0#0"; "DatePicker.ocx"
Object = "{8CD222DF-7752-11D3-9D1E-00105A19BCF2}#1.0#0"; "OAOTBar.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStateTreasurer 
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11070
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   11070
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdForceSendStTreas 
      Caption         =   "Force Treasurer"
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   4200
      TabIndex        =   22
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdTestEmail 
      Caption         =   "Test Email"
      Height          =   375
      Left            =   2640
      TabIndex        =   21
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrintSetupFile 
      Caption         =   "Print Last Treasurer Report"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton cmdAlogsOnly 
      Caption         =   "Create Alogs Only"
      Height          =   375
      Left            =   5565
      TabIndex        =   17
      Top             =   6240
      Width           =   1440
   End
   Begin OAOTitleBar.OutlookTitleBar OutlookTitle1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   979
      ForeColor       =   16777215
      Caption         =   "State Treasurer"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sdcDSN 
      DataSource      =   "adcDSN"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3030
      Width           =   1215
      DataFieldList   =   "Column(0)"
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
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "DSN#"
      Columns(0).Name =   "DSN#"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Posted Date"
      Columns(1).Name =   "Posted Date"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.Frame Frame7 
      Height          =   5025
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   9135
      Begin VB.Frame fraDSN 
         Caption         =   "Enter the following information"
         Height          =   1980
         Left            =   195
         TabIndex        =   7
         Top             =   2910
         Width           =   3015
         Begin VB.TextBox txtDSN 
            Height          =   330
            Left            =   1185
            MaxLength       =   5
            TabIndex        =   1
            Top             =   855
            Width           =   1740
         End
         Begin DatePicker.DateSelector dtStart 
            Height          =   315
            Left            =   1200
            TabIndex        =   0
            Top             =   360
            Width           =   1770
            _ExtentX        =   3122
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
         Begin DatePicker.DateSelector dtProcessDate 
            Height          =   315
            Left            =   1185
            TabIndex        =   18
            Top             =   1365
            Width           =   1770
            _ExtentX        =   3122
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
         Begin VB.Label lblProcessDate 
            Caption         =   "Process Date"
            Height          =   255
            Left            =   105
            TabIndex        =   19
            Top             =   1395
            Width           =   1065
         End
         Begin VB.Label Label2 
            Caption         =   "Deposit Sequence No"
            Height          =   465
            Left            =   75
            TabIndex        =   16
            Top             =   810
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Date to Send"
            Height          =   255
            Left            =   90
            TabIndex        =   15
            Top             =   390
            Width           =   1065
         End
      End
      Begin VB.Frame fraProcess 
         Height          =   4290
         Left            =   3240
         TabIndex        =   5
         Top             =   600
         Width           =   5715
         Begin VB.Timer Timer1 
            Left            =   4920
            Top             =   1440
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   465
            Left            =   240
            TabIndex        =   8
            Top             =   3480
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   820
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   465
            Left            =   240
            TabIndex        =   9
            Top             =   2400
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   820
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgress 
            Alignment       =   2  'Center
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2250
            TabIndex        =   11
            Top             =   3120
            Width           =   1455
         End
         Begin VB.Label lblProcess 
            Alignment       =   2  'Center
            Caption         =   "Process"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1935
            TabIndex        =   10
            Top             =   2115
            Width           =   2055
         End
         Begin VB.Label lblProcessStatus 
            Caption         =   "Enter the required information and click Process below to start the submission of total deposit information to State Treasurer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1380
            Left            =   2070
            TabIndex        =   6
            Top             =   390
            Width           =   3450
         End
         Begin VB.Image imgDSN 
            Height          =   1680
            Left            =   105
            Stretch         =   -1  'True
            Top             =   435
            Width           =   1815
         End
      End
      Begin VB.Label lblViewPriorDSNs 
         Caption         =   "View Prior DSN's"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   2475
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   2145
         Left            =   360
         Picture         =   "frmStateTreasurer.frx":0000
         Top             =   240
         Width           =   2250
      End
   End
   Begin MSComDlg.CommonDialog dlgPrintBatchFiles 
      Left            =   8040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8385
      TabIndex        =   3
      Top             =   6225
      Width           =   1215
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   90
      X2              =   9735
      Y1              =   5745
      Y2              =   5760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      BorderStyle     =   6  'Inside Solid
      X1              =   105
      X2              =   9705
      Y1              =   5775
      Y2              =   5775
   End
End
Attribute VB_Name = "frmStateTreasurer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************
' * Form Name:frmStateTreasurer
' * Form File Name: frmStateTreasurer.frm
' * Start Date: 6/1/1999
' * End Date:   7/26/1999
' * Description:
' * --------------------------------
' * The State Treasurer screen is designed using one control button
'   called "Process" which includes the following four component process:

'� Entering the Deposit Sequence Number, If Needed
'� Creating Deposit Totals Files, Four Files setup.dbf "FoxPro Format"
'� Sending E-mail Files
'� Printing Batch Files
'
'
' Mod CONSTANTS
'
Private Const MODULE As String = "State Treasurer"
'
' Mod VARIABLES
'

Dim mbDSNRequired As Boolean
Dim mbIncompleteSend As Boolean

Dim intFlag As Integer ' used for the timer
Dim msPFBatchName As String
Dim msAffBatchName As String




''********************************************************************************
''* Name: chkPriorDayData_Click
''* Description: If the user wants to process the prior Day's data
''* Created: 6/11/99 4:06 PM
''********************************************************************************
'
'Private Sub chkPriorDayData_Click()
'On Error GoTo ChkPriorDayDataErr
'
'If chkPriorDayData.Value = vbChecked Then
'
'    fraPriorDaysData.Visible = True
'    'remove the Frame DSN
'    fraDSN.Visible = False
'    fraPriorDaysData.Top = 3720
'    fraPriorDaysData.Left = 120
'    lblProcessStatus.Caption = "Enter the Prior Day's Data, then Press Process Button"
'Else
'    'remove the frame of the prior data option
'    fraPriorDaysData.Visible = False
'    'bring back the Frame DSN
'    fraDSN.Visible = True
'    If mbDSNEnableFlag = True Then
'        cmdProcess.Enabled = True
'        lblProcessStatus.Caption = "Click on the Process button"
'    Else
'        cmdProcess.Enabled = False
'        lblProcessStatus.Caption = "Enter the Deposit Sequence Number, then Press Tab Key "
'   End If
'End If
'Xit:
'    Exit Sub
'
'ChkPriorDayDataErr:
'    ShowUnexpectedError MODULE + " chkPriorDayData ", Err
'    Resume Xit
'
'
'End Sub
Private Sub SentToStTreas()

On Error GoTo SentToStTreasErr

'********************************************************************************
'* Name: SentToStTreas
'* Description:Calling the Procedure to Update the SentToStTreasIndicator TO 'Y' after Email is successed
'* the stored procedure callled "up_i_Sent_To_St_Treas"
'* Created: 7/21/99 3:39
'********************************************************************************
'
'This will alocate the Correct Posting ID based on the Posting Data to Update the up_i_Sent_To_St_Treas Procedure
'
'
Dim cmd As New ADODB.Command
Dim TempDate As String
    TempDate = dtStart.Text
    
    If mbDSNRequired = True Then
        Call iuInsertDSN
    End If
        
    Set cmd.ActiveConnection = gcnDDS
    If gStoredProcs("up_u_Sent_To_St_Treas").GetStoredProcCommand(cmd) = True Then
        cmd.Parameters("SENT_TO_ST_TREAS_IND") = "Y"
        If mbDSNRequired = True Then
            cmd.Parameters("DEP_SEQ_NUM") = txtDSN.Text
        End If
        cmd.Parameters("POSTED_DATETIME") = Format$(TempDate, "MM/dd/yyyy")
        cmd.Execute
        If cmd.Parameters("RETURN_VALUE") <> 0 Then
            GetServerErrorMsg cmd.Parameters("RETURN_VALUE"), "Error occurred adding or updating the State record."
        End If
    Else
        MsgBox "Error Updating the Sent To State Treas Stored Procedure.", vbCritical
        Set cmd = Nothing
        ExitApp
    End If
    Set cmd = Nothing

Xit:
    Exit Sub

SentToStTreasErr:

    ShowUnexpectedError MODULE + " Sent_To_St_Treas ", Err
    Resume Xit

End Sub

Private Sub cmdAlogsOnly_Click()

On Error GoTo cmdAlogsOnlyErr

Dim bErrorOccurred As Boolean

dtStart.Enabled = False
txtDSN.Enabled = False

'Check to see if this is a valid date
If Not IsDate(dtStart.Text) Then
    MsgBox "The date entered is invalid", vbInformation
    GoTo Xit
End If
    
If mbDSNRequired = True Then
    If fnCkCodeExists = True Then
        'user entered an existing DSN
        GoTo Xit
    End If
End If

'Hide the DSN Frame
'Hide the Prior Days Data
lblProgress.Visible = True
lblProcess.Visible = True
ProgressBar1.Visible = True
ProgressBar2.Visible = True
'Issues for the process Call InsertingDSNValue:
Hourglass True

SetPercentage 1
Timer1.Interval = 1
'testing
cmdProcess.Caption = "&Creating"
lblProgress.Caption = "Creating Files"
cmdProcess.Enabled = False
imgDSN.Picture = LoadPicture()


Set imgDSN.Picture = fMainForm.ImageList1.ListImages("Signboo").Picture
lblProcessStatus = "Creating Deposits Totals process is in progress... " & vbCrLf & "Please wait until the creating process is complete..."
SetPercentage 3

'call creating the SetupFiles.....
If CreateSetupFiles = False Then
    'Files were not created successfully
    Timer1.Enabled = False
    GoTo Xit
End If

If CreateALOGS = False Then
    Timer1.Enabled = False
    GoTo Xit
End If

Timer1.Enabled = True

 lblProgress.Caption = "Email Files"
 cmdProcess.Enabled = False
 imgDSN.Picture = LoadPicture()
 Set imgDSN.Picture = fMainForm.ImageList1.ListImages("Email").Picture
 lblProcessStatus = "Email Files process is in progress... " & vbCrLf & "Please wait until the sending files process is complete..."
 cmdProcess.Enabled = False
 'lblProcessStatus = "Creating Depoist Totals."
 'If you want to Print the Batch Files now click the Print button below to start the Printing process."
 cmdProcess.Caption = "&Email"
 SetPercentage 36

If SendEmailForAlogs("The attached alogs dated " & dtStart.Text & " were not received by all parties and was requested again.  Please verify whether you received this file previously and have already entered the data.  If you have, please disregard this email.") = False Then
    MsgBox "A problem occurred sending email for ALOGS. Send the files manually."
End If
    
 'Display message for the User if Error Occurs
If bErrorOccurred = False Then
      MsgBox "The ALogs Have been created and sent." & vbNewLine & _
     "No errors found.", vbInformation
     lblProcessStatus.Caption = " "
     imgDSN.Picture = LoadPicture()
     Set imgDSN.Picture = fMainForm.ImageList1.ListImages("Happy").Picture
     cmdProcess.Caption = "&Done"
     cmdProcess.Enabled = False
     cmdCancel.Enabled = True
     intFlag = 0
     ProgressBar1.value = 0
     ProgressBar1.Visible = False
     ProgressBar2.value = 0
     ProgressBar2.Visible = False
     lblProcess.Visible = True
     lblProcess.Caption = "Process is Done"
     lblProgress.Visible = False
     Err.Clear
     Timer1.Enabled = False
     Hourglass False
Else
    MsgBox "The submission of total deposit information to the State Treasurer is NOT COMPLETED." & vbNewLine & _
    "Errors found. " & Err.Description, vbInformation
    lblProcessStatus.Caption = " "
    imgDSN.Picture = LoadPicture()
    Set imgDSN.Picture = fMainForm.ImageList1.ListImages("Help02").Picture
    cmdProcess.Caption = "&Exit"
    cmdProcess.Enabled = False
    cmdCancel.Enabled = True
    intFlag = 0
    ProgressBar1.value = 0
    ProgressBar1.Visible = False
    ProgressBar2.value = 0
    ProgressBar2.Visible = False
    lblProcess.Visible = True
    lblProcess.Caption = "Process had problems"
    lblProgress.Visible = False
    Err.Clear
    Timer1.Enabled = False
    Hourglass False
End If

'
Xit:
    Timer1.Enabled = False
    Hourglass False
    Exit Sub

cmdAlogsOnlyErr:

    ShowUnexpectedError MODULE + " cmdAlogsOnly ", Err
    Resume Xit

End Sub

Private Sub cmdForceSendStTreas_Click()
    If MsgBox("Make sure proper date in Date to Send id filled correctly and also the dep seq no" & vbCrLf & "Are you sure you want to do this?", vbYesNo) = vbYes Then
        SentToStTreas
        MsgBox dtStart.Text & " has been marked complete for State Treasurer"
    End If
End Sub

Private Sub cmdPrintSetupFile_Click()
    PrintingSetupFiles
    
End Sub

'********************************************************************************
'* Name: cmdProcess_Click()
'* Description:It starts the Process:
' 1- Entering the Deposit Sequence Number, If Needed
' 2- Creating Deposit Totals Files, Four Files setup.dbf "FoxPro Format"
' 3- Sending E-mail Files
' 4- Printing Batch Files
'*
'* Created: 7/21/99 3:39
'********************************************************************************
Private Sub cmdProcess_Click()
On Error GoTo ProcessErr

Dim bErrorOccurred As Boolean
Dim iRet As Integer
dtStart.Enabled = False
txtDSN.Enabled = False
dtProcessDate.Enabled = False
'Check to see if this is a valid date
If Not IsDate(dtStart.Text) Then
    MsgBox "The date entered is invalid", vbInformation
    GoTo Xit
End If
    
'Check to see if this is a valid date
If Not IsDate(dtProcessDate.Text) Then
    MsgBox "The process date entered is invalid", vbInformation
    GoTo Xit
End If
    
If mbDSNRequired = True Then
    If fnCkCodeExists = True Then
        'user entered an existing DSN
        GoTo Xit
    End If
End If

'Hide the DSN Frame
'Hide the Prior Days Data
lblProgress.Visible = True
lblProcess.Visible = True
ProgressBar1.Visible = True
ProgressBar2.Visible = True
'Issues for the process Call InsertingDSNValue:
'1-Check the DSN is Unique Number
'2- Write DSN data to the correct Fields and to the DataBASE
'3- Check the DSN that was not send before
'4- Refresh the DSN Combo box

'Here it is :Insert DSN new value using Stored Procedures to the DSN TABLE...'
Hourglass True

SetPercentage 1
Timer1.Interval = 1
'testing
cmdProcess.Caption = "&Creating"
lblProgress.Caption = "Creating Files"
cmdProcess.Enabled = False
imgDSN.Picture = LoadPicture()


Set imgDSN.Picture = fMainForm.ImageList1.ListImages("Signboo").Picture
lblProcessStatus = "Creating Deposits Totals process is in progress... " & vbCrLf & "Please wait until the creating process is complete..."
SetPercentage 3

'call creating the SetupFiles.....
iRet = CreateSetupFiles
If iRet = 0 Then
    'Files were not created successfully
    Timer1.Enabled = False
    GoTo Xit
End If

If CreateALOGS = False Then
    Timer1.Enabled = False
    GoTo Xit
End If

Timer1.Enabled = True

 lblProgress.Caption = "Email Files"
 cmdProcess.Enabled = False
 imgDSN.Picture = LoadPicture()
 Set imgDSN.Picture = fMainForm.ImageList1.ListImages("Email").Picture
 lblProcessStatus = "Email Files process is in progress... " & vbCrLf & "Please wait until the sending files process is complete..."
 cmdProcess.Enabled = False
 'lblProcessStatus = "Creating Depoist Totals."
 'If you want to Print the Batch Files now click the Print button below to start the Printing process."
 cmdProcess.Caption = "&Email"
 SetPercentage 36

'send the file to sips
If SendFiles(iRet) = False Then
    MsgBox "You had a problem sending the files to the State Treasurer. " & vbNewLine & "Your files did NOT get to your intended target"
    bErrorOccurred = True
Else
    'Start Sending EMAIL TO THE STATE TREASURER WITH ATTACHMENT FILES (SetUp Files)
    If SendEmail = False Then
        MsgBox "You had a problem sending email to Team91. " & vbNewLine & "Your files were sent to intended target"
        bErrorOccurred = False
    End If
    If SendEmailForAlogs = False Then
        MsgBox "A problem occurred sending email for ALOGS. Send the files manually."
    End If
    
    SetPercentage 60
    Call SentToStTreas
    SetPercentage 65

End If

imgDSN.Picture = LoadPicture()
Set imgDSN.Picture = fMainForm.ImageList2.ListImages("Printer").Picture
lblProcessStatus = "Printing Files is in progress... " & vbCrLf & "Please wait until the printing process is complete..."

lblProgress.Caption = "Printing Files"
cmdProcess.Caption = "&Printing"
SetPercentage 66

'Calling a procedure for Printing Files
If PrintingSetupFiles = False Then
    bErrorOccurred = True
End If
SetPercentage 100

 'Display message for the User if Error Occurs
If bErrorOccurred = False Then
      MsgBox "The submission of total deposit information to the State Treasurer is complete." & vbNewLine & _
     "No errors found.", vbInformation
     lblProcessStatus.Caption = " "
     imgDSN.Picture = LoadPicture()
     Set imgDSN.Picture = fMainForm.ImageList1.ListImages("Happy").Picture
     cmdProcess.Caption = "&Done"
     cmdProcess.Enabled = False
     cmdCancel.Enabled = True
     intFlag = 0
     ProgressBar1.value = 0
     ProgressBar1.Visible = False
     ProgressBar2.value = 0
     ProgressBar2.Visible = False
     lblProcess.Visible = True
     lblProcess.Caption = "Process is Done"
     lblProgress.Visible = False
     Err.Clear
     Timer1.Enabled = False
     Hourglass False
Else
    MsgBox "The submission of total deposit information to the State Treasurer is NOT COMPLETED." & vbNewLine & _
    "Errors found. " & Err.Description, vbInformation
    lblProcessStatus.Caption = " "
    imgDSN.Picture = LoadPicture()
    Set imgDSN.Picture = fMainForm.ImageList1.ListImages("Help02").Picture
    cmdProcess.Caption = "&Exit"
    cmdProcess.Enabled = False
    cmdCancel.Enabled = True
    intFlag = 0
    ProgressBar1.value = 0
    ProgressBar1.Visible = False
    ProgressBar2.value = 0
    ProgressBar2.Visible = False
    lblProcess.Visible = True
    lblProcess.Caption = "Process had problems"
    lblProgress.Visible = False
    Err.Clear
    Timer1.Enabled = False
    Hourglass False
End If

'
Xit:
    Timer1.Enabled = False
    Hourglass False
    Exit Sub

ProcessErr:

    ShowUnexpectedError MODULE + " Process ", Err
    Resume Xit

End Sub
Private Function SFTPFile(ByVal sLocalName As String, ByVal sRemoteFileName) As Boolean

Dim sLine As String
Dim sRet As String
Dim sText As String
Dim ix As Integer

sLine = App.Path & "\pscp -batch -pw p@21rspX" & " """ & gsDataPath & "\" & sLocalName & """ " & "ddssftp@hes001.dhr.state.nc.us:" & "/treas/" & sRemoteFileName
'Try 3 time to send the file
For ix = 1 To 3
    fMainForm.msDosOutput = ""
    fMainForm.objDOS.CommandLine = sLine
    sRet = fMainForm.objDOS.ExecuteCommand
    DoEvents
'    If sRet <> "" And InStr(1, fMainForm.msDosOutput, "100%") > 0 Then
    If sRet <> "" And InStr(1, sRet, "100%") > 0 Then
        'MsgBox "File has been securely copied."
        SFTPFile = True
        Exit For
    Else
       If InStr(1, fMainForm.msDosOutput, "fingerprint") > 0 Then
            SaveKeyToRegistry
       End If
    End If


Next ix

Xit:

Exit Function

SFTPFileErr:
MsgBox Error
Resume Xit

End Function

Private Function SaveKeyToRegistry() As Boolean
On Error GoTo SaveKeyToRegistryErr
    Dim objReg As CRegAPI
    Set objReg = New CRegAPI
    
    objReg.Root = Hkey_Current_User
    objReg.KeyPrefix = "SOFTWARE\SimonTatham"
    Call objReg.SaveSetting("PuTTY", "SshHostKeys", "rsa2@22:199.90.16.15", "0x23,0xcd451ed5b730b01634c013a282efe06e730d9917ca1ba104757835901484b4a3453fdf4bfa254f61be01e2dd8ac9ff69c7837fed695294f7fd580d97c7e84df4e68b50e2635749c2903394556465be051ca4be03f7972a7089d2007f3254dba29475d14a0badffabd413dd30af1b310370fccf6a412694f3f730733d621ac691")
    SaveKeyToRegistry = True
Xit:
Set objReg = Nothing
Exit Function

SaveKeyToRegistryErr:
'Skip error
Resume Xit

End Function


Public Function SendFiles(iCreateProcess) As Boolean

'Dim oFtp As New FTPClass
Dim strFileName As String
Dim bRet As Boolean
On Error GoTo ErrRtn
    SendFiles = False

    
    If iCreateProcess = 1 Or iCreateProcess = 3 Then
        strFileName = "OSTMHLT.txt"

        'AS - 8/28/2014 - Replaced code to ftp file with secure ftp
        
        'bRet = FTPFile(strFileName, strFileName)
        bRet = SFTPFile(strFileName, strFileName)
        
        If bRet = False Then
            Err.Raise 10002, , "Error sending the OSTMHLT.txt"
        End If
    End If
    If iCreateProcess = 2 Or iCreateProcess = 3 Then
        strFileName = "OSTMHPA.txt"
        'bRet = FTPFile(strFileName, strFileName)
        bRet = SFTPFile(strFileName, strFileName)
        
        If bRet = False Then
            Err.Raise 10002, , "Error sending the OSTMHPA.txt"
        End If
    End If
    
    SendFiles = True

Xit:
    Exit Function

ErrRtn:
    ShowError MODULE + ".SendFiles", Err
    Resume Xit


End Function

'Private Function FTPFile(ByVal sLocalFileName As String, sDestFileName) As Boolean
''*******************************************************
''* Function: FTPFile
''* Created: 3/2/2009
''* Used to send file to ddapp directory
''*
''*******************************************************
'On Error GoTo FTPFileErr
'
'If mobjFTP.Connect <> 0 Then
'    Err.Raise 21345, "State Treasurer", "Could not connect to server"
'End If
'
'mobjFTP.RemoteDirectory = msFTPRoot & "/data"
'mobjFTP.LocalPath = gsDataPath
'mobjFTP.LocalFile = sLocalFileName
'mobjFTP.RemoteFile = sDestFileName
'mobjFTP.BinaryTransfer = False
'If mobjFTP.ChangeDirectory(mobjFTP.RemoteDirectory) <> 0 Then
'    Err.Raise COULD_NOT_CREATE_DIRECTORY, "The directory " & mobjFTP.RemoteDirectory & " could not be located."
'End If
'If mobjFTP.PutFile(True) <> 0 Then
'    Err.Raise COULD_NOT_PUT_ON_SERVER, "The file '" & mobjFTP.LocalPath & "\" & mobjFTP.LocalFile & "' could not be copied to report server"
'End If
'
'FTPFile = True
'Xit:
'mobjFTP.Disconnect
'
'Exit Function
'
'FTPFileErr:
'
'    MsgBox Error, vbInformation
'    GoTo Xit
'End Function

'********************************************************************************
'* Name:  SendEmail()
'* Description:Calling the Procedure to Update the SentToStTreas TO 'Y' after Email is successed
'* the stored procedure callled "up_u_Sent_To_St_Treas"
'* Created: 5/26/1999 3:39
'********************************************************************************

Private Function SendEmail() As Boolean
On Error GoTo emailerror

Dim rsConfig As New ADODB.Recordset
Dim sRecipEmail As String
Dim sRecipEmailCC As String

Dim iStart As Integer
Dim iEnd As Integer
Dim sPartAddr As String

Dim sSubject As String
Dim sText As String
    
Dim sDate As String
Dim oEmailMsg As New clsEmailMessage
    
'Dim sRecipName As String
rsConfig.Open "SELECT * FROM DD_CONFIG_INFO", gcnDDS, adOpenForwardOnly
If rsConfig.EOF Then
    MsgBox "Error reading configuration record.", vbCritical
    Exit Function
End If

'assign the email Options
oEmailMsg.SMTPServer = "outbound.mail.nc.gov"
oEmailMsg.From = ReadIniFile(App.Path & "\dds.ini", "Startup", "SMTPuser")
oEmailMsg.ToRecipient = ConvertNull(rsConfig!ST_TREAS_EMAIL_TO_ADDR)
oEmailMsg.CCRecipient = ConvertNull(rsConfig!ST_TREAS_EMAIL_CC_ADDR)
oEmailMsg.Message = ConvertNull(rsConfig!ST_TREAS_EMAIL_TEXT)

'based on the selection if the user decide to Post Prior Dates
sDate = dtStart.Text

'Pick the Correct Subject with or without DSN # BASED on the Value of Sum of Pa Distribution Amt
 If mbDSNRequired = False Then
    oEmailMsg.Subject = ConvertNull(rsConfig!ST_TREAS_EMAIL_SUBJ) & " for " & sDate
Else
' add the DSN & Date to your Subject Text
    oEmailMsg.Subject = ConvertNull(rsConfig!ST_TREAS_EMAIL_SUBJ) & " for " & sDate & "  DSN # " & txtDSN.Text & " on " & Now()
End If

SetPercentage 55
Set rsConfig = Nothing

'This send the email out
fMainForm.SendEmail oEmailMsg

Set oEmailMsg = Nothing

SendEmail = True

Exit Function
emailerror:
SendEmail = False
MsgBox "Error occurs  " & Err.Number & vbNewLine & Err.Description, vbCritical

End Function


'********************************************************************************
'* Name:  SendEmail()
'* Description:Calling the Procedure to Update the SentToStTreas TO 'Y' after Email is successed
'* the stored procedure callled "up_u_Sent_To_St_Treas"
'* Created: 5/26/1999 3:39
'********************************************************************************

Private Function SendEmailForAlogs(Optional ByVal sText As String) As Boolean
On Error GoTo emailALogErr
Dim rsRegion As New ADODB.Recordset
Dim sDate As String
Dim oEmailMsg As New clsEmailMessage
'Open up a registry object
    
'Dim sRecipName As String
rsRegion.Open "SELECT * FROM DD_REGION", gcnDDS, adOpenForwardOnly
If rsRegion.EOF Then
    MsgBox "Error reading configuration record.", vbCritical
    Exit Function
End If

Do Until rsRegion.EOF
    If Not IsNull(rsRegion!EMAIL_RECIPIENTS_TO) Then
        If FileExists(gsDataPath & "\ddalog-" & rsRegion!REGION & ".xls") Then
            'assign the email Options
            oEmailMsg.SMTPServer = "outbound.mail.nc.gov"
            oEmailMsg.From = ReadIniFile(App.Path & "\dds.ini", "Startup", "SMTPuser")
            oEmailMsg.ToRecipient = ConvertNull(rsRegion!EMAIL_RECIPIENTS_TO)
            oEmailMsg.CCRecipient = ConvertNull(rsRegion!EMAIL_RECIPIENTS_CC)
            If sText = vbNullString Then
                oEmailMsg.Message = "Enclosed you will find the attached ALOGS for " & dtStart.Text
            End If
            oEmailMsg.Subject = "ALOGS for " & dtStart.Text

            SetPercentage 45

           'define your attatchments
            'Attachment # 1
            If oEmailMsg.AddEmailAttachment(gsDataPath & "\ddalog-" & rsRegion!REGION & ".xls") = False Then
                MsgBox "File to attach is not found"
            End If
                       
            fMainForm.SendEmail oEmailMsg
            Set oEmailMsg = Nothing
        End If
    End If
    rsRegion.MoveNext
Loop
SetPercentage 55

Xit:

Set rsRegion = Nothing
SendEmailForAlogs = True

Exit Function
emailALogErr:
SendEmailForAlogs = False
MsgBox "Error occurs  " & Err.Number & vbNewLine & Err.Description, vbCritical
Resume Xit
Resume
End Function



'******************************************************************************
'* Name:  RefreshCombo()
'* Description:Fill list box with DSN Numbers
'*
'* Created: 6/26/1999 11:40
'********************************************************************************
Private Sub RefreshCombo()

On Error GoTo RefreshDSNComboErr
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim sSql As String

Set cmd.ActiveConnection = gcnDDS
If gStoredProcs("up_s_SeqNo").GetStoredProcCommand(cmd) = False Then
    Err.Raise 2345, , "Seqno Stored Procedure failed"
End If

'Fill list box with DSN Numbers
'Note you need to connect the Posting Date from the History Table......
'

'sSql = "SELECT DD_DEP_SEQ_NO.DEP_SEQ_NUM AS 'DSN#', DD_DEP_SEQ_NO.CREATED_DATETIME AS 'Posted Date'"
'sSql = sSql & " From DD_DEP_SEQ_NO ORDER BY DD_DEP_SEQ_NO.CREATED_DATETIME DESC,DEP_SEQ_NUM"

Set rs = cmd.Execute
'rs.CursorLocation = adUseServer
'rs.Open sSql, gcnDDS, adOpenForwardOnly, adLockReadOnly

'Clear all the items first
sdcDSN.removeAll

With rs
    Do Until .EOF
       sdcDSN.AddItem .Fields("DSN#") & vbTab & .Fields("Posted Date")
      .MoveNext
    Loop
End With
'Close the recordset
rs.Close
Xit:
    Set rs = Nothing
    Set cmd = Nothing
    
    Hourglass False
    Exit Sub

RefreshDSNComboErr:
    ShowUnexpectedError MODULE + " Refresh the DSN Values ", Err
    Resume Xit
End Sub


Private Function CheckforIncompleteSend() As Boolean

'Evaluate if the Patient Accounts are updated then DSN is required, otherwise
'DSN is NOT Require.....
On Error GoTo CheckforIncompleteSendErr

Dim rsDSN As New ADODB.Recordset
Dim sSql As String
Dim TempDate As String
mbDSNRequired = False

TempDate = dtStart.Text

'Check to see if this date is an incomplete send
sSql = "SELECT POSTING_ID"
sSql = sSql & " FROM DD_POSTING_HISTORY"
sSql = sSql & " WHERE SENT_TO_ST_TREAS_IND = 'N' AND POSTED_DATETIME >= '" & TempDate & " 00:00:00" & "' And POSTED_DATETIME <= '" & TempDate & " 23:59:59" & "' "
'Open the recordset with sSql Statement
rsDSN.Open sSql, gcnDDS
If rsDSN.EOF Then
'   No records founds...."
    CheckforIncompleteSend = False
Else
    CheckforIncompleteSend = True
End If

rsDSN.Close

'Check to see to if the DSN is required
sSql = "SELECT POSTING_ID"
sSql = sSql & " FROM DD_POSTING_HISTORY"
sSql = sSql & " WHERE PA_DISTRIBUTION_AMT > 0 AND POSTED_DATETIME >= '" & TempDate & " 00:00:00" & "' And POSTED_DATETIME <= '" & TempDate & " 23:59:59" & "' "
'Open the recordset with sSql Statement
rsDSN.Open sSql, gcnDDS, adOpenForwardOnly, adLockReadOnly
If rsDSN.EOF Then
'   No records founds...."
    mbDSNRequired = False
    txtDSN.BackColor = vbWhite
    txtDSN.Enabled = False
    cmdProcess.Enabled = True
Else
    mbDSNRequired = True
    txtDSN.Enabled = True
    txtDSN.BackColor = PALE_YELLOW
End If

rsDSN.Close
Set rsDSN = Nothing

Exit Function

CheckforIncompleteSendErr:
    
    ShowUnexpectedError "Check for Incomplete Send", Err

End Function


Private Function CheckforTransactions() As Boolean

'Evaluate if the Patient Accounts are updated then DSN is required, otherwise
'DSN is NOT Require.....
On Error GoTo CheckforTransactionsErr

Dim rsDSN As New ADODB.Recordset
Dim sSql As String

Dim TempDate As String
TempDate = Format$(CDate(dtStart.Text), "MM/dd/yyyy")
sSql = "SELECT POSTING_ID"
sSql = sSql & " FROM DD_POSTING_HISTORY"
sSql = sSql & " WHERE DD_POSTING_HISTORY.POSTED_DATETIME >= '" & TempDate & " 00:00:00" & "' And DD_POSTING_HISTORY.POSTED_DATETIME <= '" & TempDate & " 23:59:59" & "' "

'Open the recordset with sSql Statement
rsDSN.Open sSql, gcnDDS
If Not rsDSN.EOF Then
    CheckforTransactions = True
End If

'Close the recordset
rsDSN.Close

Exit Function

CheckforTransactionsErr:
    
    ShowUnexpectedError "Check for Transactions", Err

End Function


Private Function fnCkCodeExists() As Boolean

On Error GoTo fnCkCodeExistsErr

'********************************************************************************
'* Name: fnCkCodeExists
'* Description: It checks if the DSN Code exists or Not
'* Created: 7/21/99 12:40
'********************************************************************************
Dim rsDSN As ADODB.Recordset
Dim sSql As String

If txtDSN = vbNullString Then
    fnCkCodeExists = False
    Exit Function
End If

If mbDSNRequired = False Then
    'MsgBox "There are no need to add DSN"
    Exit Function
End If

Set rsDSN = New ADODB.Recordset
sSql = "SELECT DEP_SEQ_NUM FROM DD_DEP_SEQ_NO"
sSql = sSql & " WHERE DEP_SEQ_NUM = " & txtDSN
sSql = sSql & " AND CREATED_DATETIME > '" & DateAdd("m", -6, Now) & "'"
rsDSN.Open sSql, gcnDDS
    
'Check for condition of no DSN's
If rsDSN.EOF Then
    fnCkCodeExists = False
    Exit Function
Else
    Hourglass False
    MsgBox "This DSN code exists, please enter a different value", vbInformation
    'set focus to code entry field
    fnCkCodeExists = True
    txtDSN.Enabled = True
    txtDSN.SetFocus
End If

Xit:
    Exit Function

fnCkCodeExistsErr:
    ShowUnexpectedError MODULE + "fnCkCodeExists", Err
    Resume Xit


End Function

Private Sub iuInsertDSN()

On Error GoTo iuInsertDSNErr

'********************************************************************************
'* Name: iuInsertDSN
'* Description:Calling the Procedure to Update the DSN Options
'* the stored procedure callled "up_iu_DSN"
'* Created: 7/20/99 3:39
'********************************************************************************

    Dim cmd As New ADODB.Command
    '-------------
    Set cmd.ActiveConnection = gcnDDS
    If gStoredProcs("up_i_DSN").GetStoredProcCommand(cmd) = True Then
        cmd.Parameters("CREATED_BY") = gobjLoginInfo.UserId
        cmd.Parameters("DEP_SEQ_NUM") = Val(txtDSN.Text)
        cmd.Parameters("PROCESS_DATE") = Format$(dtProcessDate.Text, "MM/dd/yyyy")
        cmd.Execute
        If cmd.Parameters("RETURN_VALUE") <> 0 Then
            GetServerErrorMsg cmd.Parameters("RETURN_VALUE"), "Error occurred adding or updating the State record."
        End If
    Else
        MsgBox "Error creating the Insert the DSN Stored Procedure.", vbCritical
        Set cmd = Nothing
        ExitApp
    End If
    Set cmd = Nothing

Xit:
    Exit Sub

iuInsertDSNErr:
    ShowError MODULE + "iuInsertDSN", Err
    Resume Xit

End Sub





Private Sub DateSelector1_KeyPressError(KeyAscii As Integer, newKeyAscii As Integer)

End Sub

Private Sub CmdTestSecureSend_Click()
Dim bRet As Boolean
'bRet = FTPFile("OSTMHLT.txt", "OSTMHLT.txt")
'bRet = fMainForm.SecureSend("C:\test\op.txt", "op.txt")
    MsgBox bRet
    'frmAbout.Show vbModal, Me
End Sub

Private Sub cmdSendEmail_Click()

End Sub

Private Sub cmdTestEmail_Click()
On Error GoTo TestEmailErr
Dim rsConfig As New ADODB.Recordset
Dim sRecipEmail As String
Dim sRecipEmailCC As String

Dim iStart As Integer
Dim iEnd As Integer
Dim sPartAddr As String

Dim sSubject As String
Dim sText As String
    
Dim sDate As String
Dim sEmailAddr As String
Dim oEmailMsg As New clsEmailMessage
    
'Dim sRecipName As String
rsConfig.Open "SELECT * FROM DD_CONFIG_INFO", gcnDDS, adOpenForwardOnly
If rsConfig.EOF Then
    Err.Raise 13011, "Test Email", "Error reading configuration record."
End If

'assign the email Options
oEmailMsg.SMTPServer = "outbound.mail.nc.gov"
oEmailMsg.From = ReadIniFile(App.Path & "\dds.ini", "Startup", "SMTPuser")
sEmailAddr = InputBox("Enter email address or leave blank for default recipients")
If sEmailAddr <> "" Then
    oEmailMsg.ToRecipient = sEmailAddr
    oEmailMsg.CCRecipient = sEmailAddr
Else
    oEmailMsg.ToRecipient = ConvertNull(rsConfig!ST_TREAS_EMAIL_TO_ADDR)
    oEmailMsg.CCRecipient = ConvertNull(rsConfig!ST_TREAS_EMAIL_CC_ADDR)
End If

oEmailMsg.Message = "Testing Email. PLease ignore"

'based on the selection if the user decide to Post Prior Dates
sDate = dtStart.Text

' add the DSN & Date to your Subject Text
    oEmailMsg.Subject = "Testing " & ConvertNull(rsConfig!ST_TREAS_EMAIL_SUBJ) & " for " & sDate & "  DSN # " & txtDSN.Text & " on " & Now()


SetPercentage 55
Set rsConfig = Nothing

'This send the email out
fMainForm.SendEmail oEmailMsg

Set oEmailMsg = Nothing

Exit Sub
TestEmailErr:
MsgBox Error


End Sub

Private Sub Command1_Click()
If CreateALOGS = False Then
End If

End Sub

Private Sub Command2_Click()

End Sub

Private Sub dtStart_Change()

cmdProcess.Enabled = False
If Len(dtStart.Text) > 7 And IsDate(dtStart.Text) Then
    If Year(dtStart.Text) > Year(DateAdd("yyyy", -5, Now)) Then
        If CheckforTransactions = False Then
            MsgBox "No transactions need to be sent to the State Treasurer for this date", vbInformation
        Else
            If CheckforIncompleteSend = False Then
                mbIncompleteSend = False
                mbDSNRequired = False
                txtDSN.BackColor = vbWhite
                txtDSN.Enabled = False
                cmdProcess.Enabled = True
                lblProcessStatus.Caption = "The following day will be resent to the State Treasurer.  Press the Process button to continue."
            Else
                mbIncompleteSend = True
                If mbDSNRequired = True Then
                    txtDSN.BackColor = PALE_YELLOW
                    txtDSN.Enabled = True
                    txtDSN.SetFocus
                    lblProcessStatus.Caption = "Enter the Deposit Sequence Number and then press the Process button"
                Else
                    txtDSN.BackColor = vbWhite
                    txtDSN.Enabled = False
                    lblProcessStatus.Caption = "Press the Process button to continue"
                End If
            End If
        End If
    Else
        MsgBox "No transactions need to be sent to the State Treasurer for this date", vbInformation
    End If
End If

End Sub



Private Sub dtStart_GotFocus()
                
SetSelected
                

End Sub

Private Sub Form_Activate()
    fMainForm.SetMainToolbar True
    RefreshCombo
End Sub

Private Sub Form_Deactivate()
fMainForm.SetMainToolbar True

End Sub



Private Sub Label3_Click()

End Sub



Private Sub Form_Unload(Cancel As Integer)

'Set mobjFTP = Nothing

End Sub

Private Sub OutlookTitle1_IconClick()
    
    If cmdCancel.Enabled = True Then
        Unload Me
    End If
    
End Sub



Private Sub sdcDSN_CloseUp()

dtStart.Text = Format$(sdcDSN.Columns(1).Text, "MM/dd/yyyy")

End Sub
' Is this valid function for checking AlphaNumeric field? - EP 03/31/2021
' Mid Function - Returns or replaces a specified number of characters from a string.
' Syntax: Mid(string,start[,length])[=string1]
Public Function IsAlphaNumeric(strInputText As String) As Boolean
Dim intCounter As Integer
Dim strCompare As String
Dim strInput As String
IsAlphaNumeric = False
'Iterate through each character and determine if its a number or letter.
For intCounter = 1 To Len(strInputText)
    strCompare = Mid$(strInputText, intCounter, 1)
    strInput = Mid$(strInputText, intCounter + 1, Len(strInputText))
    If strCompare Like ("[A-Z]") Or strCompare Like ("[a-z]") Or strCompare Like ("0-9") Then
        IsAlphaNumeric = True
    Else
        IsAlphaNumeric = False
        Exit Function
    End If
 Next intCounter

End Function


Private Sub txtDSN_Change()
' Replaced the following:
'   If IsNumeric(txtDSN) Then
' with this: - EP 03/31/2021
If IsAlphaNumeric(txtDSN) Then
    cmdProcess.Enabled = True
Else
    cmdProcess.Enabled = False
End If

End Sub

Private Sub txtDSN_GotFocus()
SetSelected

End Sub

Private Sub txtDSN_KeyPress(KeyAscii As Integer)
'********************************************************************************
'* Name: txtDSN_KeyPress(KeyAscii As Integer)
'* Description:Cotrolling the Entry for the DSN
'*
'* Created: 6/14/99 3:39
'********************************************************************************
'
'Modified line below to allow search for alpha and numeric characters - EP 03/31/2021.
'If KeyAscii >= 48 And KeyAscii <= 57 Then  'Allow entry only from 0-9
If ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then  'Allow 0-9,a-z and A-Z
ElseIf KeyAscii = 8 Then 'Allow a BACKSPACE key.
Else
    KeyAscii = 0   ' This will cause the character to be ignored
End If

End Sub


Private Sub Form_Load()
'********************************************************************************
'* Name: Form_Load
'* Description:
'* Created: 6/01/99 11:10 AM
'********************************************************************************
On Error GoTo Form_LoadErr

Dim rsConfig As New ADODB.Recordset
Dim sSql As String
    
    sSql = "SELECT PA_BATCH_NAME,PF_BATCH_NAME FROM DD_CONFIG_INFO"
    rsConfig.Open sSql, gcnDDS, 0, 1
    msPFBatchName = rsConfig!PF_BATCH_NAME
    msAffBatchName = rsConfig!PA_BATCH_NAME
    
    
    Set rsConfig = Nothing
'AS - 8/28/2014 Ftp no longer used
'Set up FTP
'Set mobjFTP = Nothing
'Set mobjFTP = New FTPClass
'mobjFTP.RemoteSite = ReadIniFile(App.Path & "\" & App.EXEName & ".ini", "Startup", "SybaseData")
'mobjFTP.UserName = "pfsrpt"
'mobjFTP.Password = "arc$T3"
Set OutlookTitle1.Picture = fMainForm.imlToolbarIcons.ListImages("State Treasurer").Picture

cmdProcess.Enabled = False
'lblProcessStatus.Caption = "Enter the Deposit Sequence Number, then Press Tab Key "
imgDSN.Picture = LoadPicture()
'Set imgDSN.Picture = frmMain.imlToolbarIcons.ListImages("Deposit Sequence No").Picture
Set imgDSN.Picture = fMainForm.ImageList1.ListImages("DSN").Picture
lblProgress.Visible = False
lblProcess.Visible = False
ProgressBar1.Visible = False
ProgressBar2.Visible = False
sdcDSN.Visible = True
txtDSN.BackColor = vbWhite
dtStart.BackColor = PALE_YELLOW
txtDSN.Enabled = False
mbDSNRequired = False

If gobjLoginInfo.UserId = "epecounis" Then
    cmdForceSendStTreas.Visible = True
Else
    cmdForceSendStTreas.Visible = False
End If

    


'Make sure that the frame for selection prior days data is invisible

RefreshCombo

dtProcessDate.Text = Format$(Now, "MM/dd/yyyy")

'CALLING PROCEDURES
'Call EnableDSN

Xit:
    Hourglass False
    Exit Sub

Form_LoadErr:
    ShowUnexpectedError MODULE + "Form_Load", Err
    Resume Xit
End Sub

Private Sub Timer1_Timer()
'********************************************************************************
'* Name: Timer1_Timer
'* Description: Procedure to get a time running for the progress bar
'* Created: 7/21/99 12:40
'********************************************************************************
Dim i As Integer

ProgressBar2.value = intFlag

Select Case intFlag
  Case 1 To 35
          ProgressBar1.value = intFlag / 35 * 100
  Case 36 To 65
          ProgressBar1.value = ((intFlag - 36) / 29) * 100
  Case 66 To 100
          ProgressBar1.value = ((intFlag - 66) / 34) * 100
End Select
 
End Sub

Private Sub SetPercentage(ByVal iPercent As Integer)

intFlag = iPercent
DoEvents

End Sub
Private Sub cmdCancel_Click()
'Initial a value for the DSN .....
txtDSN.Text = "00000"
Unload Me

End Sub

Public Function PrintingSetupFiles() As Boolean
'********************************************************************************
'* Name:  PrintingSetupFiles()
'* Description: This will Call Crystall Report Files "SETUP" FILES, to Print
'* Parameters:
'* Created: 7/16/1999
'********************************************************************************
On Error GoTo PrintingSetupFilesErr

'File Name: Setup.dbf
'------------------------------------------------------------
'
'This is a Crystal Report Coding
Dim IntRet As Integer
Dim sCriteria As String

Dim cnData As New ADODB.Connection
Dim rsPAHeader As New ADODB.Recordset
Dim rsPADetail As New ADODB.Recordset
Dim rsPFDetail As New ADODB.Recordset

Dim rsPFHeader As New ADODB.Recordset
Dim crRpt As CRAXDRT.Report

'Open the Fox Pro Connection ....
'cnData.Open "DSN=DDSData"
On Error Resume Next
cnData.Open "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & gsDataPath & "\dd.mdb"
'Try opening with Jet 4.0
If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo PrintingSetupFilesErr
    cnData.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gsDataPath & "\dd.mdb"
End If
On Error GoTo PrintingSetupFilesErr
rsPAHeader.Open "SELECT * FROM " & msAffBatchName & "H", cnData, adOpenStatic
rsPFHeader.Open "SELECT * FROM " & msPFBatchName & "H", cnData, adOpenStatic
If rsPAHeader.EOF Or rsPFHeader.EOF Then
    MsgBox "Error locating header records.", vbCritical
    Exit Function
End If

rsPADetail.Open "SELECT * FROM " & msAffBatchName, cnData, adOpenStatic
SetPercentage 75
Set crRpt = gcrApp.OpenReport(gsDataPath & "\rSetup.RPT", crOpenReportByTempCopy)
'Print Report for Patient Account
'crRpt.Database.Tables.Item(1)..DataFiles(0) = gsDataPath & "\" & msAffBatchName & ".dbf"
crRpt.ParameterFields.Item(1).AddCurrentValue Format$(rsPAHeader!tdate, "MM/dd/yyyy")
crRpt.ParameterFields.Item(2).AddCurrentValue Format$(rsPAHeader!pdate, "MM/dd/yyyy")
crRpt.ParameterFields.Item(3).AddCurrentValue CDbl(rsPAHeader!tot_amt)
crRpt.Database.SetDataSource rsPADetail
crRpt.ReadRecords

crRpt.PrintOut False
Set crRpt = Nothing
SetPercentage 85
'Print Report for Personal Funds

rsPFDetail.Open "SELECT * FROM " & msPFBatchName, cnData, adOpenStatic
Set crRpt = gcrApp.OpenReport(gsDataPath & "\SetupPF.RPT", crOpenReportByTempCopy)
'crRpt.DataFiles(0) = gsDataPath & "\" & msPFBatchName & ".dbf"
crRpt.ParameterFields.Item(1).AddCurrentValue Format$(rsPFHeader!tdate, "MM/dd/yyyy")
crRpt.ParameterFields.Item(2).AddCurrentValue Format$(rsPFHeader!pdate, "MM/dd/yyyy")
crRpt.ParameterFields.Item(3).AddCurrentValue CDbl(rsPFHeader!tot_amt)
'Trigger the event at run time
crRpt.Database.SetDataSource rsPFDetail
crRpt.ReadRecords

crRpt.PrintOut False

PrintingSetupFiles = True
SetPercentage 95
'Rest the Mouse Pointer to the default
Xit:

Set rsPAHeader = Nothing
Set rsPFHeader = Nothing
Set rsPADetail = Nothing
Set rsPFDetail = Nothing
Set cnData = Nothing

Exit Function
PrintingSetupFilesErr:
ShowError MODULE + " Printing Files ", Err
Resume Xit
End Function



Public Function CreateSetupFiles() As Integer
'********************************************************************************
'* Name:   GetTotalTransactionsAmount()
'* Description: To Get the total value of the Transaction amount for PA Amount
'* Parameters:
'* Created: 7/16/1999
'********************************************************************************
On Error GoTo CreateSetupFilesErr
Dim rsTransaction As New ADODB.Recordset
Dim rsInst As New ADODB.Recordset
Dim dbData As DAO.Database
Dim rsPASetup As DAO.Recordset
Dim rsPFSetup As DAO.Recordset
Dim rsPASetupH As DAO.Recordset
Dim rsPFSetupH As DAO.Recordset

Dim sSql As String
Dim dtUseDate As Date
Dim ix As Integer
Dim dPAAmount As Double
Dim dPFAmount As Double
Dim dTotalPAAmount As Double
Dim dTotalPFAmount As Double
Dim sConnect As String
Dim rsConfig As New ADODB.Recordset
Dim iPAFile As Integer
Dim iPFFile As Integer
'Assign room for 20 details
Dim sPFDetails(20) As String
Dim iDtl As Integer
Dim sLine As String
Dim sAmt As String
        
    rsConfig.Open "SELECT * FROM DD_CONFIG_INFO", gcnDDS, adOpenForwardOnly
    If rsConfig.EOF Then
        MsgBox "Error reading configuration record.", vbCritical
        Exit Function
    End If

    Set dbData = OpenDatabase(gsDataPath & "\" & DD_DATABASE_NAME)
    
' 11/18/2003 Remved the links to Fox Pro Tables.  We don't need them anymore.
'    If LinkFoxProTable(dbData, msPFBatchName) = False Then
'        MsgBox "Could not link to Fox Pro table"
'        Exit Function
'    End If
'    If LinkFoxProTable(dbData, msPFBatchName & "H") = False Then
'        MsgBox "Could not link to Fox Pro table"
'        Exit Function
'    End If
'    If LinkFoxProTable(dbData, msAffBatchName) = False Then
'        MsgBox "Could not link to Fox Pro table"
'        Exit Function
'    End If
'    If LinkFoxProTable(dbData, msAffBatchName & "H") = False Then
'        MsgBox "Could not link to Fox Pro table"
'        Exit Function
'    End If
    
    
SetPercentage 5
    
    'delete all records in the setup.dbf and setuph.dbf tables for Patient and personal Funds
    sSql = "DELETE FROM " & msPFBatchName
    dbData.Execute sSql
    
    sSql = "DELETE FROM " & msPFBatchName & "H"
    dbData.Execute sSql
    
    sSql = "DELETE FROM " & msAffBatchName
    dbData.Execute sSql
    
    sSql = "DELETE FROM " & msAffBatchName & "H"
    dbData.Execute sSql
    
    
    
SetPercentage 7
    iPFFile = FreeFile()
    Open gsDataPath & "\OSTMHLT.txt" For Output As iPFFile
    iPAFile = FreeFile()
    Open gsDataPath & "\OSTMHPA.txt" For Output As iPAFile
    
    Set rsPASetup = dbData.OpenRecordset("SELECT * FROM " & msAffBatchName, dbOpenDynaset)
    Set rsPFSetup = dbData.OpenRecordset("SELECT * FROM " & msPFBatchName, dbOpenDynaset)
    Set rsPASetupH = dbData.OpenRecordset("SELECT * FROM " & msAffBatchName & "H", dbOpenDynaset)
    Set rsPFSetupH = dbData.OpenRecordset("SELECT * FROM " & msPFBatchName & "H", dbOpenDynaset)
    
    dtUseDate = CDate(dtStart.Text)
        
    'Sum up the total amount Get the PA and PF distribution amount
    sSql = "SELECT DD_POSTING_HISTORY.INSTITUTION_CODE as 'INST_CODE', DR_CR_FLAG, SUM(DD_POSTING_HISTORY.PF_DISTRIBUTION_AMT) AS 'TotalPFAmt', SUM(DD_POSTING_HISTORY.PA_DISTRIBUTION_AMT) AS 'TotalPAAmt'"
    sSql = sSql & " FROM DD_POSTING_HISTORY"
    sSql = sSql & " WHERE DD_POSTING_HISTORY.POSTED_DATETIME >= '" & dtUseDate & " 00:00:00" & "' And DD_POSTING_HISTORY.POSTED_DATETIME <= '" & dtUseDate & " 23:59:59" & "'"
    sSql = sSql & " GROUP BY DD_POSTING_HISTORY.INSTITUTION_CODE, DR_CR_FLAG"
    sSql = sSql & " ORDER BY DD_POSTING_HISTORY.INSTITUTION_CODE"
    
    'Open the recordset with sSql Statement
    rsTransaction.Open sSql, gcnDDS
    Set rsInst = New ADODB.Recordset
    sSql = "SELECT INSTITUTION_CODE,INSTITUTION_NAME,DD_VENDOR_ID_NUM FROM PF_INSTITUTION WHERE RECORD_STATUS = 'A'"
    rsInst.Open sSql, gcnPFS, adOpenStatic
    With rsTransaction
        
SetPercentage 9
    
    Do Until rsInst.EOF
        dPAAmount = 0
        dPFAmount = 0
        If .RecordCount > 0 Then .MoveFirst
        .Find "INST_CODE = '" & rsInst!INSTITUTION_CODE & "'", , adSearchForward
        'If there is no match then there is no transaction for the institution
        If .EOF Then
            'Do Nothing
        Else
            Select Case !DR_CR_FLAG
            Case "DR"
                dPAAmount = -1 * !TotalPAAmt
                dPFAmount = -1 * !TotalPFAmt
            Case "CR"
                dPAAmount = !TotalPAAmt
                dPFAmount = !TotalPFAmt
            Case Else
                MsgBox "The DR/CR Flag contains """ & !DR_CR_FLAG & """", vbCritical
                CreateSetupFiles = 0
                Exit Function
            End Select
            
            'See if we have another type of record for DR/CR
            .Find "INST_CODE = '" & rsInst!INSTITUTION_CODE & "'", 1, adSearchForward
            If Not .EOF Then
                Select Case !DR_CR_FLAG
                Case "DR"
                    dPAAmount = dPAAmount + (-1 * !TotalPAAmt)
                    dPFAmount = dPFAmount + (-1 * !TotalPFAmt)
                Case "CR"
                    dPAAmount = dPAAmount + !TotalPAAmt
                    dPFAmount = dPFAmount + !TotalPFAmt
                Case Else
                    MsgBox "The DR/CR Flag contains """ & !DR_CR_FLAG & """", vbCritical
                    CreateSetupFiles = 0
                    Exit Function
                End Select
            End If
        End If
            
        'IF the pa amount or to pf amount is greater than zero
        'and the institution has an unknown vendor then fail process
        If rsInst!DD_VENDOR_ID_NUM = "Unknown" Then
            If dPAAmount <> 0 Or dPFAmount <> 0 Then
                MsgBox rsInst!INSTITUTION_NAME & " is not set up with a valid vendor id.  Please fix this and resend", vbInformation
                CreateSetupFiles = 0
                Exit Function
            End If
        Else
            'Make records for detail record Setup file
        
            'Create entry for Personal Funds Setup
            If dPFAmount <> 0 Then
                iDtl = iDtl + 1
                sLine = rsInst!DD_VENDOR_ID_NUM & String(20 - Len(rsInst!DD_VENDOR_ID_NUM), " ")
                sLine = sLine & "," & rsInst!INSTITUTION_NAME & String(40 - Len(rsInst!INSTITUTION_NAME), " ")
                sLine = sLine & "," & String(20, " ")
                sAmt = Format$(dPFAmount * 100, String(11, "0"))
                sLine = sLine & "," & sAmt
                sLine = sLine & Space(26)
                
                sPFDetails(iDtl) = sLine
                Debug.Print sLine
                                
            
            End If
            
            rsPFSetup.AddNew
            rsPFSetup!Vendor = rsInst!DD_VENDOR_ID_NUM
            rsPFSetup!payee = Trim$(Left$(rsInst!INSTITUTION_NAME, 25))
            rsPFSetup!stifno = "" 'This is Blank field per the Requirements
            rsPFSetup!AMOUNT = CDbl(dPFAmount)
            rsPFSetup.Update
        
            'Accumulate the accumulators
            dTotalPAAmount = dTotalPAAmount + dPAAmount
            dTotalPFAmount = dTotalPFAmount + dPFAmount
            'Move to the next institution
        End If
        
        rsInst.MoveNext
        If intFlag < 35 Then
            SetPercentage intFlag + 1
        End If
    Loop

    'Create entry for Patient Setup detail
    rsPASetup.AddNew
    rsPASetup!Vendor = rsConfig!PA_VENDOR_ID_NUM
    'rsPASetup!payee = Trim$(Left$(rsInst!INSTITUTION_NAME, 25))
    rsPASetup!payee = "Mental Health Summary"
    
    rsPASetup!stifno = "" 'This is Blank field per the Requirements
    rsPASetup!AMOUNT = dTotalPAAmount
    rsPASetup.Update


    'Put in a record for the PA header files

    rsPASetupH.AddNew
    rsPASetupH!tdate = dtUseDate
    rsPASetupH!pdate = Format(dtProcessDate.Text, "MM/dd/yyyy")
    rsPASetupH!tot_amt = dTotalPAAmount
    rsPASetupH.Update

    'Create the file for PA Accounting
    'Create the header record
    sAmt = Format$(dTotalPAAmount * 100, String(15, "0"))
    sLine = "HRT,OSTMHPA,O," & Format$(dtUseDate, "MM/dd/yyyy") & "," & Format(dtProcessDate.Text, "MM/dd/yyyy") & "," & sAmt & String(69, " ")
    Print #iPAFile, sLine
                
    'Create the detail record
    sLine = rsConfig!PA_VENDOR_ID_NUM & String(20 - Len(rsConfig!PA_VENDOR_ID_NUM), " ")
    sLine = sLine & "," & "Mental Health Summary" & String(40 - Len("Mental Health Summary"), " ")
    sLine = sLine & "," & String(20, " ")
    sAmt = Format$(dTotalPAAmount * 100, String(11, "0"))
    sLine = sLine & "," & sAmt
    sLine = sLine & Space(26)
    Print #iPAFile, sLine


    'Create the Personal Funds Records
    'Put in a record for the PF header files
    sAmt = Format$(dTotalPFAmount * 100, String(15, "0"))
    sLine = "HRT,OSTMHLT,O," & Format$(dtUseDate, "MM/dd/yyyy") & "," & Format(dtProcessDate.Text, "MM/dd/yyyy") & "," & sAmt & String(69, " ")
    Print #iPFFile, sLine
    'Get the details
    For ix = 1 To 20
        If sPFDetails(ix) = "" Then
            Exit For
        End If
        Print #iPFFile, sPFDetails(ix)
    Next ix

    rsPFSetupH.AddNew
    rsPFSetupH!tdate = dtUseDate
    rsPFSetupH!pdate = Format(dtProcessDate.Text, "MM/dd/yyyy")
    rsPFSetupH!tot_amt = dTotalPFAmount
    rsPFSetupH.Update
    
    End With
    If iDtl > 0 Then
        If dTotalPAAmount <> 0 Then
            CreateSetupFiles = 3
        Else
            CreateSetupFiles = 1
        End If
    Else
        If dTotalPAAmount <> 0 Then
            CreateSetupFiles = 2
        End If
    End If
Xit:
    On Error Resume Next
    SetPercentage 35
    Set rsTransaction = Nothing
    Set rsInst = Nothing
    Set rsConfig = Nothing
    Set rsPASetup = Nothing
    Set rsPFSetup = Nothing
    Set rsPASetupH = Nothing
    Set rsPFSetupH = Nothing
    Set dbData = Nothing
    Close #iPFFile
    Close #iPAFile
    Exit Function

CreateSetupFilesErr:

    ShowError MODULE + "Create Setup Files", Err
    Resume Xit

End Function



Private Function LinkFoxProTable(dbsTemp As Database, strTable As String, Optional ByVal strSourceTable As String) As Boolean

'ConnectOutput(dbsTemp As Database, strTable As String, strConnect As String, _   strSourceTable As String)
   
   Dim tdfLinked As TableDef
   Dim intTemp As Integer   ' Create a new TableDef, set its Connect and
   Dim strConnect As String
   ' SourceTableName properties based on the passed
   ' arguments, and append it to the TableDefs collection.
    strConnect = "FoxPro 2.0;DATABASE=" & gsDataPath
    If strSourceTable = "" Then
        strSourceTable = strTable
    End If
On Error Resume Next
   dbsTemp.TableDefs.Delete strTable

On Error GoTo FoxProErr
   Set tdfLinked = dbsTemp.CreateTableDef(strTable)
   tdfLinked.Connect = strConnect
   tdfLinked.SourceTableName = strSourceTable
   dbsTemp.TableDefs.Append tdfLinked
   Set tdfLinked = Nothing
   LinkFoxProTable = True

Exit Function
FoxProErr:
    LinkFoxProTable = False

End Function
Private Function CreateALOGS() As Boolean
'AS - 10/3/2014 Modified to specify excel version of 2003
Dim oExcel As New Excel.Application
Dim oBook As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim rsSeq As New ADODB.Recordset
Dim rsInst As New ADODB.Recordset
Dim sSql As String
Dim sRegionHold As String
Dim sLastWorksheet As String
Dim rsTransactions As New ADODB.Recordset
Dim iRow As Integer
Dim sPostedDate As String
Dim sSeqNum As String
Dim sSentToDate As String

On Error GoTo CreateALOGSErr

sPostedDate = dtStart.Text

If mbIncompleteSend = False Then
    'We need to get the sequence number and date
    sSql = "SELECT DISTINCT DD_DEP_SEQ_NO.DEP_SEQ_NUM, DD_DEP_SEQ_NO.CREATED_DATETIME,DD_DEP_SEQ_NO.PROCESS_DATE"
    sSql = sSql & " FROM DD_POSTING_HISTORY, DD_DEP_SEQ_NO"
    sSql = sSql & " WHERE DD_POSTING_HISTORY.DEP_SEQ_NUM = DD_DEP_SEQ_NO.DEP_SEQ_NUM"
    sSql = sSql & " AND DD_POSTING_HISTORY.POSTED_DATETIME >= '" & sPostedDate & " 00:00:00'"
    sSql = sSql & " AND DD_POSTING_HISTORY.POSTED_DATETIME <= '" & sPostedDate & " 23:59:59'"
    sSql = sSql & " AND DD_DEP_SEQ_NO.CREATED_DATETIME >= '" & sPostedDate & " 00:00:00'"
    rsSeq.Open sSql, gcnDDS, adOpenStatic
    If rsSeq.EOF Then
        sSentToDate = Format$(dtProcessDate.Text)
        sSeqNum = ""
    Else
        If IsNull(rsSeq!PROCESS_DATE) Then
            sSentToDate = Format$(rsSeq!CREATED_DATETIME, "MM/dd/yyyy")
        Else
            sSentToDate = Format$(rsSeq!PROCESS_DATE, "MM/dd/yyyy")
        End If
        sSeqNum = rsSeq!DEP_SEQ_NUM
    End If
Else
    'Use the values currently on screen
    sSentToDate = Format(dtProcessDate.Text, "MM/dd/yyyy")
    sSeqNum = txtDSN.Text
End If


sSql = "SELECT * FROM PF_INSTITUTION ORDER BY DD_NCAS_REGION"
rsInst.Open sSql, gcnPFS, adOpenStatic
oExcel.DisplayAlerts = False
Kill gsDataPath & "\ddalog-*.xls"

Do Until rsInst.EOF
    'Create a new Excel Spreadsheet for each region
    If Not IsNull(rsInst!DD_NCAS_REGION) Then
        Set rsTransactions = Nothing
        'Sum up the total amount Get the PA and PF distribution amount
        sSql = "SELECT DISTINCT DD_POSTING_HISTORY.INSTITUTION_CODE,DD_INCOME_SOURCE_TYPE.NCAS_ACCOUNT,DD_INCOME_SOURCE_TYPE.PA_INCOME_SRC_TYPE, SUM(CASE WHEN DR_CR_FLAG = 'DR' THEN -1 WHEN DR_CR_FLAG = 'CR' THEN 1 END * DD_POSTING_HISTORY.PF_DISTRIBUTION_AMT) AS 'TotalPFAmt', SUM(CASE WHEN DR_CR_FLAG = 'DR' THEN -1 WHEN DR_CR_FLAG = 'CR' THEN 1 END * DD_POSTING_HISTORY.PA_DISTRIBUTION_AMT) AS 'TotalPAAmt'"
        sSql = sSql & " FROM DD_POSTING_HISTORY, DD_INCOME_SOURCE_TYPE"
        sSql = sSql & " WHERE DD_POSTING_HISTORY.INCOME_SOURCE_TYPE_ID = DD_INCOME_SOURCE_TYPE.INCOME_SOURCE_TYPE_ID AND DD_POSTING_HISTORY.INSTITUTION_CODE = '" & rsInst!INSTITUTION_CODE & "' AND DD_POSTING_HISTORY.POSTED_DATETIME >= '" & sPostedDate & " 00:00:00" & "' And DD_POSTING_HISTORY.POSTED_DATETIME <= '" & sPostedDate & " 23:59:59" & "'"
        sSql = sSql & " GROUP BY DD_POSTING_HISTORY.INSTITUTION_CODE, DD_INCOME_SOURCE_TYPE.NCAS_ACCOUNT, DD_INCOME_SOURCE_TYPE.PA_INCOME_SRC_TYPE"
    
        'Open the recordset with sSql Statement
        rsTransactions.Open sSql, gcnDDS
        If Not rsTransactions.EOF Then
            If sRegionHold <> rsInst!DD_NCAS_REGION Then
                If sRegionHold <> "" Then
                    oBook.Worksheets("ALOG").Delete
                    'oBook.SaveAs gsDataPath & "\ddalog-" & sRegionHold & ".xls", 56
                    oBook.SaveAs gsDataPath & "\ddalog-" & sRegionHold & ".xls"
                    
                    oBook.Close
                End If
                sLastWorksheet = "ALOG"
                Set oBook = Nothing
                sRegionHold = rsInst!DD_NCAS_REGION
                Set oBook = oExcel.Workbooks.Add(gsDataPath & "\alog.xlt")
            End If
        
            oBook.Worksheets("ALOG").Copy After:=Worksheets(sLastWorksheet)
            Set oSheet = oBook.Sheets("ALOG (2)")
            sLastWorksheet = rsInst!AFFINITY_DB_NAME
            oSheet.Name = sLastWorksheet
            oSheet.Activate
            
            'Fill in the fields needed
            With oSheet
            .Cells(2, 2) = UCase(rsInst!INSTITUTION_NAME)
            .Cells(1, 10) = sSeqNum
            .Cells(3, 10) = sSentToDate
            .Cells(21, 2) = Format$(Now, "MM/dd/yyyy")
            
            'Enter in the transaction data
            iRow = 8
            Do Until rsTransactions.EOF
                
                Select Case rsTransactions!PA_INCOME_SRC_TYPE
                Case "S1"
                    .Cells(iRow, 5) = "Social Security"
                Case "V1"
                    .Cells(iRow, 5) = "Veteran's Administration"
                Case "R1"
                    .Cells(iRow, 5) = "Railroad Retirement"
                Case "C1"
                    .Cells(iRow, 5) = "Civil Service"
                Case "A1"
                    .Cells(iRow, 5) = "Armed Services"
                Case "P1"
                    .Cells(iRow, 5) = "Misc Federal"
                Case "I1"
                    'Add a row line
                    iRow = iRow + 1
                    .Cells(iRow, 5) = "Supplemental Social Security"
                Case Else
                    .Cells(iRow, 5) = rsTransactions!PA_INCOME_SRC_TYPE
                End Select
                
                .Cells(iRow, 1) = rsTransactions!NCAS_ACCOUNT
                
                If rsTransactions!PA_INCOME_SRC_TYPE = "I1" Then
                    .Cells(iRow, 2) = "801"
                Else
                    .Cells(iRow, 2) = rsInst!DD_NCAS_CENTER
                End If
                
                .Cells(iRow, 8) = CCur(rsTransactions!TotalPAAmt)
                .Cells(iRow, 9) = CCur(rsTransactions!TotalPFAmt)
                .Cells(iRow, 13) = "DIRECT DEPOSIT"
                
                'Advance the row
                iRow = iRow + 1
                rsTransactions.MoveNext
            Loop
            End With
            
        End If
    End If
    rsInst.MoveNext
Loop
'Save the ending excel workbook

oBook.Worksheets("ALOG").Delete
oBook.SaveAs gsDataPath & "\ddalog-" & sRegionHold & ".xls"
oBook.Close

CreateALOGS = True

Xit:
oExcel.DisplayAlerts = True
Set rsInst = Nothing
Set rsSeq = Nothing
Set rsTransactions = Nothing
Set oSheet = Nothing
Set oBook = Nothing
oExcel.Quit
Set oExcel = Nothing

Exit Function
CreateALOGSErr:
If Err = 53 Then
    Resume Next
Else
    MsgBox Error, vbInformation
    CreateALOGS = False
    Resume Xit
    
End If

End Function
