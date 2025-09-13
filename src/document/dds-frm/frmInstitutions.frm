VERSION 5.00
Object = "{D9D1F94F-AEDB-11D2-9C3C-00105A19BCF2}#1.0#0"; "OAOTitle.ocx"
Begin VB.Form frmInstitutions 
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   9930
   WindowState     =   2  'Maximized
   Begin VB.Frame fraDetails 
      Caption         =   "Details"
      Height          =   5415
      Left            =   4095
      TabIndex        =   6
      Top             =   720
      Width           =   5625
      Begin VB.TextBox txtDDDBName 
         Height          =   285
         Left            =   2670
         MaxLength       =   20
         TabIndex        =   19
         Top             =   1995
         Width           =   2700
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
         Height          =   285
         HelpContextID   =   30037
         Left            =   2685
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   16
         Top             =   675
         WhatsThisHelpID =   30036
         Width           =   2700
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   285
         HelpContextID   =   30037
         Left            =   2685
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1110
         WhatsThisHelpID =   30036
         Width           =   2700
      End
      Begin VB.TextBox txtDDSendReportTo 
         Enabled         =   0   'False
         Height          =   1485
         HelpContextID   =   30037
         Left            =   2670
         Locked          =   -1  'True
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   2430
         WhatsThisHelpID =   30036
         Width           =   2700
      End
      Begin VB.TextBox txtDDVendorIDNumber 
         Height          =   285
         Left            =   2670
         MaxLength       =   15
         TabIndex        =   18
         Top             =   1545
         Width           =   2700
      End
      Begin VB.TextBox txtModifiedBy 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4725
         Width           =   1935
      End
      Begin VB.TextBox txtModifiedDate 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MMMM d, yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   4965
         Width           =   1935
      End
      Begin VB.TextBox txtCreatedBy 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   4125
         Width           =   1935
      End
      Begin VB.TextBox txtCreatedDate 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MMMM d, yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4365
         Width           =   1935
      End
      Begin VB.Label lblDDDBName 
         Caption         =   "DD Database Name:"
         Height          =   255
         Left            =   315
         TabIndex        =   25
         Top             =   2010
         Width           =   2055
      End
      Begin VB.Label lblCode 
         Caption         =   "Institution Code:"
         Height          =   210
         Left            =   330
         TabIndex        =   24
         Top             =   705
         Width           =   2250
      End
      Begin VB.Label lblName 
         Caption         =   "Institution Name:"
         Height          =   210
         Left            =   330
         TabIndex        =   23
         Top             =   1125
         Width           =   2250
      End
      Begin VB.Label lblSendReportTo 
         Caption         =   "Send Report to Email address:"
         Height          =   210
         Left            =   315
         TabIndex        =   22
         Top             =   2475
         Width           =   2250
      End
      Begin VB.Label lblddvendoridnum 
         Caption         =   "DD Vendor ID Number:"
         Height          =   255
         Left            =   315
         TabIndex        =   21
         Top             =   1545
         Width           =   2055
      End
      Begin VB.Label lblCreatedDate 
         Caption         =   "Record Created On:"
         Height          =   255
         Left            =   1995
         TabIndex        =   15
         Top             =   4365
         Width           =   1485
      End
      Begin VB.Label lblCreatedBy 
         Caption         =   "Record Created By:"
         Height          =   255
         Left            =   1995
         TabIndex        =   14
         Top             =   4125
         Width           =   1485
      End
      Begin VB.Label lblModifiedBy 
         Caption         =   "Record Modified By:"
         Height          =   255
         Left            =   1995
         TabIndex        =   7
         Top             =   4725
         Width           =   1485
      End
      Begin VB.Label lblModifiedDate 
         Caption         =   "Record Modified On:"
         Height          =   255
         Left            =   1995
         TabIndex        =   8
         Top             =   4965
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   8550
      TabIndex        =   3
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7350
      TabIndex        =   2
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Frame fraCodes 
      Caption         =   "Codes"
      Height          =   5415
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   2055
      Begin VB.ListBox lstCodes 
         Height          =   4740
         ItemData        =   "frmInstitutions.frx":0000
         Left            =   225
         List            =   "frmInstitutions.frx":0002
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   930
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6150
      TabIndex        =   4
      Top             =   6480
      Width           =   1095
   End
   Begin OAOTitle.OutlookTitle OutTitle 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   979
      ForeColor       =   16777215
      Picture         =   "frmInstitutions.frx":0004
      Caption         =   "Institutions"
   End
   Begin VB.Line lin1 
      BorderColor     =   &H80000010&
      BorderStyle     =   6  'Inside Solid
      X1              =   360
      X2              =   9675
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line lin2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   360
      X2              =   9660
      Y1              =   6375
      Y2              =   6360
   End
End
Attribute VB_Name = "frmInstitutions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ********************************************************************************
' * Description:
' * The Institutions Screen is used to view, add, edit, and delete Institution
'*  codes from the system. This screen serves 4 primary goals:
' *  1) Display all active Institution codes in the system (secure)
' *  2) Add a new Institution code (secure)
' *  3) Edit an existing Institution code (secure)
' *  4) Delete an existing Institution code (secure)
' *
' * Methods:
' *
' * Properties:
' *
' * Associations:
' *
' * Events:
' *
' * Revisions:
' *  1/22/99    bsr Added comments.
' *  8/13/99 - Added two fields to Institution Maintenance
' *            UpdateFields() - fill in the value for DD Vendor ID Number and DD Database Name
'              DataValidate() - check DD Vendor ID Number for Nulls
'              iuInstitution() - fill parameter fields for DD Vendor ID Number and DD Database Name
'              ChangeScreenMode - For add mode enable and blank both fields
'                                 For Edit Mode change background to yellow for DD Vendor ID Number.  Enable both fields for entry.
'               GotFocus - change background color to blue when the fields get focus
' *
' ********************************************************************************
' Mod INTERFACES
' Mod EVENTS

' Mod CONSTANTS
Private Const INSERT_STR As String = "I"
Private Const UPDATE_STR As String = "U"
Private Const MODULE As String = "Institutions"

' Mod ENUMS
' Mod TYPES

' Mod DECLARES
Private mcmdInstitution As New ADODB.Command
Private mrsInstitution As New ADODB.Recordset

' Mod VARIABLES
Private msUpdateMode As String
Private msMessage As String
Private msTitle As String
Private Style As VbMsgBoxStyle
Private dblInstitutionID As Double
Private msLastSelected As String

Private Sub Form_Activate()

    'Set the apropriate buttons for the main toolbar
    Call fMainForm.SetMainToolbar(False)

End Sub

Private Sub Form_Load()
'********************************************************************************
'* Name: Form_Load
'*
'* Description:
'* Created: 1/22/99 3:55:25 PM
'********************************************************************************
    
    ChangeScreenMode (VIEW_MODE)
End Sub

Private Sub cmdEdit_Click()
'********************************************************************************
'* Name: cmdEdit_Click
'*
'* Description:
'* Created: 1/22/99 4:08:04 PM
'********************************************************************************
    msUpdateMode = UPDATE_STR
    ChangeScreenMode (EDIT_MODE)
End Sub

Private Sub cmdApply_Click()

On Error GoTo cmdApply_ClickErr

'********************************************************************************
'* Name: cmdApply_Click
'*
'* Description:
'* Created: 1/22/99 4:02:11 PM
'********************************************************************************
    msLastSelected = txtCode
    Hourglass True
    'check to see if this code already exists
    If fnCkCodeExists = False Then
        'this code doesn't exist so check that the fields are filled out properly
        If DataValidation = True Then
            'the fields are filled out so execute procedure to update the record
            Call iuInstitution
            'reset the screen to view mode
            ChangeScreenMode (VIEW_MODE)
            SelectLastSelected
        Else
        'the fields are not filled out properly so don't execute the procedure
        End If
    Else
        'this code already exists so don't do anything else
    End If

Xit:
    Hourglass False
    Exit Sub

cmdApply_ClickErr:
    ShowUnexpectedError MODULE + "cmdApply_Click", Err
    Resume Xit


End Sub

Private Sub cmdOK_Click()

On Error GoTo cmdOK_ClickErr

'********************************************************************************
'* Name: cmdOK_Click
'*
'* Description:
'* Created: 1/22/99 4:12:47 PM
'********************************************************************************

    'check to see if this code already exists
    Hourglass True
    If fnCkCodeExists = False Then
        'this code doesn't exist so check that the fields are filled out properly
        If DataValidation = True Then
            'the fields are filled out so execute procedure to update the record
            Call iuInstitution
            'now close the form
            Unload Me
        Else
        'the fields are not filled out properly so don't execute the procedure
        End If
    Else
        'this code already exists so don't do anything else
    End If

Xit:
    Hourglass False
    Exit Sub

cmdOK_ClickErr:
    ShowUnexpectedError MODULE + "cmdOk_Click", Err
    Resume Xit


End Sub

Private Sub cmdCancel_Click()

On Error GoTo cmdCancel_ClickErr

'********************************************************************************
'* Name: cmdCancel_Click
'*
'* Description:
'* Created: 1/22/99 4:12:57 PM
'********************************************************************************
    msLastSelected = txtCode
    Style = vbYesNo + vbQuestion + vbDefaultButton2 + vbApplicationModal
    msMessage = "Are you sure you want to cancel?"
        
    Select Case msUpdateMode
        
        Case INSERT_STR
            If MsgBox(msMessage, Style) = vbNo Then
                'go back
            Else
                'cancel and lose changes
                ChangeScreenMode (VIEW_MODE)
                SelectLastSelected
            End If

        Case UPDATE_STR
            If MsgBox(msMessage, Style) = vbNo Then
                'go back
            Else
                'cancel and lose changes
                ChangeScreenMode (VIEW_MODE)
                SelectLastSelected
            End If
            
        Case vbNullString
            'cancel and lose changes
            Unload Me
    End Select

Xit:
    Exit Sub

cmdCancel_ClickErr:
    ShowUnexpectedError MODULE + "cmdCancel_Click", Err
    Resume Xit


End Sub
Private Sub SelectLastSelected()
'*************************************************
'* Selects the last entered
'************************************************
Dim iX As Long
    For iX = 0 To lstCodes.ListCount - 1
        If lstCodes.List(iX) = msLastSelected Then
            lstCodes.Selected(iX) = True
            lstCodes.SetFocus
            Exit For
        End If
    Next iX
    
End Sub

Private Sub UpdateList()

On Error GoTo UpdateListErr

'********************************************************************************
'* Name: UpdateList
'*
'* Description:
'* Created: 1/22/99 4:16:20 PM
'********************************************************************************
    Hourglass True
    lstCodes.Clear
    lstCodes.Enabled = True
    Set mcmdInstitution.ActiveConnection = gcnPFS
    With mcmdInstitution
    .CommandType = adCmdText
    .CommandText = "SELECT INSTITUTION_ID,INSTITUTION_CODE,INSTITUTION_NAME,NCAS_CO_NUM,NCAS_CODE,STD_ALLOW_TIMEFRAME_CODE,SYS1099_HOSPITAL_CODE,SKD_RPT_NTWK_PATH, DD_VENDOR_ID_NUM, DD_SEND_REPORT_TO, i.AFFINITY_DB_NAME,i.CREATED_BY,i.CREATED_DATETIME,i.LAST_MOD_BY,i.LAST_MOD_DATE FROM PF_INSTITUTION i , PF_STD_ALLOW_TMEFRM t, PF_NCAS_ACCTG_DATA a WHERE i.RECORD_STATUS = " & "'A" & "' AND i.STD_ALLOW_TIMEFRAME_ID = t.STD_ALLOW_TIMEFRAME_ID AND i.NCAS_ACCTG_DATA_ID = a.NCAS_ACCTG_DATA_ID ORDER BY INSTITUTION_CODE"
    Set mrsInstitution = .Execute
    End With
    With mrsInstitution
    Do Until .EOF
        lstCodes.AddItem !INSTITUTION_CODE
        .MoveNext
    Loop
    .MoveFirst
    End With
    Call UpdateFields

Xit:
    Hourglass False
    Exit Sub

UpdateListErr:
    ShowUnexpectedError MODULE + "UpdateList", Err
    Resume Xit


End Sub

Private Sub UpdateFields()

On Error GoTo UpdateFieldsErr

'********************************************************************************
'* Name: UpdateFields
'*
'* Description:
'* Created: 1/22/99 4:16:44 PM
'********************************************************************************

    With mrsInstitution
        txtCode = !INSTITUTION_CODE
        txtName = !INSTITUTION_NAME
        txtDDVendorIDNumber = !DD_VENDOR_ID_NUM
        txtDDDBName = ConvertNull(!AFFINITY_DB_NAME)
        txtDDSendReportTo = ConvertNull(!DD_SEND_REPORT_TO)
        txtCreatedBy = !CREATED_BY
        txtCreatedDate = !CREATED_DATETIME
        txtModifiedBy = ConvertNull(!LAST_MOD_BY)
        txtModifiedDate = ConvertNull(!LAST_MOD_DATE)
        dblInstitutionID = !INSTITUTION_ID
    End With

Xit:
    Exit Sub

UpdateFieldsErr:
    ShowUnexpectedError MODULE + "UpdateFields", Err
    Resume Xit


End Sub

Private Sub lstCodes_Click()

On Error GoTo lstCodes_ClickErr

'********************************************************************************
'* Name: lstCodes_Click
'*
'* Description:
'* Created: 1/22/99 4:17:19 PM
'********************************************************************************

    With mrsInstitution
        .MoveFirst
        While .Fields("INSTITUTION_CODE") <> lstCodes.Text
            .MoveNext
        Wend
    End With
    Call UpdateFields
    cmdEdit.Enabled = True

Xit:
    Exit Sub

lstCodes_ClickErr:
    ShowUnexpectedError MODULE + "lstCodes_Click", Err
    Resume Xit


End Sub

Private Function fnCkCodeExists() As Boolean

On Error GoTo fnCkCodeExistsErr

'********************************************************************************
'* Name: fnCkCodeExists
'*
'* Description:
'* Created: 1/22/99 4:18:02 PM
'********************************************************************************

    If msUpdateMode = INSERT_STR Then
    'if trying to insert a new code then see if it already exists, otherwise it doesn't matter
        Dim bExists As Boolean
        With mrsInstitution
            bExists = False
            .MoveFirst
            Do Until .EOF
                If .Fields("INSTITUTION_CODE") <> Trim(UCase(txtCode.Text)) Then
                    bExists = False
                    .MoveNext
                Else
                    bExists = True
                    Exit Do
                End If
            Loop
            If bExists = True Then
                Style = vbOKOnly + vbExclamation + vbApplicationModal
                msTitle = "Code Exists"
                msMessage = "This code exists, please enter a different value"
                Hourglass False
                If MsgBox(msMessage, Style, msTitle) = vbOK Then
                    fnCkCodeExists = True
                    'set focus to code entry field
                    txtCode.SetFocus
                End If
                Hourglass True
            Else
                fnCkCodeExists = False
            End If
        End With
    Else
        fnCkCodeExists = False
    End If

Xit:
    Exit Function

fnCkCodeExistsErr:
    ShowUnexpectedError MODULE + "fnCkCodeExists", Err
    Resume Xit


End Function

Private Function DataValidation() As Boolean

On Error GoTo DataValidationErr

'********************************************************************************
'* Name: DataValidation
'*
'* Description:
'* Created: 1/22/99 4:18:46 PM
'********************************************************************************

    Dim strEmptyFields As String
    Style = vbOKOnly + vbExclamation + vbApplicationModal
    msMessage = "The following data is required:    " & vbCrLf & vbCrLf
    msTitle = "Invalid Data"
    
    If txtCode.Text = vbNullString Then
        strEmptyFields = " Code"
        txtCode.SetFocus
    End If
    
    If txtName.Text = vbNullString And strEmptyFields <> vbNullString Then
        strEmptyFields = strEmptyFields & vbCrLf & " Name"
    Else
        If txtName.Text = vbNullString Then
        strEmptyFields = strEmptyFields & " Name"
        txtName.SetFocus
        End If
    End If
          
    
    If txtDDVendorIDNumber.Text = vbNullString And strEmptyFields <> vbNullString Then
        strEmptyFields = strEmptyFields & vbCrLf & " DD Vendor ID Number"
    Else
        If txtDDVendorIDNumber.Text = vbNullString Then
        strEmptyFields = strEmptyFields & " DD Vendor Id Number"
        txtDDVendorIDNumber.SetFocus
        End If
    End If
    
    
    If strEmptyFields <> vbNullString Then
        msMessage = msMessage & strEmptyFields & vbCrLf
        If MsgBox(msMessage, Style, msTitle) = vbOK Then
            DataValidation = False
        End If
    Else
        DataValidation = True
    End If


Xit:
    Exit Function

DataValidationErr:
    ShowUnexpectedError MODULE + "DataValidation", Err
    Resume Xit


End Function

Private Sub iuInstitution()

On Error GoTo iuInstitutionErr

'********************************************************************************
'* Name: iuInstitution
'*
'* Description:
'* Created: 1/22/99 4:20:26 PM
'********************************************************************************

        Dim cmd As New ADODB.Command
        Set cmd.ActiveConnection = gcnDDS
        If gStoredProcs("up_u_Institution_DDS").GetStoredProcCommand(cmd) = True Then
            cmd.Parameters("institution_code") = Trim(UCase(txtCode.Text))
            cmd.Parameters("institution_name") = Trim(txtName.Text)
            cmd.Parameters("dd_vendor_id_num") = Trim(txtDDVendorIDNumber.Text)
            If Trim$(txtDDDBName.Text) = vbNullString Then
                cmd.Parameters("affinity_db_name") = Null
            Else
                cmd.Parameters("affinity_db_name") = Trim$(txtDDDBName.Text)
            End If
            If Trim$(txtDDSendReportTo.Text) = vbNullString Then
                cmd.Parameters("dd_send_report_to") = Null
            Else
                cmd.Parameters("dd_send_report_to") = Trim$(txtDDSendReportTo.Text)
            End If
            cmd.Parameters("user_id") = gobjLoginInfo.UserId
            cmd.Parameters("update_mode") = msUpdateMode
            cmd.Parameters("institution_id") = dblInstitutionID
            cmd.Parameters("called_from_another_proc") = "N"
            cmd.Execute
            If cmd.Parameters("RETURN_VALUE") <> 0 Then
                GetServerErrorMsg cmd.Parameters("RETURN_VALUE"), "Error occurred adding or updating the Institution record."
            End If
        Else
            MsgBox "Error creating the Insert/Update Institution Stored Procedure.", vbCritical
            Set cmd = Nothing
            ExitApp
        End If
        Set cmd = Nothing

Xit:
    Exit Sub

iuInstitutionErr:
    ShowUnexpectedError MODULE + "iuInstitution", Err
    Resume Xit
    

End Sub



Private Sub ChangeScreenMode(ByVal iMode As ScreenMode)
On Error GoTo ChangeScreenModeErr
'********************************************************************************
'* Name: ChangeScreenMode
'*
'* Description:
'*   This subroutine will change the background of certain controls and enable
'*   or disable controls and buttons depending on whether you are adding a code,
'*   editing a code, or viewing the active codes.
'
'* Parameters: iMode - The choices are ADD_MODE, VIEW_MODE or UPDATE_MODE
'* Created: 5/23/00 10:51 AM
'********************************************************************************
    'Change the mode for the screen
    
    Select Case iMode
    
    Case VIEW_MODE
        'restore the background of the controls to white while in view mode
        txtCode.BackColor = DFLT_WHITE
        txtName.BackColor = DFLT_WHITE
        txtDDVendorIDNumber.BackColor = DFLT_WHITE
        'enable the controls that are modifiable during view
        'set focus
        'disable the controls that cannot be changed during view
        txtCode.Enabled = False
        txtCode.Locked = True
        txtName.Enabled = False
        txtName.Locked = True
        txtDDVendorIDNumber.Enabled = False
        txtDDVendorIDNumber.Locked = True
        txtDDDBName.Enabled = False
        txtDDDBName.Locked = True
        txtDDSendReportTo.Enabled = False
        txtDDSendReportTo.Locked = True
        'disable all buttons not available during view
        cmdEdit.Enabled = False
        cmdApply.Enabled = False
        cmdOK.Enabled = False
        Call UpdateList
        msUpdateMode = vbNullString
    
    Case EDIT_MODE
        'change the background of the controls that are mandatory for an add
        txtCode.BackColor = PALE_YELLOW
        txtName.BackColor = PALE_YELLOW
        txtDDVendorIDNumber.BackColor = PALE_YELLOW
        'enable the controls that are modifiable during an add
        txtName.Enabled = True
        txtName.Locked = False
        txtDDVendorIDNumber.Enabled = True
        txtDDVendorIDNumber.Locked = False
        txtDDDBName.Enabled = True
        txtDDDBName.Locked = False
        txtDDSendReportTo.Enabled = True
        txtDDSendReportTo.Locked = False
        'set focus to first control
        txtName.SetFocus
        'enable all buttons available during an edit
        cmdOK.Enabled = True
        cmdApply.Enabled = True
        'disable the controls that cannot be changed during an add
        lstCodes.Enabled = False
        'disable all buttons not available during an edit
        cmdEdit.Enabled = False
    
    End Select

Xit:
    Exit Sub
    
ChangeScreenModeErr:
    ShowError MODULE + ".ChangeScreenMode", Err
    Resume Xit
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'********************************************************************************
'* Name: Form_Unload
'*
'* Description:
'* Created: 1/27/99 3:44:43 PM
'********************************************************************************
On Error Resume Next
mrsInstitution.Close
Set mrsInstitution = Nothing
Set mcmdInstitution = Nothing
End Sub

Private Sub txtCode_GotFocus()

Call SetSelected

End Sub

Private Sub txtDDDBName_GotFocus()

    Call SetSelected

End Sub

Private Sub txtDDVendorIDNumber_GotFocus()
    
    Call SetSelected

End Sub

Private Sub txtName_GotFocus()

Call SetSelected

End Sub

Private Sub txtDDSendReportTo_GotFocus()

Call SetSelected

End Sub

