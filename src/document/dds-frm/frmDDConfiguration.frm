VERSION 5.00
Object = "{8CD222DF-7752-11D3-9D1E-00105A19BCF2}#1.0#0"; "OAOTBar.ocx"
Begin VB.Form frmDDConfiguration 
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   10155
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Affinity Data Acquisition Information"
      Height          =   1380
      Left            =   120
      TabIndex        =   34
      Top             =   3975
      Width           =   4095
      Begin VB.TextBox txtLookback 
         Height          =   375
         Left            =   1425
         TabIndex        =   11
         Top             =   795
         Width           =   2550
      End
      Begin VB.TextBox txtDataRefresh 
         Height          =   375
         Left            =   1425
         TabIndex        =   9
         Top             =   300
         Width           =   2535
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Visit Lookback (months):"
         Height          =   495
         Left            =   180
         TabIndex        =   10
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Refresh (minutes):"
         Height          =   375
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   1215
      End
   End
   Begin OAOTitleBar.OutlookTitleBar OutlookTitle1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   979
      ForeColor       =   16777215
      Caption         =   "DDS Configuration"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Patient Account Posting Information"
      Height          =   1380
      Left            =   120
      TabIndex        =   32
      Top             =   2445
      Width           =   4095
      Begin VB.TextBox txtFT1InsuranceCode 
         Height          =   375
         Left            =   1425
         TabIndex        =   5
         Top             =   300
         Width           =   2535
      End
      Begin VB.TextBox txtPATEnteringArea 
         Height          =   375
         Left            =   1425
         TabIndex        =   7
         Top             =   795
         Width           =   2550
      End
      Begin VB.Label Label2 
         Caption         =   "Insurance Code:"
         Height          =   375
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "PAT Code Entering Area:"
         Height          =   495
         Left            =   345
         TabIndex        =   6
         Top             =   780
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "Modify"
      Height          =   375
      Left            =   4560
      TabIndex        =   26
      Top             =   6525
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "FUNB ID's"
      Height          =   1380
      Left            =   120
      TabIndex        =   31
      Top             =   960
      Width           =   4095
      Begin VB.TextBox txtFUNBReceiverID 
         Height          =   375
         Left            =   1410
         TabIndex        =   3
         Top             =   795
         Width           =   2565
      End
      Begin VB.TextBox txtFUNBSenderID 
         Height          =   375
         Left            =   1410
         TabIndex        =   1
         Top             =   315
         Width           =   2565
      End
      Begin VB.Label lblFUNBReceiverID 
         Alignment       =   1  'Right Justify
         Caption         =   "Receiver ID:"
         Height          =   375
         Left            =   105
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblFUNBSenderID 
         Alignment       =   1  'Right Justify
         Caption         =   "Sender ID:"
         Height          =   375
         Left            =   105
         TabIndex        =   0
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.Frame fraSendEmail 
      Caption         =   "State Treasurer Information"
      Height          =   4395
      Left            =   4350
      TabIndex        =   30
      Top             =   960
      Width           =   5340
      Begin VB.TextBox txtPFBatchName 
         Height          =   285
         Left            =   1725
         MaxLength       =   15
         TabIndex        =   17
         Top             =   1125
         Width           =   3450
      End
      Begin VB.TextBox txtPABatchName 
         Height          =   285
         Left            =   1725
         MaxLength       =   15
         TabIndex        =   15
         Top             =   765
         Width           =   3450
      End
      Begin VB.TextBox txtPAVENDORIDNUM 
         Height          =   285
         Left            =   1725
         MaxLength       =   15
         TabIndex        =   13
         Top             =   405
         Width           =   3450
      End
      Begin VB.TextBox txtNoteText 
         Height          =   1620
         Left            =   1725
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   2595
         Width           =   3450
      End
      Begin VB.TextBox txtsubject 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1725
         TabIndex        =   23
         Top             =   2205
         Width           =   3450
      End
      Begin VB.TextBox txtcc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1725
         TabIndex        =   21
         Top             =   1845
         Width           =   3450
      End
      Begin VB.TextBox txtTo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1725
         TabIndex        =   19
         Top             =   1485
         Width           =   3450
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "PF Batch File Name:"
         Height          =   240
         Left            =   195
         TabIndex        =   16
         Top             =   1155
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "PA Batch File Name:"
         Height          =   240
         Left            =   135
         TabIndex        =   14
         Top             =   795
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text:"
         Height          =   195
         Left            =   195
         TabIndex        =   24
         Top             =   2625
         Width           =   1455
      End
      Begin VB.Label lblVID 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendor ID Number:"
         Height          =   240
         Left            =   195
         TabIndex        =   12
         Top             =   450
         Width           =   1455
      End
      Begin VB.Label lblSubject 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subj&ect:"
         Height          =   195
         Left            =   195
         TabIndex        =   22
         Top             =   2220
         Width           =   1455
      End
      Begin VB.Label lblCc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cc:"
         Height          =   195
         Left            =   195
         TabIndex        =   20
         Top             =   1875
         Width           =   1455
      End
      Begin VB.Label lblTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&To:"
         Height          =   195
         Left            =   195
         TabIndex        =   18
         Top             =   1515
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5880
      TabIndex        =   27
      Top             =   6525
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   28
      Top             =   6525
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   8520
      TabIndex        =   29
      Top             =   6525
      Width           =   1095
   End
End
Attribute VB_Name = "frmDDConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************
' * Form Name:frmDDConfiguration
' * Form File Name: DDConfiguration.frm
' * Start Date: 6/19/1999
' * End Date:   7/26/1999
' * Description:
' * --------------------------------
' * The DD Configuration Screen is divided into
'   several frames as follows:
'   � Patient Account Vendor ID
'   � FUNB ID's,
'   � Patient Account Posting Info., and
'   � Email Options
'
'
'
' Mod CONSTANTS
Private Const MODULE As String = "DDS Configurations"
Private Const strInsert As String = "I"
Private Const strUpdate As String = "U"

Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private cmdVendorID As New ADODB.Command


' Mod VARIABLES
Private iEditMode As ScreenMode
Private strUpdateMode As String
Private strMessage As String
Private strTitle As String
Private Style As VbMsgBoxStyle
Dim dbVendorIDNumberID As Double

Private Sub cmdApply_Click()
On Error GoTo cmdApply_ClickErr

'********************************************************************************
'* Name: cmdApply_Click
'* Description:
'* Created: 6/8/99 4:06 PM
'********************************************************************************

    'check to see if this code already exists
    If fnCkCodeExists = False Then
        'this code doesn't exist so check that the fields are filled out properly
        If DataValidation = True Then
            'the fields are filled out so execute procedure to update the record
            Call UpdateDDConfig
             
            'reset the screen to view mode
            ChangeScreenMode (VIEW_MODE)
        Else
        'the fields are not filled out properly so don't execute the procedure
        End If
    Else
        'this code already exists so don't do anything else
    End If

Xit:
    Exit Sub

cmdApply_ClickErr:
    ShowUnexpectedError MODULE + "cmdApply_Click", Err
    Resume Xit
End Sub



Private Sub cmdCancel_Click()
'********************************************************************************
'* Name: cmdCancel_Click
'*
'* Description:
'* Created: 6/8/99 5:30 PM
'********************************************************************************
'Unload Me
    On Error GoTo cmdCancel_ClickErr
    Style = vbYesNo + vbQuestion + vbDefaultButton2 + vbApplicationModal
    strMessage = "Are you sure you want to cancel?"
        
    Select Case strUpdateMode
         'No Insert
        'Case strInsert
          '  If MsgBox(strMessage, Style) = vbNo Then
          '        go back
          '  Else
          '         cancel and lose changes
          '         ChangeScreenMode (VIEW_MODE)
          '  End If

        Case strUpdate
            If MsgBox(strMessage, Style) = vbNo Then
                'go back
            Else
                'cancel and lose changes
                ChangeScreenMode (VIEW_MODE)
            End If
            
        Case vbNullString
'             Call CloseConnection
            'cancel and lose changes
            Unload Me
    End Select

Xit:
   Exit Sub

cmdCancel_ClickErr:
'   ShowUnexpectedError MODULE + "cmdCancel_Click", Err
   Resume Xit
End Sub


Private Sub cmdModify_Click()
'********************************************************************************
'* Name: cmdModify_Click
'* Description:To change the current Options the user would click on the Modify button.
'* Created: 6/21/99 3:26 PM
'********************************************************************************
strUpdateMode = strUpdate
ChangeScreenMode (EDIT_MODE)
End Sub

Private Sub cmdOK_Click()
On Error GoTo cmdOK_ClickErr

'********************************************************************************
'* Name: cmdOk_Click
'* Description:If the OK button is clicked and required fields pass data validation,
'* update the database with changes. Refresh the screen with changes.
'* Created: 6/23/99 5:21 PM
'********************************************************************************
    
    'check to see if this code already exists
    If fnCkCodeExists = False Then
       'this code doesn't exist so check that the fields are filled out properly
        If DataValidation = True Then
            'the fields are filled out so execute procedure to update the record
            Call UpdateDDConfig
            'now close the form
            Unload Me
        Else
        'the fields are not filled out properly so don't execute the procedure
        End If
    Else
        'this code already exists so don't do anything else
    End If

Xit:
    Exit Sub

cmdOK_ClickErr:
    'ShowUnexpectedError MODULE + "cmdOK_Click", Err
    Resume Xit

End Sub

Private Sub Form_Activate()

    fMainForm.SetMainToolbar True
    
End Sub

Private Sub Form_Deactivate()
    
    fMainForm.SetMainToolbar False

End Sub

Private Sub Form_Load()
'********************************************************************************
'* Name: Form_Load
'* Description:
'* Created: 6/19/99 6:10 PM
'********************************************************************************
Set OutlookTitle1.Picture = fMainForm.imlToolbarIcons.ListImages("DDS Configuration").Picture
'fraDetails.Visible = False
cmdOK.Enabled = False
cmdApply.Enabled = False

'Call UpdateList Procedure
Call UpdateList
ChangeScreenMode (VIEW_MODE)
End Sub
'********************************************************************************
'* Name: UpdateList()
'* Description:
'* Created: 6/25/99 4:06 PM
'********************************************************************************
Public Sub UpdateList()
On Error GoTo UpdateListError
Dim sSql As String
Dim rs As New ADODB.Recordset
sSql = "SELECT * FROM DD_CONFIG_INFO"
sSql = sSql & " ORDER BY PA_VENDOR_ID_NUM"
rs.Open sSql, gcnDDS, adOpenForwardOnly
With rs
     .MoveFirst
    txtPAVENDORIDNUM.Text = ConvertNull(!PA_VENDOR_ID_NUM)
    txtPAVENDORIDNUM = !PA_VENDOR_ID_NUM
    txtFUNBSenderID = !SENDER_ID
    txtFUNBReceiverID = !RECEIVER_ID
    txtFT1InsuranceCode = !FT1_INSURANCE_CODE
    txtPATEnteringArea = !PATCODE_ENTERING_AREA
    txtTo = !ST_TREAS_EMAIL_TO_ADDR
    txtsubject = !ST_TREAS_EMAIL_SUBJ
    txtcc = ConvertNull(!ST_TREAS_EMAIL_CC_ADDR)
    txtNoteText = ConvertNull(!ST_TREAS_EMAIL_TEXT)
    dbVendorIDNumberID = !CONFIG_ID
    txtDataRefresh = ConvertNull(!DATA_REFRESH_RATE)
    txtLookback = ConvertNull(!DATA_LOOKBACK)
    txtPABatchName = ConvertNull(!PA_BATCH_NAME)
    txtPFBatchName = ConvertNull(!PF_BATCH_NAME)
    End With
    
Xit:
    Hourglass False
    Exit Sub

UpdateListError:
    ShowUnexpectedError MODULE + "UpdateList", Err
    Resume Xit
End Sub



Private Sub Form_Unload(Cancel As Integer)
'********************************************************************************
'* Name: Form_Unload
'* Description:
'* Created: 5/11/1999 10:11:15 AM
'********************************************************************************
On Error Resume Next
 Call CloseConnection
End Sub


'
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
'* Created: 6/08/99 4:23
'********************************************************************************
    'Change the mode for the screen
    
    Hourglass True
    iEditMode = iMode
    
    Select Case iMode
    
    Case VIEW_MODE
        'restore the background of the controls to white while in view mode
        txtPAVENDORIDNUM.BackColor = DFLT_WHITE
        txtPAVENDORIDNUM.Enabled = False
        txtPAVENDORIDNUM.Locked = True
        Call UpdateList
        Call FreezeTextControls
        cmdApply.Enabled = False
        cmdOK.Enabled = False
        strUpdateMode = vbNullString
        cmdModify.Enabled = True

'    Case ADD_MODE
'    We do not add any thing for this Screen, so Add mode is enable
'
'
    Case EDIT_MODE
'       change the background of the controls that are mandatory for an add
'
        Call UnFreezeTextControls
        cmdOK.Enabled = True
        cmdApply.Enabled = True
        cmdModify.Enabled = False
    End Select
    

Xit:
    Hourglass False
    Exit Sub

ChangeScreenModeErr:
    ShowUnexpectedError MODULE + "ChangeScreenMode", Err
    Resume Xit


End Sub
Private Sub cmdEdit_Click()
'********************************************************************************
'* Name: cmdEdit_Click
'* Description:
'* Created: 1/28/:99 3:55:10 PM
'********************************************************************************
    
    strUpdateMode = strUpdate
    ChangeScreenMode (EDIT_MODE)
End Sub



Public Sub CloseConnection()
'********************************************************************************
'* Name: CloseConnection
'* Description:
'* Created: 6/21/:99 3:55:10 PM
'********************************************************************************
On Error Resume Next
rs.Close
conn.Close
'Clear the recordset and the connection
Set rs = Nothing
Set conn = Nothing
End Sub


Public Sub FreezeTextControls()
'********************************************************************************
'* Name: FreezeEmailOptions
'* Description: Update the controls for the Email Option Frame
'* Created: 7/11/:99 5:55:13 PM
'********************************************************************************
    
    txtFUNBReceiverID.BackColor = DFLT_WHITE
    txtFUNBSenderID.BackColor = DFLT_WHITE
    txtFUNBSenderID.Enabled = False
    txtFUNBSenderID.Locked = True
    txtFUNBReceiverID.Enabled = False
    txtFUNBReceiverID.Locked = True
    txtFT1InsuranceCode.BackColor = DFLT_WHITE
    txtPATEnteringArea.BackColor = DFLT_WHITE
    txtFT1InsuranceCode.Enabled = False
    txtFT1InsuranceCode.Locked = True
    txtPATEnteringArea.Enabled = False
    txtPATEnteringArea.Locked = True
    txtTo.BackColor = DFLT_WHITE
    txtcc.BackColor = DFLT_WHITE
    txtsubject.BackColor = DFLT_WHITE
    txtNoteText.BackColor = DFLT_WHITE
    txtTo.Enabled = False
    txtTo.Locked = True
    txtcc.Enabled = False
    txtcc.Locked = True
    txtsubject.Enabled = False
    txtsubject.Locked = True
    txtNoteText.Enabled = False
    txtNoteText.Locked = True

    txtDataRefresh.BackColor = DFLT_WHITE
    txtLookback.BackColor = DFLT_WHITE
    txtPABatchName.BackColor = DFLT_WHITE
    txtPFBatchName.BackColor = DFLT_WHITE
    txtDataRefresh.Enabled = False
    txtDataRefresh.Locked = True
    txtLookback.Enabled = False
    txtLookback.Locked = True
    txtPABatchName.Enabled = False
    txtPABatchName.Locked = True
    txtPFBatchName.Enabled = False
    txtPFBatchName.Locked = True


End Sub

Public Sub UnFreezeTextControls()
'********************************************************************************
'* Name: UnFreezeEmailOptions()
'* Description: Update the controls for the Email Option Frame
'* Created: 7/12/:99 1:55:13 PM
'********************************************************************************
        txtPAVENDORIDNUM.BackColor = PALE_YELLOW
        txtPAVENDORIDNUM.Enabled = True
        txtPAVENDORIDNUM.Locked = False

        txtFUNBReceiverID.BackColor = PALE_YELLOW
        txtFUNBSenderID.BackColor = PALE_YELLOW
        txtFUNBSenderID.Enabled = True
        txtFUNBSenderID.Locked = False
        txtFUNBReceiverID.Enabled = True
        txtFUNBReceiverID.Locked = False

        txtFT1InsuranceCode.BackColor = PALE_YELLOW
        txtPATEnteringArea.BackColor = PALE_YELLOW
        txtFT1InsuranceCode.Enabled = True
        txtFT1InsuranceCode.Locked = False
        txtPATEnteringArea.Enabled = True
        txtPATEnteringArea.Locked = False

        txtTo.BackColor = PALE_YELLOW
        txtcc.BackColor = PALE_YELLOW
        
        txtsubject.BackColor = PALE_YELLOW
        txtNoteText.BackColor = PALE_YELLOW
        'lblTOMultiple.BackColor = DFLT_WHITE
        txtTo.Enabled = True
        txtTo.Locked = False
        txtcc.Enabled = True
        txtcc.Locked = False
        txtsubject.Enabled = True
        txtsubject.Locked = False
'        txtDate.Enabled = True
'        txtDate.Locked = False
        txtNoteText.Enabled = True
        txtNoteText.Locked = False

        txtDataRefresh.BackColor = PALE_YELLOW
        txtLookback.BackColor = PALE_YELLOW
        txtPABatchName.BackColor = PALE_YELLOW
        txtPFBatchName.BackColor = PALE_YELLOW
        txtDataRefresh.Enabled = True
        txtDataRefresh.Locked = False
        txtLookback.Enabled = True
        txtLookback.Locked = False
        txtPABatchName.Enabled = True
        txtPABatchName.Locked = False
        txtPFBatchName.Enabled = True
        txtPFBatchName.Locked = False


End Sub

Private Function fnCkCodeExists() As Boolean

On Error GoTo fnCkCodeExistsErr

'********************************************************************************
'* Name: fnCkCodeExists
'* Description:
'* Created: 6/02/99 3:40
'********************************************************************************

    If strUpdateMode = strInsert Then
    'if trying to insert a new code then see if it already exists, otherwise it doesn't matter
        Dim bExists As Boolean
        Hourglass True
        With rs
            bExists = False
            .MoveFirst
            Do Until .EOF
                If .Fields("INCOME_SOURCE_TYPE_CODE") <> Trim(UCase(txtPAVENDORIDNUM.Text)) Then
                    bExists = False
                    .MoveNext
                Else
                    bExists = True
                    Exit Do
                End If
            Loop
            If bExists = True Then
                Style = vbOKOnly + vbExclamation + vbApplicationModal
                strTitle = "Code Exists"
                strMessage = "This code exists, please enter a different value"
                Hourglass False
                If MsgBox(strMessage, Style, strTitle) = vbOK Then
                    fnCkCodeExists = True
                    'set focus to code entry field
                    txtPAVENDORIDNUM.SetFocus
                End If
            Else
                fnCkCodeExists = False
            End If
        End With
    Else
        fnCkCodeExists = False
    End If

Xit:
    Hourglass False
    Exit Function

fnCkCodeExistsErr:
    ShowUnexpectedError MODULE + "fnCkCodeExists", Err
    Resume Xit


End Function
Private Sub UpdateDDConfig()

On Error GoTo UpdateDDConfigErr

'********************************************************************************
'* Name: iuVendorIDNumber
'* Description: Calling the Procedure to Update the DD Configuration Screen Options
'* the stored procedure callled "up_iu_Config"
'* Created: 6/02/99 3:39
'********************************************************************************

     Hourglass True
Dim ix As Integer
    Set cmdVendorID.ActiveConnection = gcnDDS

    If gStoredProcs("up_iu_Config").GetStoredProcCommand(cmdVendorID) = True Then
        With cmdVendorID
             For ix = 0 To cmdVendorID.Parameters.Count - 1
                cmdVendorID.Parameters(ix) = Null
             Next ix

            .Parameters("config_id") = dbVendorIDNumberID
            .Parameters("pa_vendor_id_num") = Trim(UCase(txtPAVENDORIDNUM.Text))
            .Parameters("sender_id") = Trim(UCase(txtFUNBSenderID.Text))
            .Parameters("receiver_id") = Trim(UCase(txtFUNBReceiverID.Text))
            .Parameters("st_treas_email_to_addr") = Trim(LCase(txtTo.Text))
            .Parameters("ft1_insurance_code") = Trim(UCase(txtFT1InsuranceCode.Text))
            .Parameters("patcode_entering_area") = Trim(UCase(txtPATEnteringArea.Text))
            .Parameters("st_treas_email_cc_addr") = Trim(LCase(txtcc.Text))
            .Parameters("st_treas_email_text") = Trim(UCase(txtNoteText.Text))
            .Parameters("st_treas_email_subj") = Trim(UCase(txtsubject.Text))
            .Parameters("data_refresh_rate") = txtDataRefresh
            .Parameters("data_lookback") = txtLookback
            .Parameters("pa_batch_name") = Trim$(UCase(txtPABatchName.Text))
            .Parameters("pf_batch_name") = Trim$(UCase(txtPFBatchName.Text))
            .Parameters("called_from_another_proc") = "N"
            .Parameters("update_mode") = strUpdateMode
            .Execute
        If .Parameters("RETURN_VALUE") <> 0 Then
            GetServerErrorMsg .Parameters("RETURN_VALUE"), "Error occurred adding or updating the State record."
        End If
        End With
    Else
        MsgBox "Error creating the DD Configuration Stored Procedure.", vbCritical
        Set cmdVendorID = Nothing
        ExitApp
    End If
    Set cmdVendorID = Nothing

Xit:
    Hourglass False
    Exit Sub
    
UpdateDDConfigErr:
    ShowUnexpectedError MODULE + "VendorIDNumber", Err
    Resume Xit

End Sub
Private Function DataValidation() As Boolean

On Error GoTo DataValidationErr

'********************************************************************************
'* Name: DataValidation
'* Description: To validate the Data for NOT NULL Fields in Sybase
'* Created: 7/20/99 11:24 am
'********************************************************************************
    Dim strEmptyFields As String
    Style = vbOKOnly + vbExclamation + vbApplicationModal
    strTitle = "Invalid Data"
    
    ValidateEmptyString "Code", txtPAVENDORIDNUM, strEmptyFields
    ValidateEmptyString "FUNB Sender ID", txtFUNBSenderID, strEmptyFields
    ValidateEmptyString "FUNB Receiver ID", txtFUNBReceiverID, strEmptyFields
    ValidateEmptyString "Insurance Code", txtFT1InsuranceCode, strEmptyFields
    ValidateEmptyString "PAT Entering Area", txtPATEnteringArea, strEmptyFields
    ValidateEmptyString "Email address to", txtTo, strEmptyFields
    ValidateEmptyString "Email Subject", txtsubject, strEmptyFields
    If ValidateEmptyString("Data Refresh", txtDataRefresh, strEmptyFields) = True Then
        'Check to make sure value is numeric
        ValidateValue "Data Refresh", txtDataRefresh, strEmptyFields, 1000000
    End If
    
    If ValidateEmptyString("Visit Lookback", txtLookback, strEmptyFields) = True Then
        'Check to make sure value is numeric
        ValidateValue "Visit Lookback", txtLookback, strEmptyFields, 1000000
    End If
    
    ValidateEmptyString "PA Batch File Name", txtPABatchName, strEmptyFields
    ValidateEmptyString "PF Batch File Name", txtPFBatchName, strEmptyFields
    
    
    txtTo = Trim$(txtTo)
    If Right$(txtTo, 1) = "," Then
        txtTo = Left$(txtTo, Len(txtTo) - 1)
    End If
       
    txtcc = Trim$(txtcc)
    If Right$(txtcc, 1) = "," Then
        txtcc = Left$(txtcc, Len(txtcc) - 1)
    End If
       
    
    If strEmptyFields <> vbNullString Then
        strMessage = strEmptyFields & vbCrLf
        If MsgBox(strMessage, Style, strTitle) = vbOK Then
            DataValidation = False
        End If
    Else
        DataValidation = True
    End If
    strEmptyFields = ""

Xit:
    Exit Function

DataValidationErr:
    ShowUnexpectedError MODULE + "DataValidation", Err
    Resume Xit


End Function

Private Function ValidateEmptyString(ByVal sDescription, ByRef oCtl As Control, ByRef sMsg)
On Error GoTo ValidateEmptyStringErr
    
    If Trim$(oCtl.Text) <> vbNullString Then
        ValidateEmptyString = True
    Else
        ValidateEmptyString = False
        If sMsg = vbNullString Then
            sMsg = sMsg & sDescription & " is empty."
            oCtl.SetFocus
        Else
            sMsg = sMsg & vbCrLf & sDescription & " is empty."
        End If
    End If

Exit Function

ValidateEmptyStringErr:
    sMsg = sMsg & "Error validating string"
    ValidateEmptyString = False
    

End Function

Private Function ValidateValue(ByVal sDescription, ByRef oCtl As Control, ByRef sMsg, Optional ByVal dMaxValue As Double)
On Error GoTo ValidateValueErr
    
    If IsNumeric(oCtl.Text) Then
        If IsNumeric(dMaxValue) Then
            If CDbl(oCtl.Text) < dMaxValue Then
                ValidateValue = True
            Else
                ValidateValue = False
                If sMsg = vbNullString Then
                    sMsg = sMsg & sDescription & " is greater than or equal to maximum value."
                    oCtl.SetFocus
                Else
                    sMsg = sMsg & vbCrLf & sDescription & " is greater than or equal to maximum value."
                End If
            End If
        Else
            ValidateValue = True
        End If
    Else
        ValidateValue = False
        If sMsg = vbNullString Then
            sMsg = sMsg & sDescription & " is not numeric."
            oCtl.SetFocus
        Else
            sMsg = sMsg & vbCrLf & sDescription & " is not numeric."
        End If
    End If

Exit Function

ValidateValueErr:

    sMsg = sMsg & "Error determining value" & vbCrLf
    ValidateValue = False
    

End Function


Private Sub OutlookTitle1_IconClick()
    If cmdCancel.Enabled = True Then
        Unload Me
    End If

End Sub

Private Sub txtcc_GotFocus()
Call SetSelected
End Sub

Private Sub txtcc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 59 Then
        KeyAscii = 44
    End If

    If KeyAscii = 44 And Right$(txtcc, 1) = "," Then
        KeyAscii = 0
    End If

End Sub




Private Sub txtDataRefresh_GotFocus()
Call SetSelected
End Sub

Private Sub txtFT1InsuranceCode_GotFocus()
Call SetSelected
End Sub

Private Sub txtFUNBReceiverID_GotFocus()
Call SetSelected
End Sub


Private Sub txtFUNBSenderID_GotFocus()
Call SetSelected
End Sub



Private Sub txtLookback_GotFocus()
Call SetSelected
End Sub

Private Sub txtNoteText_GotFocus()
Call SetSelected
End Sub

Private Sub txtPABatchName_GotFocus()
Call SetSelected
End Sub

Private Sub txtPATEnteringArea_GotFocus()
Call SetSelected
End Sub

Private Sub txtPAVENDORIDNUM_GotFocus()
'********************************************************************************
'* Name: txtPAVENDORIDNUM
'* Description: To focus on the PA Vendor ID Number
'* Created: 6/29/99 6:04:
'********************************************************************************
Call SetSelected

End Sub



Private Sub txtPFBatchName_GotFocus()
Call SetSelected
End Sub



Private Sub txtsubject_GotFocus()
Call SetSelected
End Sub

Private Sub txtTo_GotFocus()
Call SetSelected
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 59 Then
        KeyAscii = 44
    End If

    If KeyAscii = 44 And Right$(txtTo, 1) = "," Then
        KeyAscii = 0
    End If

End Sub

