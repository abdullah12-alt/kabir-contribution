VERSION 5.00
Object = "{D9D1F94F-AEDB-11D2-9C3C-00105A19BCF2}#1.0#0"; "OAOTitle.ocx"
Begin VB.Form frmRegions 
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
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   2805
      TabIndex        =   1
      Top             =   930
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2805
      TabIndex        =   2
      Top             =   1410
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2805
      TabIndex        =   3
      Top             =   1890
      Width           =   975
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Details"
      Height          =   5415
      Left            =   4095
      TabIndex        =   11
      Top             =   720
      Width           =   5625
      Begin VB.TextBox txtToEmail 
         Height          =   870
         Left            =   2670
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1095
         Width           =   2700
      End
      Begin VB.TextBox txtRegion 
         Enabled         =   0   'False
         Height          =   285
         HelpContextID   =   30037
         Left            =   2670
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   4
         Top             =   675
         WhatsThisHelpID =   30036
         Width           =   2700
      End
      Begin VB.TextBox txtCCEmail 
         Enabled         =   0   'False
         Height          =   945
         HelpContextID   =   30037
         Left            =   2670
         Locked          =   -1  'True
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   2070
         WhatsThisHelpID =   30036
         Width           =   2700
      End
      Begin VB.TextBox txtModifiedBy 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   14
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
         Left            =   3525
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   4965
         Width           =   1935
      End
      Begin VB.Label lblDDDBName 
         Caption         =   "Send Email TO address:"
         Height          =   255
         Left            =   315
         TabIndex        =   19
         Top             =   1140
         Width           =   2055
      End
      Begin VB.Label lblCode 
         Caption         =   "Region:"
         Height          =   210
         Left            =   330
         TabIndex        =   18
         Top             =   705
         Width           =   2250
      End
      Begin VB.Label lblSendReportTo 
         Caption         =   "Send Email CC address:"
         Height          =   210
         Left            =   315
         TabIndex        =   17
         Top             =   2100
         Width           =   2250
      End
      Begin VB.Label lblModifiedBy 
         Caption         =   "Record Modified By:"
         Height          =   255
         Left            =   1995
         TabIndex        =   12
         Top             =   4725
         Width           =   1485
      End
      Begin VB.Label lblModifiedDate 
         Caption         =   "Record Modified On:"
         Height          =   255
         Left            =   1995
         TabIndex        =   13
         Top             =   4965
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   8565
      TabIndex        =   7
      Top             =   6495
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7350
      TabIndex        =   8
      Top             =   6495
      Width           =   1095
   End
   Begin VB.Frame fraCodes 
      Caption         =   "Regions"
      Height          =   5415
      Left            =   360
      TabIndex        =   10
      Top             =   720
      Width           =   2055
      Begin VB.ListBox lstCodes 
         Height          =   4740
         ItemData        =   "frmRegions.frx":0000
         Left            =   225
         List            =   "frmRegions.frx":0002
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6150
      TabIndex        =   9
      Top             =   6495
      Width           =   1095
   End
   Begin OAOTitle.OutlookTitle OutTitle 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   979
      ForeColor       =   16777215
      Picture         =   "frmRegions.frx":0004
      Caption         =   "Regions"
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
Attribute VB_Name = "frmRegions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ********************************************************************************
' * Description:
' * The Regions Screen is used to view, add, edit, and delete Region
'*  codes from the system. This screen serves 4 primary goals:
' *  1) Display all Regions in the system
' *  2) Add a new Region
' *  3) Edit an existing Region
' *  4) Delete an existing Region
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
' *
' ********************************************************************************
' Mod INTERFACES
' Mod EVENTS

' Mod CONSTANTS
Private Const MODULE As String = "Regions"

Private Enum UpdateMode
    VIEW_MODE
    ADD_MODE
    DELETE_MODE
    EDIT_MODE
End Enum
' Mod ENUMS
' Mod TYPES

' Mod DECLARES
Private mcmdRegion As New ADODB.Command
Private mrsRegion As New ADODB.Recordset

' Mod VARIABLES
Private miUpdateMode As UpdateMode
Private msMessage As String
Private msTitle As String
Private Style As VbMsgBoxStyle
Private dblRegionID As Double
Private msLastSelected As String

Private Sub cmdDelete_Click()

On Error GoTo cmdDelete_ClickErr

'********************************************************************************
'* Name: cmdDelete_Click
'* Description:
'* Created: 6/2/2001 3:26 PM
'********************************************************************************
Dim strMessage, strTitle As String
    Style = vbYesNo + vbExclamation + vbDefaultButton2 + vbApplicationModal
    strMessage = "Are you sure you want to delete this region?"
    strTitle = "Confirm Region Deletion"
    If MsgBox(strMessage, Style, strTitle) = vbNo Then
        'disable buttons
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    Else
        miUpdateMode = DELETE_MODE
        'run delete procedure
        Call delRegion
        'update the list
        Call UpdateList
        'disable buttons
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    End If

Xit:
    Exit Sub

cmdDelete_ClickErr:
    ShowUnexpectedError MODULE + "cmdDelete_Click", Err
    Resume Xit



End Sub

Private Sub cmdNew_Click()
    
    ChangeScreenMode (ADD_MODE)

End Sub

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
    msLastSelected = txtRegion
    Hourglass True
    'check to see if this code already exists
    If fnCkCodeExists = False Then
        'this code doesn't exist so check that the fields are filled out properly
        If DataValidation = True Then
            'the fields are filled out so execute procedure to update the record
            Call iuRegion
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
            Call iuRegion
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
    msLastSelected = txtRegion
    Style = vbYesNo + vbQuestion + vbDefaultButton2 + vbApplicationModal
    msMessage = "Are you sure you want to cancel?"
        
    If miUpdateMode = ADD_MODE Then
        If MsgBox(msMessage, Style) = vbNo Then
            'go back
        Else
            'cancel and lose changes
            ChangeScreenMode (VIEW_MODE)
            SelectLastSelected
        End If
    ElseIf miUpdateMode = EDIT_MODE Then
        If MsgBox(msMessage, Style) = vbNo Then
            'go back
        Else
            'cancel and lose changes
            ChangeScreenMode (VIEW_MODE)
            SelectLastSelected
        End If
    Else
        'cancel and lose changes
        Unload Me
    End If

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
    Set mcmdRegion.ActiveConnection = gcnDDS
    With mcmdRegion
    .CommandType = adCmdText
    .CommandText = "SELECT REGION_ID,REGION,EMAIL_RECIPIENTS_TO,EMAIL_RECIPIENTS_CC,LAST_MOD_BY,LAST_MOD_DATETIME FROM DD_REGION ORDER BY REGION"
    Set mrsRegion = .Execute
    End With
    With mrsRegion
    Do Until .EOF
        lstCodes.AddItem !REGION
        .MoveNext
    Loop
    If Not .BOF Then
        .MoveFirst
        Call UpdateFields
    End If
    End With

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

    With mrsRegion
        txtRegion = !REGION
        txtToEmail = !EMAIL_RECIPIENTS_TO
        txtCCEmail = ConvertNull(!EMAIL_RECIPIENTS_CC)
        txtModifiedBy = ConvertNull(!LAST_MOD_BY)
        txtModifiedDate = ConvertNull(!LAST_MOD_DATETIME)
        dblRegionID = !REGION_ID
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

    With mrsRegion
        .MoveFirst
        While .Fields("REGION") <> lstCodes.Text
            .MoveNext
        Wend
    End With
    Call UpdateFields
    cmdEdit.Enabled = True
    cmdNew.Enabled = True
    cmdDelete.Enabled = True
    

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

If mrsRegion.BOF And mrsRegion.EOF Then
    fnCkCodeExists = False
Else
    If miUpdateMode = ADD_MODE Then
    'if trying to insert a new code then see if it already exists, otherwise it doesn't matter
        Dim bExists As Boolean
        With mrsRegion
            bExists = False
            .MoveFirst
            Do Until .EOF
                If .Fields("REGION") <> Trim(txtRegion.Text) Then
                    bExists = False
                    .MoveNext
                Else
                    bExists = True
                    Exit Do
                End If
            Loop
            If bExists = True Then
                Style = vbOKOnly + vbExclamation + vbApplicationModal
                msTitle = "Region Exists"
                msMessage = "This region exists, please enter a different value"
                Hourglass False
                If MsgBox(msMessage, Style, msTitle) = vbOK Then
                    fnCkCodeExists = True
                    'set focus to code entry field
                    txtRegion.SetFocus
                End If
                Hourglass True
            Else
                fnCkCodeExists = False
            End If
        End With
    Else
        fnCkCodeExists = False
    End If
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
    
    If txtRegion.Text = vbNullString Then
        strEmptyFields = " Region"
        txtRegion.SetFocus
    End If
    
    If txtToEmail.Text = vbNullString And strEmptyFields <> vbNullString Then
        strEmptyFields = strEmptyFields & vbCrLf & " Send Email To address"
    Else
        If txtToEmail.Text = vbNullString Then
        strEmptyFields = strEmptyFields & " Send Email To address"
        txtToEmail.SetFocus
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

Private Sub iuRegion()

On Error GoTo iuRegionErr
'********************************************************************************
'* Name: iuRegion
'*
'* Description:
'* Created: 7/9/2001 4:20:26 PM
'********************************************************************************

        Dim cmd As New ADODB.Command
        Dim iX As Integer
        Set cmd.ActiveConnection = gcnDDS
        If gStoredProcs("up_iud_Regions").GetStoredProcCommand(cmd) = True Then
            
            For iX = 1 To cmd.Parameters.Count - 1
                cmd.Parameters(iX) = Null
            Next iX
            If miUpdateMode = EDIT_MODE Then
                cmd.Parameters("region_id") = dblRegionID
            End If
            cmd.Parameters("region") = Trim$(txtRegion.Text)
            cmd.Parameters("email_recipients_to") = Trim$(txtToEmail.Text)
            If txtCCEmail <> "" Then
                cmd.Parameters("email_recipients_cc") = Trim$(txtCCEmail.Text)
            End If
            cmd.Parameters("user_id") = gobjLoginInfo.UserId
            If miUpdateMode = ADD_MODE Then
                cmd.Parameters("update_mode") = "I"
            ElseIf miUpdateMode = EDIT_MODE Then
                cmd.Parameters("update_mode") = "U"
            Else
                Err.Raise 34567, , "Update mode was not set properly"
            End If
            cmd.Execute
            If cmd.Parameters("RETURN_VALUE") <> 0 Then
                GetServerErrorMsg cmd.Parameters("RETURN_VALUE"), "Error occurred adding or updating the Region record."
            End If
        Else
            MsgBox "Error creating the Insert/Update Region Stored Procedure.", vbCritical
            Set cmd = Nothing
            ExitApp
        End If

Xit:
    
    Set cmd = Nothing
    Exit Sub

iuRegionErr:
    ShowUnexpectedError MODULE + "iuRegion", Err
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
    
    Case ADD_MODE
        
        'clear the controls
        txtRegion.Text = vbNullString
        txtToEmail.Text = vbNullString
        txtCCEmail.Text = vbNullString

        'Change background color of mandatory fields
        txtRegion.BackColor = PALE_YELLOW
        txtToEmail.BackColor = PALE_YELLOW

        'enable the controls that are modifiable during an add
        txtRegion.Enabled = True
        txtRegion.Locked = False
        txtToEmail.Enabled = True
        txtToEmail.Locked = False
        txtCCEmail.Enabled = True
        txtCCEmail.Locked = False

        'set focus to first control
        txtRegion.SetFocus
        'enable all buttons available during an add
        cmdOK.Enabled = True
        cmdApply.Enabled = True
        'disable the controls that cannot be changed during an add
        lstCodes.Enabled = False
        'disable all buttons not available during an add
        cmdNew.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False

        miUpdateMode = ADD_MODE

    Case VIEW_MODE
        'restore the background of the controls to white while in view mode
        txtRegion.BackColor = DFLT_WHITE
        txtToEmail.BackColor = DFLT_WHITE
        txtCCEmail.BackColor = DFLT_WHITE
        'enable the controls that are modifiable during view
        'set focus
        'disable the controls that cannot be changed during view
        txtRegion.Enabled = False
        txtRegion.Locked = True
        txtToEmail.Enabled = False
        txtToEmail.Locked = True
        txtCCEmail.Enabled = False
        txtCCEmail.Locked = True
        'disable all buttons not available during view
        cmdNew.Enabled = True
        cmdDelete.Enabled = False
        cmdEdit.Enabled = False
        cmdApply.Enabled = False
        cmdOK.Enabled = False
        Call UpdateList
        miUpdateMode = VIEW_MODE
    
    Case EDIT_MODE
        'change the background of the controls that are mandatory for an add
        txtRegion.BackColor = PALE_YELLOW
        txtToEmail.BackColor = PALE_YELLOW
        'enable the controls that are modifiable during an add
        txtRegion.Enabled = True
        txtRegion.Locked = False
        txtToEmail.Enabled = True
        txtToEmail.Locked = False
        txtCCEmail.Enabled = True
        txtCCEmail.Locked = False
        'set focus to first control
        txtRegion.SetFocus
        'enable all buttons available during an edit
        cmdOK.Enabled = True
        cmdApply.Enabled = True
        'disable the controls that cannot be changed during an add
        lstCodes.Enabled = False
        'disable all buttons not available during an edit
        cmdEdit.Enabled = False
        cmdNew.Enabled = False
        cmdDelete.Enabled = False
        miUpdateMode = EDIT_MODE
    End Select

Xit:
    Exit Sub
    
ChangeScreenModeErr:
    ShowError MODULE + ".ChangeScreenMode", Err
    Resume
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'********************************************************************************
'* Name: Form_Unload
'*
'* Description:
'* Created: 1/27/99 3:44:43 PM
'********************************************************************************
On Error Resume Next
mrsRegion.Close
Set mrsRegion = Nothing
Set mcmdRegion = Nothing
End Sub

Private Sub txtCCEmail_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 59 Then
        KeyAscii = 44
    End If

    If KeyAscii = 44 And Right$(txtCCEmail, 1) = "," Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtRegion_GotFocus()

Call SetSelected

End Sub

Private Sub txtToEmail_GotFocus()

    Call SetSelected

End Sub


Private Sub txtCCEmail_GotFocus()

Call SetSelected

End Sub

Private Sub delRegion()

On Error GoTo delRegionErr

'********************************************************************************
'* Name: delRegion
'* Description:
'* Created: 7/9/2001 2:40:05 PM
'********************************************************************************
    
    'Hourglass True

    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = gcnDDS
    If gStoredProcs("up_iud_Regions").GetStoredProcCommand(cmd) = True Then
        cmd.Parameters("region_id") = dblRegionID
        cmd.Parameters("user_id") = gobjLoginInfo.UserId
        cmd.Parameters("update_mode") = "D"
        cmd.Execute
        If cmd.Parameters("RETURN_VALUE") <> 0 Then
            GetServerErrorMsg cmd.Parameters("RETURN_VALUE"), "Error occurred deleting the Region record."
        End If
    Else
        MsgBox "Error creating the Delete Region Stored Procedure.", vbCritical
        Set cmd = Nothing
        ExitApp
    End If
    Set cmd = Nothing

Xit:
    Hourglass False
    Exit Sub

delRegionErr:
    ShowUnexpectedError MODULE + "delRegion", Err
    Resume Xit

End Sub

Private Sub txtToEmail_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 59 Then
        KeyAscii = 44
    End If

    If KeyAscii = 44 And Right$(txtToEmail, 1) = "," Then
        KeyAscii = 0
    End If

End Sub
