VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{8CD222DF-7752-11D3-9D1E-00105A19BCF2}#1.0#0"; "OAOTBar.ocx"
Begin VB.Form frmDDSUsers 
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7515
   ScaleWidth      =   10455
   WindowState     =   2  'Maximized
   Begin OAOTitleBar.OutlookTitleBar OutlookTitle1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   979
      ForeColor       =   16777215
      Caption         =   "Security-DDS Users"
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   8640
      TabIndex        =   11
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   9615
      Begin VB.Frame fraUserDetails 
         Caption         =   "Details"
         Height          =   3165
         Left            =   3480
         TabIndex        =   15
         Top             =   240
         Width           =   5880
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox txtCreatedBy 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1575
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1575
         End
         Begin VB.TextBox txtModifiedBy 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   2400
            Width           =   1575
         End
         Begin VB.TextBox txtPasswordConfirm 
            Height          =   285
            HelpContextID   =   30037
            IMEMode         =   3  'DISABLE
            Left            =   1530
            MaxLength       =   50
            PasswordChar    =   "*"
            TabIndex        =   9
            Top             =   1185
            WhatsThisHelpID =   30036
            Width           =   2235
         End
         Begin VB.CheckBox chkDisabled 
            Caption         =   "Account Disabled"
            Height          =   255
            Left            =   4020
            TabIndex        =   10
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtFirstName 
            Height          =   285
            HelpContextID   =   30037
            Left            =   3780
            MaxLength       =   50
            TabIndex        =   7
            Top             =   540
            WhatsThisHelpID =   30036
            Width           =   1905
         End
         Begin VB.TextBox txtLastName 
            Height          =   285
            HelpContextID   =   30037
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   6
            Top             =   555
            WhatsThisHelpID =   30036
            Width           =   2235
         End
         Begin VB.TextBox txtPassword 
            Height          =   285
            HelpContextID   =   30037
            IMEMode         =   3  'DISABLE
            Left            =   1530
            MaxLength       =   50
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   870
            WhatsThisHelpID =   30036
            Width           =   2235
         End
         Begin MSMask.MaskEdBox edbUserID 
            Height          =   285
            Left            =   1530
            TabIndex        =   5
            Top             =   240
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "?AAAAAAAAA"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Caption         =   "Record Modified By:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   2400
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "Record Modified On:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   2640
            Width           =   1485
         End
         Begin VB.Label Label3 
            Caption         =   "Record Created By:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1800
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Record Created On:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label lblPassword 
            Caption         =   "Password:"
            Height          =   285
            Left            =   150
            TabIndex        =   19
            Top             =   900
            Width           =   1335
         End
         Begin VB.Label lblPasswordConfirm 
            Caption         =   "Confirm Password:"
            Height          =   285
            Left            =   150
            TabIndex        =   18
            Top             =   1230
            Width           =   1335
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblUserID 
            Caption         =   "Login Name:"
            Height          =   285
            Left            =   150
            TabIndex        =   17
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lblUserName 
            Caption         =   "Last, First Name:"
            Height          =   285
            Left            =   150
            TabIndex        =   16
            Top             =   585
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   780
         Width           =   975
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   360
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame fraUsers 
         Caption         =   "User Accounts"
         Height          =   4380
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2055
         Begin VB.ListBox lstUsers 
            Height          =   3960
            ItemData        =   "frmDDSUsers.frx":0000
            Left            =   120
            List            =   "frmDDSUsers.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.Line lin1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      X1              =   120
      X2              =   9720
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   9720
      Y1              =   5760
      Y2              =   5760
   End
End
Attribute VB_Name = "frmDDSUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msLastSelected As String

Private Enum SybaseUserPrivileges
    NON_USER
    IS_USER_ON_SYSTEM
    IS_USER_ON_DDS_DATABASE
    IS_USER_ON_DDS_DATABASE_AND_A_USER
End Enum
'********************************************************************************
' * Form Name:frmDDSUSERS
' * Form File Name: frmDDSUSERS.frm
' * Start Date: 6/29/1999
' * End Date:   7/26/1999
' * Description:
' * --------------------------------
' * The USER SECURITY Screen allows the user to View, Add and Modify users accounts
'
'
' Mod CONSTANTS
Private Const strActive As String = "A"
Private Const strInactive As String = "I"
Private Const strInsert As String = "I"
Private Const strUpdate As String = "U"
Private Const strSSOUser As String = "dds"
Private Const strSSOPassword As String = "presented"
Private Const MODULE As String = "Security - "
Private Const EXTENDSTR As String = "P#Ssa(fC"
' Mod ENUMS
' Mod TYPES
' Mod DECLARES
Private cmdSecurity As New ADODB.Command
'Private rsRole As New ADODB.Recordset
Private rsUser As New ADODB.Recordset
' Mod VARIABLES
Private iEditMode As ScreenMode
Private strUpdateMode As String
Private strMessage As String
Private strTitle As String
Private Style As VbMsgBoxStyle
'Private dblRoleID As Double
Private strStatus As String
Private strNewPassword As String
Private Sub Form_Activate()

    'Set the main toolbar to not see the institution dropdown or report icon
    fMainForm.SetMainToolbar True

End Sub

Private Sub Form_Deactivate()
    fMainForm.SetMainToolbar False
End Sub

Private Sub Form_Load()

On Error GoTo Form_LoadErr

'********************************************************************************
'* Name: Form_Load
'*
'* Description:
'* Parameters:
'* Created: 6/3/99 2:59:07 PM
'********************************************************************************
    Hourglass True
    '*******************************************
    'Populate the role tab
    '*******************************************
'        sstabSecurity.Tab = 1
'Load an image to the Outlook title
Set OutlookTitle1.Picture = fMainForm.imlToolbarIcons.ListImages("Security").Picture

        Call UpdateList
    '**********************************************************************
    'Populate the user tab last so it will always display first at startup
    '**********************************************************************
'        sstabSecurity.Tab = 0
        ChangeScreenMode (VIEW_MODE)

Xit:
    Hourglass False
    Exit Sub

Form_LoadErr:
    ShowUnexpectedError MODULE + "Form_Load", Err
    Resume Xit

End Sub

Private Sub SelectLastSelected()
'*************************************************
'* Selects the last entered
'************************************************
Dim ix As Long
    For ix = 0 To lstUsers.ListCount - 1
        If lstUsers.List(ix) = msLastSelected Then
            lstUsers.Selected(ix) = True
            lstUsers.SetFocus
            Exit For
        End If
    Next ix
    
End Sub

Private Sub cmdNew_Click()
'********************************************************************************
'* Name: cmdNew_Click
'* Description:
'* Created: 6/2/99 3:26 PM
'********************************************************************************
    strUpdateMode = strInsert
    Hourglass True
    ChangeScreenMode (ADD_MODE)
    Hourglass False
End Sub

Private Sub cmdEdit_Click()
'********************************************************************************
'* Name: cmdEdit_Click
'* Description:
'* Created: 6/10/99 3:26 PM
'********************************************************************************
    strUpdateMode = strUpdate
    Hourglass True
    ChangeScreenMode (EDIT_MODE)
    Hourglass False
End Sub

Private Sub cmdDelete_Click()

On Error GoTo cmdDelete_ClickErr

'********************************************************************************
'* Name: cmdDelete_Click
'*
'* Description:
'* Parameters:
'* Created: 6/03/99 3:38 PM
'********************************************************************************
    Style = vbYesNo + vbExclamation + vbDefaultButton2 + vbApplicationModal
    strMessage = "Are you sure you want to delete this user?"
    strTitle = "Confirm Deletion"
    If MsgBox(strMessage, Style, strTitle) = vbNo Then
        'disable buttons
    Else
        'run delete procedure
        Call delUser
        'update the list
        Hourglass True
        ChangeScreenMode (VIEW_MODE)
        Call UpdateList
    End If

Xit:
    Hourglass False
    Exit Sub

cmdDelete_ClickErr:
    ShowUnexpectedError MODULE + "cmdDelete_Click", Err
    Resume Xit

End Sub

Private Sub cmdApply_Click()

On Error GoTo cmdApply_ClickErr

'********************************************************************************
'* Name: cmdApply_Click
'*
'* Description:
'* Parameters:
'* Created: 6/3/99 3:39
    msLastSelected = edbUserID

    Hourglass True
    If DataValidation = True Then
    'the fields are filled out so check to see if this login exists in Sybase
        '3/18/2009 - AS - Will no longer check priviliges
        'If CheckUserPrivileges = False Then
        'this sybase login doesn't exist so check to see if the user id exists in the PFS database
        If fnCkUserExists = False Then
        'this user id doesn't exist, so check to see if the password entry is ok
            If fnPasswordValidation = True Then
            'the password entry was ok, so check to see that the user has access to at least one institution
                'If fnCkInstitutionRights = True Then
                'access to at least one institution has been granted so execute the procedure to update the record
                    Call iuUser
                    'reset the screen to view mode
                    Call UpdateList
                    ChangeScreenMode (VIEW_MODE)
                    SelectLastSelected
                'Else
                'user was not assigned any rights to an institution so don't do anything else
                'End If
            Else
            'there was a problem with the password entry so don't do anything else
            End If
        Else
        'this user id already exists so don't do anything else
        End If
'        Else
'        'this login exists in Sybase so don't do anything else
'        End If
    Else
    'the fields are not filled out properly so don't do anything else
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
'* Parameters:
'* Created: 6/3/99 3:36 PM
'********************************************************************************
    
    Hourglass True
    If DataValidation = True Then
    'the fields are filled out so check to see if this login exists in Sybase
        If CheckUserPrivileges = False Then
        'this sybase login doesn't exist so check to see if the user id exists in the PFS database
            If fnCkUserExists = False Then
            'this user id doesn't exist, so check to see if the password entry is ok
                If fnPasswordValidation = True Then
                'the password entry was ok, so check to see that the user has access to at least one institution
                    'If fnCkInstitutionRights = True Then
                    'access to at least one institution has been granted so execute the procedure to update the record
                        Call iuUser
                        'now close the form
                        Unload Me
                    'Else
                    'user was not assigned any rights to an institution so don't do anything else
                   ' End If
                Else
                'there was a problem with the password entry so don't do anything else
                End If
            Else
            'this user id already exists so don't do anything else
            End If
        Else
        'this login exists in Sybase so don't do anything else
        End If
    Else
    'the fields are not filled out properly so don't do anything else
    End If

Xit:
    Hourglass False
    Exit Sub

cmdOK_ClickErr:
    ShowUnexpectedError MODULE + "cmdOK_Click", Err
    Resume Xit

End Sub

Private Sub cmdCancel_Click()

On Error GoTo cmdCancel_ClickErr

'********************************************************************************
'* Name: cmdCancel_Click
'*
'* Description:
'* Parameters:
'* Created:
'********************************************************************************
    msLastSelected = edbUserID
    
    Style = vbYesNo + vbQuestion + vbDefaultButton2 + vbApplicationModal
    strMessage = "Are you sure you want to cancel?"
    
    Select Case strUpdateMode
        Case strInsert
'            sstabSecurity.Tab = 0
            If MsgBox(strMessage, Style) = vbNo Then
                'go back
            Else
                'cancel and lose changes
                Hourglass True
                ChangeScreenMode (VIEW_MODE)
                SelectLastSelected
            End If

        Case strUpdate
'            sstabSecurity.Tab = 0
            If MsgBox(strMessage, Style) = vbNo Then
                'go back
            Else
                'cancel and lose changes
                Hourglass True
                ChangeScreenMode (VIEW_MODE)
                SelectLastSelected
            End If
            
        Case vbNullString
            'cancel and lose changes
            Unload Me
    End Select

Xit:
    Hourglass False
    Exit Sub

cmdCancel_ClickErr:
    ShowUnexpectedError MODULE + "cmdCancel_Click", Err
    Resume Xit

End Sub
    

Private Sub lstUsers_Click()

On Error GoTo lstUsers_ClickErr

'********************************************************************************
'* Name: lstUsers_Click
'*
'* Description:
'* Parameters:
'* Created: 6/03/99 2:13
'********************************************************************************
    Hourglass True
    With rsUser
        .MoveFirst
        While .Fields("USER_ID") <> lstUsers.Text
            .MoveNext
        Wend
    End With
    Call UpdateFields
    If lstUsers.Text = "MASTER" Then
        cmdEdit.Enabled = False
        'cmdPermissions.Enabled = False
    Else
        cmdEdit.Enabled = True
'        cmdPermissions.Enabled = True
    End If
    
Xit:
    Hourglass False
    Exit Sub

lstUsers_ClickErr:
    ShowUnexpectedError MODULE + "lstUsers_Click", Err
    Resume Xit

End Sub



Private Sub UpdateList()

On Error GoTo UpdateListErr

'********************************************************************************
'* Name: UpdateList
'*
'* Description:
'* Parameters:
'* Created: 6/3/99
'********************************************************************************
    
'    If sstabSecurity.Tab = 0 Then
        lstUsers.Clear
        lstUsers.Enabled = True
        Set cmdSecurity.ActiveConnection = gcnDDS
        With cmdSecurity
            .CommandType = adCmdText
            '.CommandText = "SELECT * FROM DD_USER where USER_ID = " & "'" & gobjLoginInfo.UserId & "'"
            .CommandText = "SELECT * FROM DD_USER ORDER BY LOWER(USER_ID) "
'                           "(Select USER_ID FROM PF_INSTITUTION_USER where INSTITUTION_ID IN " & _
'                           "(Select INSTITUTION_ID FROM PF_INSTITUTION_USER where USER_ID = " & "'" & gobjLoginInfo.UserId & "'))"
            Set rsUser = .Execute
        End With
        With rsUser
    
            Do Until .EOF
                lstUsers.AddItem ConvertNull(!USER_ID)
                .MoveNext
            Loop
            .MoveFirst
        End With
        Call UpdateFields

Xit:
    Exit Sub

UpdateListErr:
    Hourglass False
    ShowUnexpectedError MODULE + "UpdateList", Err
    Resume Xit

End Sub

Private Sub UpdateFields()

On Error GoTo UpdateFieldsErr

'********************************************************************************
'* Name: UpdateFields
'*
'* Description:
'* Parameters:
'* Created: 6/30/99 3:04:29 PM
'********************************************************************************

        '*******************************************
        ' fill in the user detail information
        '*******************************************
        Dim strAcctDisabled As String
        With rsUser
            edbUserID = !USER_ID
            txtLastName = !USER_LAST_NAME
            txtFirstName = !USER_FIRST_NAME
            strAcctDisabled = !RECORD_STATUS
            txtPassword.Text = "********"
            txtPasswordConfirm.Text = "********"
            txtCreatedBy = !CREATED_BY
            txtCreatedDate = Format(!CREATED_DATETIME, "MM/DD/YYYY")
            txtModifiedBy = ConvertNull(!LAST_MOD_BY)
            txtModifiedDate = Format((ConvertNull(!LAST_MOD_DATETIME)), "MM/DD/YYYY")
            'dblIncomeSourceTypeID = !INCOME_SOURCE_TYPE_ID
            
            
            If strAcctDisabled = strActive Then
                chkDisabled = 0
            End If
            If strAcctDisabled = strInactive Then
                chkDisabled = 1
            End If
        End With
    '*******************************************
    ' end of user tab update
    '*******************************************
'    End If

'    If sstabSecurity.Tab = 1 Then
    '*******************************************
    ' start of role tab update
    '*******************************************
'        With rsRole
''            txtRoleCode = !ROLE_CODE
''            txtRoleDescription = !ROLE_DESCR
'            dblRoleID = !ROLE_ID
'            Call UpdateRolePermissions
'        End With
    '*******************************************
    ' end of role tab update
    '*******************************************
'    End If

Xit:
    Exit Sub

UpdateFieldsErr:
    Hourglass False
    ShowUnexpectedError MODULE + "UpdateFields", Err
    Resume Xit

End Sub



Private Function DataValidation() As Boolean

On Error GoTo DataValidationErr

'********************************************************************************
'* Name: DataValidation
'*
'* Description:
'* Parameters:
'* Created: 6/03/99 3:54
'********************************************************************************
    
    Dim strEmptyFields As String
    Style = vbOKOnly + vbExclamation + vbApplicationModal
    strMessage = "The following data is required:    " & vbCrLf & vbCrLf
    strTitle = "Invalid Data"
    
    If Trim(edbUserID.Text) = vbNullString Then
        strEmptyFields = " Login Name "
        edbUserID.SetFocus
    End If

    If Trim(txtLastName.Text) = vbNullString And strEmptyFields <> vbNullString Then
        strEmptyFields = strEmptyFields & vbCrLf & " Last name "
    Else
        If Trim(txtLastName.Text) = vbNullString Then
        strEmptyFields = strEmptyFields & " Last name "
        txtLastName.SetFocus
        End If
    End If
    
    If Trim(txtFirstName.Text) = vbNullString And strEmptyFields <> vbNullString Then
        strEmptyFields = strEmptyFields & vbCrLf & " First name "
    Else
        If Trim(txtFirstName.Text) = vbNullString Then
        strEmptyFields = strEmptyFields & " First name "
        txtFirstName.SetFocus
        End If
    End If

    If strEmptyFields <> vbNullString Then
        strMessage = strMessage & strEmptyFields & vbCrLf
        Hourglass False
        If MsgBox(strMessage, Style, strTitle) = vbOK Then
            DataValidation = False
        End If
    Else
        DataValidation = True
    End If

Xit:
    Exit Function

DataValidationErr:
    Hourglass False
    ShowUnexpectedError MODULE + "DataValidation", Err
    Resume Xit

End Function

Private Function fnCkUserExists() As Boolean

On Error GoTo fnCkUserExistsErr

'********************************************************************************
'* Name: fnCkUserExists
'*
'* Description:
'* Parameters:
'* Created:
'********************************************************************************

Dim sSql As String
Dim rs As New ADODB.Recordset


    'if trying to insert a new user id, then see if it already exists, otherwise it doesn't matter
    If strUpdateMode = strInsert Then
        sSql = "select USER_ID from DD_USER where USER_ID = " & "'" & Trim(edbUserID.Text) & "'"
        rs.Open sSql, gcnDDS, adOpenForwardOnly, adLockReadOnly, adCmdText
        Dim bUserExists As Boolean
        If rs.EOF = True Then
            bUserExists = False
        Else
            bUserExists = True
        End If
        rs.Close
        Set rs = Nothing
        If bUserExists = True Then
            Style = vbOKOnly + vbExclamation + vbApplicationModal
            strTitle = "PFS Login Exists"
            strMessage = "This login name exists, please enter a different value"
            Hourglass False
            If MsgBox(strMessage, Style, strTitle) = vbOK Then
                fnCkUserExists = True
                'set focus to user id entry field
                edbUserID.SetFocus
            End If
        Else
            fnCkUserExists = False
        End If
    End If

Xit:
    Exit Function

fnCkUserExistsErr:
    Hourglass False
    ShowUnexpectedError MODULE + "fnCkUserExists", Err
    Resume Xit

End Function

Private Sub iuUser()

On Error GoTo iuUserErr

'********************************************************************************
'* Name: iuUser
'*
'* Description:
'* Parameters:
'* Created: 7/1/1999
'********************************************************************************

    If chkDisabled = 1 Then
          strStatus = strInactive
    ElseIf chkDisabled = 0 Then
          strStatus = strActive
    End If
    
    Dim ix As Integer
    
'Procedure called will:
' - add the user to the pfs database and assign to public group
'        sp_adduser 'userid'  (SYNTAX)
' - add the pfs database user information (PF_USER)
    Dim cmdUser As New ADODB.Command
    Dim oSHA256 As CSHA256
    Set oSHA256 = New CSHA256

    Set cmdUser.ActiveConnection = gcnDDS
        If gStoredProcs("up_iu_User_New").GetStoredProcCommand(cmdUser) = True Then
            cmdUser.Parameters("user_id") = Trim(edbUserID.Text)
            cmdUser.Parameters("user_last_name") = Trim(txtLastName.Text)
            cmdUser.Parameters("user_first_name") = Trim(txtFirstName.Text)
            cmdUser.Parameters("record_status") = strStatus
            'Only pass the password if there is a change
            If strNewPassword <> vbNullString Then
                cmdUser.Parameters("password") = oSHA256.SHA256(Trim(edbUserID.Text) & Trim(strNewPassword) & EXTENDSTR)
            Else
                cmdUser.Parameters("password") = Null
            End If
            cmdUser.Parameters("created_by") = gobjLoginInfo.UserId
            cmdUser.Parameters("called_from_another_proc") = "N"
            cmdUser.Parameters("update_mode") = strUpdateMode
            cmdUser.Execute
            If cmdUser.Parameters("RETURN_VALUE") <> 0 Then
                GetServerErrorMsg cmdUser.Parameters("RETURN_VALUE"), "Error occurred adding or updating user information."
            End If
        Else
            MsgBox "Error creating Insert/Update User Stored Procedure.", vbCritical
            Set cmdUser = Nothing
            ExitApp
        End If
    
    Set cmdUser = Nothing
    If strNewPassword <> vbNullString Then
        With gobjLoginInfo
            If .UserId = Trim(edbUserID.Text) Then .UserPassword = strNewPassword
            .ConnectString = "Provider=MSDASQL.1;DRIVER={" & .DBDriver & "};UID=" & .UserId & ";PWD=" & .UserPassword & ";database=" & .DBName & ";SRVR=" & .DDSServer
        End With
    End If

   'Close the connection with sso rights
            
    
Xit:
    Set oSHA256 = Nothing
     Exit Sub

iuUserErr:
    Hourglass False
    MsgBox Error, vbInformation
    Resume Xit

End Sub

Private Sub delUser()
'********************************************************************************
'* Name: delUser
'*
'* Description: This procedure must do the following:
'* Get an active connection with sso rights using database pfs
'*    1 - Delete the user roles assigned to the userid (DD_USER_ROLES)
'*    2 - Delete the institutions rights assigned to the userid (DD_INSTITUTION_USER)
'*    3 - Delete the userid from the user table (DD_USER)
'*    4 - Drop the user from the pfs database (sp_dropuser 'userid')
'*    5 - Drop the login from the Sybase SQL server (sp_droplogin 'userid')
'* Parameters:
'* Created: 7/1/1999
'********************************************************************************
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = gcnDDS
    If gStoredProcs("up_d_User").GetStoredProcCommand(cmd) = True Then
        cmd.Parameters("ddsuser_id") = Trim(edbUserID.Text)
        cmd.Execute
    End If
    Set cmd = Nothing
End Sub

Private Sub ChangeScreenMode(ByVal iMode As ScreenMode)

On Error GoTo ChangeScreenModeErr

'********************************************************************************
'* Name: ChangeScreenMode
'*
'*
'* Description:
'*   This subroutine will change the background of certain controls and enable
'*   or disable controls and buttons depending on whether you are adding a code,
'*   editing a code, or viewing the active codes.
'*
'* Parameters: iMode - The choices are ADD_MODE, VIEW_MODE or UPDATE_MODE
'* Created:
'********************************************************************************

    'Change the mode for the screen
    iEditMode = iMode
    
    Select Case iMode
    
    Case VIEW_MODE
        'Delete button is not currently being used, so disable it and hide it
        cmdDelete.Enabled = False
        cmdDelete.Visible = False
'        clear the controls
'        edbUserID.Text = vbNullString
        txtFirstName.Text = vbNullString
        txtLastName.Text = vbNullString
        txtPassword.Text = vbNullString
        txtPasswordConfirm.Text = vbNullString
        chkDisabled = 0
'        lstUserInstitutions.Clear
'        lstAvailableInstitutions.Clear
'        lstUserRoles.Clear
'        lstAvailableRoles.Clear
        'restore the background of the controls to white while in view mode
'        edbUserID.BackColor = DFLT_WHITE
        txtFirstName.BackColor = DFLT_WHITE
        txtLastName.BackColor = DFLT_WHITE
        txtPassword.BackColor = DFLT_WHITE
        txtPasswordConfirm.BackColor = DFLT_WHITE

        'enable the controls that are modifiable during view
        'set focus
        'enable all buttons available during initial view
        lstUsers.Enabled = True
        cmdNew.Enabled = True
        'disable the controls that cannot be changed during view
        '*******************'
        '**** Users tab ****'
        '*******************'
'        edbUserID.Enabled = False
        'edbUserID.Locked = True
        txtFirstName.Enabled = False
        txtFirstName.Locked = True
        txtLastName.Enabled = False
        txtLastName.Locked = True
        txtPassword.Enabled = False
        txtPassword.Locked = True
        txtPasswordConfirm.Enabled = False
        txtPasswordConfirm.Locked = True
        chkDisabled.Enabled = False
'        lstUserInstitutions.Enabled = False
'        lstAvailableInstitutions.Enabled = False
'        lstUserRoles.Enabled = False
'        lstAvailableRoles.Enabled = False
        '*******************'
        '**** Roles tab ****'
        '*******************'
'        txtRoleCode.Enabled = False
'        txtRoleCode.Locked = True
'        txtRoleDescription.Enabled = False
'        txtRoleDescription.Locked = True
        'disable all buttons not available during view
        cmdEdit.Enabled = False
        'cmdPermissions.Enabled = False
        cmdApply.Enabled = False
        cmdOK.Enabled = False
'        cmdAddInstitution.Enabled = False
'        cmdRemoveInstitution.Enabled = False
'        cmdAddRole.Enabled = False
'        cmdRemoveRole.Enabled = False
        Call UpdateList
        strUpdateMode = vbNullString
                
    Case ADD_MODE
        'clear the controls
'        edbUserID.Text = vbNullString
        edbUserID.Text = vbNullString
        txtFirstName.Text = vbNullString
        txtLastName.Text = vbNullString
        txtPassword.Text = vbNullString
        txtPasswordConfirm.Text = vbNullString
        txtCreatedBy.Text = vbNullString
        txtCreatedDate.Text = vbNullString
        txtModifiedBy = vbNullString
        txtModifiedDate = vbNullString
        
        chkDisabled = 0
'        lstUserInstitutions.Clear
'        lstAvailableInstitutions.Clear
'        lstUserRoles.Clear
'        lstAvailableRoles.Clear
        'change the background of the controls that are mandatory for an add
'        edbUserID.BackColor = PALE_YELLOW
        txtFirstName.BackColor = PALE_YELLOW
        txtLastName.BackColor = PALE_YELLOW
        txtPassword.BackColor = PALE_YELLOW
        txtPasswordConfirm.BackColor = PALE_YELLOW
        'enable the controls that are modifiable during an add
'        edbUserID.Enabled = True
        'edbUserID.Locked = False
        txtFirstName.Enabled = True
        txtFirstName.Locked = False
        txtLastName.Enabled = True
        txtLastName.Locked = False
        txtPassword.Enabled = True
        txtPassword.Locked = False
        txtPasswordConfirm.Enabled = True
        txtPasswordConfirm.Locked = False
        chkDisabled.Enabled = True
'        lstUserInstitutions.Enabled = True
'        lstUserRoles.Enabled = True
'        lstAvailableInstitutions.Enabled = True
'        lstAvailableRoles.Enabled = True
        'set focus to first control
'        edbUserID.SetFocus
        'enable all buttons available during an add
        cmdOK.Enabled = True
        cmdApply.Enabled = True
        'disable the controls that cannot be changed during an add
        lstUsers.Enabled = False
        'disable all buttons not available during an add
        cmdNew.Enabled = False
        cmdEdit.Enabled = False
        
' This section left for future coding for Permission
'
'        cmdPermissions.Enabled = False
'        '******************************************
'        ' populate the available institutions list
'        '******************************************
'        Dim rsAvailable As New ADODB.Recordset
'        With cmdSecurity
'            .CommandType = adCmdText
'            .CommandText = "Select INSTITUTION_NAME FROM PF_INSTITUTION where INSTITUTION_ID IN (Select INSTITUTION_ID FROM PF_INSTITUTION_USER where USER_ID = " & "'" & gobjLoginInfo.UserId & "'" & ") and RECORD_STATUS = 'A'"
'            Set rsAvailable = .Execute
'        End With
'        With rsAvailable
'            Do Until .EOF
'                lstAvailableInstitutions.AddItem !INSTITUTION_NAME
'                .MoveNext
'            Loop
'        End With
        '*******************************************
        ' populate the available roles list
        '*******************************************
'        With cmdSecurity
'            .CommandType = adCmdText
'            .CommandText = "Select ROLE_CODE FROM PF_ROLE where RECORD_STATUS = 'A'"
'            Set rsAvailable = .Execute
'        End With
'        With rsAvailable
'            Do Until .EOF
''                lstAvailableRoles.AddItem !ROLE_CODE
'                .MoveNext
'            Loop
'        End With
'        rsAvailable.Close
'        Set rsAvailable = Nothing
'
'

        
    Case EDIT_MODE
        'change the background of the controls that are mandatory for an edit
'        edbUserID.BackColor = PALE_YELLOW
        txtFirstName.BackColor = PALE_YELLOW
        txtLastName.BackColor = PALE_YELLOW
        txtPassword.BackColor = PALE_YELLOW
        txtPasswordConfirm.BackColor = PALE_YELLOW
        'enable the controls that are modifiable during an edit
        txtFirstName.Enabled = True
        txtFirstName.Locked = False
        txtLastName.Enabled = True
        txtLastName.Locked = False
        txtPassword.Enabled = True
        txtPassword.Locked = False
        txtPasswordConfirm.Enabled = True
        txtPasswordConfirm.Locked = False
        If lstUsers.Text = gobjLoginInfo.UserId Then
            chkDisabled.Enabled = False
        Else
            chkDisabled.Enabled = True
        End If
        'lstUserInstitutions.Enabled = True
'        lstUserRoles.Enabled = True
'        lstAvailableInstitutions.Enabled = True
'        lstAvailableRoles.Enabled = True
        'set focus to first control
        txtLastName.SetFocus
        'enable all buttons available during an add
        cmdOK.Enabled = True
        cmdApply.Enabled = True
        'disable the controls that cannot be changed during an edit
        lstUsers.Enabled = False
        'disable all buttons not available during an edit
'        cmdNew.Enabled = False
'        cmdEdit.Enabled = False
'        cmdDelete.Enabled = False
       ' cmdPermissions.Enabled = False
    
    End Select

Xit:
    Exit Sub

ChangeScreenModeErr:
    Hourglass False
    ShowUnexpectedError MODULE + "ChangeScreenMode", Err
    Resume Xit

End Sub

Private Sub Form_Unload(Cancel As Integer)

'********************************************************************************
'* Name: Form_Unload
'*
'* Description:
'* Parameters:
'* Created:
'********************************************************************************
    
    On Error Resume Next
'    rsRole.Close
'    Set rsRole = Nothing
    rsUser.Close
    Set rsUser = Nothing
    Set cmdSecurity = Nothing

End Sub


Private Function fnPasswordValidation() As Boolean

On Error GoTo fnPasswordValidationErr

'********************************************************************************
'* Name: fnPasswordValidation
'*
'* Description:
'* Parameters:
'* Created: 3/30/99 3:18:20 PM
'********************************************************************************
            
    If Len(txtPassword.Text) < 6 Or Len(txtPasswordConfirm.Text) < 6 Then
        Style = vbOKOnly + vbExclamation + vbApplicationModal
        strTitle = "Password entry error"
        strMessage = "Password specified is too short. Minimum length of acceptable passwords is 6 characters."
        Hourglass False
        If MsgBox(strMessage, Style, strTitle) = vbOK Then
            fnPasswordValidation = False
            'set focus to password entry field
            txtPassword.SetFocus
        End If
    Else
        If txtPassword.Text <> txtPasswordConfirm.Text Then
            Style = vbOKOnly + vbExclamation + vbApplicationModal
            strTitle = "Password entry error"
            strMessage = "The password did not confirm properly. Please make sure the password and confirm password fields are identical."
            Hourglass False
            If MsgBox(strMessage, Style, strTitle) = vbOK Then
                fnPasswordValidation = False
                'set focus to password entry field
                txtPassword.SetFocus
            End If
        Else
            fnPasswordValidation = True
        End If
    End If
    
    If fnPasswordValidation = True Then
        If strUpdateMode = strUpdate Then
            If txtPassword.Text = "********" Then
                strNewPassword = vbNullString
            Else
                strNewPassword = txtPassword.Text
            End If
        Else
            If strUpdateMode = strInsert Then
                strNewPassword = txtPassword.Text
            End If
        End If
    End If
    
Xit:
    Exit Function

fnPasswordValidationErr:
    Hourglass False
    ShowUnexpectedError MODULE + "fnPasswordValidation", Err
    Resume Xit

End Function

Private Sub OutlookTitle1_IconClick()
    
    If cmdCancel.Enabled = True Then
        Unload Me
    End If

End Sub

Private Sub txtPasswordConfirm_GotFocus()
    Call SetSelected
End Sub

Private Sub txtPassword_GotFocus()
    Call SetSelected
End Sub

Private Sub txtFirstName_GotFocus()
    Call SetSelected
End Sub

Private Sub txtLastName_GotFocus()
    Call SetSelected
End Sub

Private Sub edbUserID_GotFocus()
    Call SetSelected
End Sub

Private Function CheckUserPrivileges() As Boolean

On Error GoTo CheckUserPrivilegesErr

'********************************************************************************
'* Name: CheckUserPrivileges
'*
'* Description:
'* Parameters:
'* Created: 3/30/99 3:18:52 PM
'* Modifications:
'* check to see if sybase user exists in system
'* and sybase user for just this database

'********************************************************************************

Dim cnnPFSsso As New ADODB.Connection
Dim bSybaseLoginExists As Boolean
Dim bDBUserExists As Boolean
Dim rsSybaseLogin As New ADODB.Recordset
Dim rsDBUser As New ADODB.Recordset
Dim sSql As String
    
    
    'Get an active connection with sso rights
    
    With gobjLoginInfo
    cnnPFSsso.ConnectionString = "DRIVER={" & .DBDriver & "};UID=" & strSSOUser & ";PWD=" & strSSOPassword & ";database=" & .DBName & ";SRVR=" & .DDSServer
    End With
    cnnPFSsso.CursorLocation = adUseClient
    cnnPFSsso.Open
    
    If strUpdateMode = strInsert Then
    'if trying to insert a new user id, then see if this Sybase login already exists, otherwise it doesn't matter
        bSybaseLoginExists = False
        bDBUserExists = False
        
        sSql = "Select * from dds.dbo.sysusers where name = " & "'" & edbUserID.Text & "'"
        rsDBUser.Open sSql, cnnPFSsso, adOpenForwardOnly, adLockReadOnly, adCmdText
        If rsDBUser.EOF = True Then
            bDBUserExists = False
        Else
            bDBUserExists = True
        End If
        
        rsDBUser.Close
        Set rsDBUser = Nothing
            
        sSql = "Select * from master.dbo.syslogins where name = " & "'" & edbUserID.Text & "'"
        rsSybaseLogin.Open sSql, cnnPFSsso, adOpenForwardOnly, adLockReadOnly, adCmdText
        If rsSybaseLogin.EOF = True Then
            bSybaseLoginExists = False
        Else
            bSybaseLoginExists = True
        End If
        rsSybaseLogin.Close
        Set rsSybaseLogin = Nothing
        
        If bSybaseLoginExists = True And bDBUserExists = True Then
            Style = vbOKOnly + vbExclamation + vbApplicationModal
            strTitle = "Sybase Login Exists"
            strMessage = "This Sybase login and username already exist, please enter a different value"
            Hourglass False
            If MsgBox(strMessage, Style, strTitle) = vbOK Then
                CheckUserPrivileges = True
                'set focus to user id entry field
                edbUserID.SetFocus
            End If
        Else
            If bDBUserExists = True Then
                Style = vbOKOnly + vbExclamation + vbApplicationModal
                strTitle = "Sybase Username Exists"
                strMessage = "This Sybase username already exists, please enter a different value"
                Hourglass False
                If MsgBox(strMessage, Style, strTitle) = vbOK Then
                    CheckUserPrivileges = True
                    'set focus to user id entry field
                    edbUserID.SetFocus
                End If
            Else
                CheckUserPrivileges = False
            End If
        End If
    End If

'Close the connection with sso rights
    cnnPFSsso.Close

Xit:
    Exit Function

CheckUserPrivilegesErr:
    Hourglass False
    ShowUnexpectedError MODULE + "CheckUserPrivileges", Err
    Resume Xit
'
End Function

'This section left for future coding for Permission
'
'Private Function fnCkInstitutionRights() As Boolean
'
'On Error GoTo fnCkInstitutionRightsErr
'
''********************************************************************************
''* Name: fnCkInstitutionRights
''*
''* Description:
''* Parameters:
''* Created: 3/30/99 3:19:31 PM
''********************************************************************************
'
'    If lstUserInstitutions.ListCount = 0 Then
'        Style = vbOKOnly + vbExclamation + vbApplicationModal
'        strTitle = "Institution Access"
'        strMessage = "User must be assigned access to at least one institution, please select an institution. "
'        Hourglass False
'        If MsgBox(strMessage, Style, strTitle) = vbOK Then
'            fnCkInstitutionRights = False
'        End If
'    Else
'        fnCkInstitutionRights = True
'    End If
'
'Xit:
'    Exit Function
'
'fnCkInstitutionRightsErr:
'    Hourglass False
'    ShowUnexpectedError MODULE + "fnCkInstitutionRights", Err
'    Resume Xit
'
'End Function

'********************************************************************************
''* Name: ConvertNull
''*
''* Description: 'This function is to Check if it is Null value for the field
''* Parameters: vValue
''* Created: 5/30/99 4:11:11 PM
''********************************************************************************

Public Function ConvertNull(ByVal vValue As Variant) As Variant
    If IsNull(vValue) Then
        ConvertNull = vbNullString
    Else
        ConvertNull = vValue
    End If
End Function
