VERSION 5.00
Object = "{8CD222DF-7752-11D3-9D1E-00105A19BCF2}#1.0#0"; "OAOTBar.ocx"
Begin VB.Form frmIncomeSourceTypes 
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   10005
   WindowState     =   2  'Maximized
   Begin OAOTitleBar.OutlookTitleBar OutlookTitle1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   979
      ForeColor       =   16777215
      Caption         =   "Income Source Types"
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   6420
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   6420
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   6420
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Codes"
      Height          =   5520
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   9495
      Begin VB.Frame Frame4 
         Caption         =   "Details"
         Height          =   5115
         Left            =   4080
         TabIndex        =   10
         Top             =   240
         Width           =   5295
         Begin VB.TextBox txtNCASAccount 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   32
            Top             =   3000
            WhatsThisHelpID =   30036
            Width           =   2655
         End
         Begin VB.TextBox txtPACode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   24
            Top             =   1215
            WhatsThisHelpID =   30036
            Width           =   2655
         End
         Begin VB.TextBox txtPAPMTCode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   23
            Top             =   1572
            WhatsThisHelpID =   30036
            Width           =   2655
         End
         Begin VB.TextBox txtPAPMTREVCode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   22
            Top             =   1929
            WhatsThisHelpID =   30036
            Width           =   2655
         End
         Begin VB.TextBox txtPFDEPTRANSCode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   21
            Top             =   2286
            WhatsThisHelpID =   30036
            Width           =   2655
         End
         Begin VB.TextBox txtPFDEPREVTRANSCode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   20
            Top             =   2643
            WhatsThisHelpID =   30036
            Width           =   2655
         End
         Begin VB.TextBox txtStart 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   19
            Top             =   3357
            WhatsThisHelpID =   30036
            Width           =   2655
         End
         Begin VB.TextBox txtLength 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   18
            Top             =   3720
            WhatsThisHelpID =   30036
            Width           =   2655
         End
         Begin VB.TextBox txtDescription 
            Enabled         =   0   'False
            Height          =   435
            HelpContextID   =   30037
            Left            =   2490
            Locked          =   -1  'True
            MaxLength       =   80
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   660
            WhatsThisHelpID =   30036
            Width           =   2655
         End
         Begin VB.TextBox txtCreatedBy 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   210
            Left            =   2490
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   4140
            Width           =   2655
         End
         Begin VB.TextBox txtModifiedBy 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   2490
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   4530
            Width           =   2595
         End
         Begin VB.TextBox txtFUNBCode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   4
            Top             =   300
            WhatsThisHelpID =   30036
            Width           =   2655
         End
         Begin VB.Label Label14 
            Caption         =   "NCAS Account:"
            Height          =   285
            Left            =   210
            TabIndex        =   33
            Top             =   3015
            Width           =   2055
         End
         Begin VB.Label Label7 
            Caption         =   "PA Code:"
            Height          =   285
            Left            =   210
            TabIndex        =   31
            Top             =   1230
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "PA PMT Code:"
            Height          =   285
            Left            =   210
            TabIndex        =   30
            Top             =   1590
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "PA PMT REV Code:"
            Height          =   285
            Left            =   210
            TabIndex        =   29
            Top             =   1950
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "PF DEP TRANS Code:"
            Height          =   285
            Left            =   210
            TabIndex        =   28
            Top             =   2295
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "PF DEP REV TRANS Code:"
            Height          =   285
            Left            =   210
            TabIndex        =   27
            Top             =   2655
            Width           =   2055
         End
         Begin VB.Label Label12 
            Caption         =   "DD Number Start Position:"
            Height          =   285
            Left            =   210
            TabIndex        =   26
            Top             =   3375
            Width           =   2055
         End
         Begin VB.Label Label13 
            Caption         =   "DD Number Length:"
            Height          =   285
            Left            =   210
            TabIndex        =   25
            Top             =   3735
            Width           =   2055
         End
         Begin VB.Label Label6 
            Caption         =   "Description:"
            Height          =   285
            Left            =   240
            TabIndex        =   16
            Top             =   645
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Record Modified By:"
            Height          =   255
            Left            =   195
            TabIndex        =   13
            Top             =   4530
            Width           =   1485
         End
         Begin VB.Label Label3 
            Caption         =   "Record Created By:"
            Height          =   255
            Left            =   195
            TabIndex        =   12
            Top             =   4140
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "FUNB Code:"
            Height          =   405
            Left            =   240
            TabIndex        =   11
            Top             =   330
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.ListBox lstIncomeSource 
         Height          =   2790
         ItemData        =   "frmIncomeSourceTypes.frx":0000
         Left            =   240
         List            =   "frmIncomeSourceTypes.frx":0007
         TabIndex        =   1
         Top             =   375
         Width           =   2535
      End
   End
   Begin VB.Line lin1 
      BorderColor     =   &H80000010&
      BorderStyle     =   6  'Inside Solid
      X1              =   120
      X2              =   9600
      Y1              =   6300
      Y2              =   6300
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   9600
      Y1              =   6300
      Y2              =   6300
   End
End
Attribute VB_Name = "frmIncomeSourceTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************
' * Form Name:frmIncomeSourceType
' * Form File Name: frmIncomeSourceType.frm
' * Start Date: 6/19/1999
' * End Date:   7/26/1999
' * Description:
' * --------------------------------
' * The The Income Source Types screen allows the user to view
'   · modify and add new Codes based on the Income source types
'
'
' Mod CONSTANTS
Private Const MODULE As String = "Income_Source_Types"
Private Const strInsert As String = "I"
Private Const strUpdate As String = "U"
'

' Mod DECLARES
Private cmd As New ADODB.Command
Private rsIncomeSourceType As New ADODB.Recordset

' Mod VARIABLES
Private iEditMode As ScreenMode
Private strUpdateMode As String
Private strMessage As String
Private strTitle As String
Private Style As VbMsgBoxStyle
Private dblIncomeSourceTypeID As Double
Private msLastSelected As String

 
Private Sub Form_Activate()
    'Set the main toolbar to not see the some icons
    fMainForm.SetMainToolbar True

End Sub


Private Sub cmdCancel_Click()

'********************************************************************************
'* Name: cmdCancel_Click
'* Description:
'* Created: 6/2/99 5:30 PM
'********************************************************************************
    'Unload Me
    On Error GoTo cmdCancel_ClickErr
    msLastSelected = txtFUNBCode

    Style = vbYesNo + vbQuestion + vbDefaultButton2 + vbApplicationModal
    strMessage = "Are you sure you want to cancel?"
        
    Select Case strUpdateMode
        
        Case strInsert
            If MsgBox(strMessage, Style) = vbNo Then
                'go back
            Else
                'cancel and lose changes
                ChangeScreenMode (VIEW_MODE)
                SelectLastSelected
            End If

        Case strUpdate
            If MsgBox(strMessage, Style) = vbNo Then
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

Private Sub cmdNew_Click()

'********************************************************************************
'* Name: cmdNew_Click
'* Description:
'* Created: 6/2/99 3:26 PM
'********************************************************************************
   
    strUpdateMode = strInsert
    ChangeScreenMode (ADD_MODE)
End Sub
Private Sub cmdEdit_Click()
'********************************************************************************
'* Name: cmdEdit_Click
'* Description:
'* Created: 6/2/99 3:26 PM
'********************************************************************************
    
    strUpdateMode = strUpdate
    ChangeScreenMode (EDIT_MODE)
End Sub

Private Sub cmdDelete_Click()

On Error GoTo cmdDelete_ClickErr

'********************************************************************************
'* Name: cmdDelete_Click
'* Description:
'* Created: 6/2/99 3:26 PM
'********************************************************************************
    Style = vbYesNo + vbExclamation + vbDefaultButton2 + vbApplicationModal
    strMessage = "Are you sure you want to delete this code?"
    strTitle = "Confirm Code Deletion"
    If MsgBox(strMessage, Style, strTitle) = vbNo Then
        'disable buttons
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    Else
        'run delete procedure
        Call delIncomeSourceType
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

Private Sub cmdApply_Click()

On Error GoTo cmdApply_ClickErr

'********************************************************************************
'* Name: cmdApply_Click
'* Description:
'* Created: 6/2/99 3:27 PM
'********************************************************************************
    msLastSelected = txtFUNBCode
    'check to see if this code already exists
    If fnCkCodeExists = False Then
        'this code doesn't exist so check that the fields are filled out properly
        If DataValidation = True Then
            'the fields are filled out so execute procedure to update the record
            Call iuIncomeSourceType
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
    Exit Sub

cmdApply_ClickErr:
   ShowUnexpectedError MODULE + "cmdApply_Click", Err
    Resume Xit
End Sub


Private Sub SelectLastSelected()
'*************************************************
'* Selects the last entered
'************************************************
Dim iX As Long
    For iX = 0 To lstIncomeSource.ListCount - 1
        If lstIncomeSource.List(iX) = msLastSelected Then
            lstIncomeSource.Selected(iX) = True
            lstIncomeSource.SetFocus
            Exit For
        End If
    Next iX
    
End Sub

Private Sub cmdOK_Click()

On Error GoTo cmdOK_ClickErr

'********************************************************************************
'* Name: cmdOk_Click
'* Description:
'* Created: 6/2/99 3:29 PM
'********************************************************************************
    
    'check to see if this code already exists
    If fnCkCodeExists = False Then
       'this code doesn't exist so check that the fields are filled out properly
        If DataValidation = True Then
            'the fields are filled out so execute procedure to update the record
            Call iuIncomeSourceType
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
    ShowUnexpectedError MODULE + "cmdOK_Click", Err
    Resume Xit


End Sub

Private Sub Form_Deactivate()
    
    fMainForm.SetMainToolbar False

End Sub

Private Sub Form_Load()
'********************************************************************************
'* Name: Form_Load
'* Description:
'* Created: 6/26/99 11:10 AM
'********************************************************************************
    Set OutlookTitle1.Picture = fMainForm.imlToolbarIcons.ListImages("Income Source Types").Picture
    Call UpdateList
    'Call the Procedure based on the mode
    ChangeScreenMode (VIEW_MODE)
    cmdOK.Enabled = False
    cmdApply.Enabled = False
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
'* Created: 6/02/99 6:09 PM
'********************************************************************************
    'Change the mode for the screen
    
    Hourglass True
    iEditMode = iMode
    
    Select Case iMode
    
    Case VIEW_MODE
        'restore the background of the controls to white while in view mode
        txtFUNBCode.BackColor = DFLT_WHITE
        txtDescription.BackColor = DFLT_WHITE
        txtPACode.BackColor = DFLT_WHITE
        txtPAPMTCode.BackColor = DFLT_WHITE
        txtPAPMTREVCode.BackColor = DFLT_WHITE
        txtPFDEPTRANSCode.BackColor = DFLT_WHITE
        txtPFDEPREVTRANSCode.BackColor = DFLT_WHITE
        txtNCASAccount.BackColor = DFLT_WHITE
        txtStart.BackColor = DFLT_WHITE
        txtLength.BackColor = DFLT_WHITE
        
        'enable the controls that are modifiable during view
        'set focus
        'enable all buttons available during initial view
        cmdNew.Enabled = True
        'disable the controls that cannot be changed during view
        txtFUNBCode.Enabled = False
        txtFUNBCode.Locked = True
        txtDescription.Enabled = False
        txtDescription.Locked = True
        txtPACode.Enabled = False
        txtPACode.Locked = True
        txtPAPMTCode.Enabled = False
        txtPAPMTCode.Locked = True
        txtPAPMTREVCode.Enabled = False
        txtPAPMTREVCode.Locked = True
        txtPFDEPTRANSCode.Enabled = False
        txtPFDEPTRANSCode.Locked = True
        txtPFDEPREVTRANSCode.Enabled = False
        txtPFDEPREVTRANSCode.Locked = True
        txtNCASAccount.Enabled = False
        txtNCASAccount.Locked = True
        txtStart.Enabled = False
        txtStart.Locked = True
        txtLength.Enabled = False
        txtLength.Locked = True
        
        'disable all buttons not available during view
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        cmdApply.Enabled = False
        cmdOK.Enabled = False
        Call UpdateList
        'Call UpdateIncomeSourceList
        strUpdateMode = vbNullString
        
    Case ADD_MODE
        'clear the controls
        txtFUNBCode.Text = vbNullString
        txtDescription.Text = vbNullString
        txtPACode.Text = vbNullString
        txtPAPMTCode.Text = vbNullString
        txtPAPMTREVCode.Text = vbNullString
        txtPFDEPTRANSCode.Text = vbNullString
        txtPFDEPREVTRANSCode.Text = vbNullString
        txtNCASAccount.Text = vbNullString
        txtCreatedBy.Text = vbNullString
        txtModifiedBy = vbNullString
        txtStart = vbNullString
        txtLength = vbNullString
        
        'change the background of the controls that are mandatory for an add
        txtFUNBCode.BackColor = PALE_YELLOW
        txtDescription.BackColor = PALE_YELLOW
        txtPACode.BackColor = PALE_YELLOW
        txtPAPMTCode.BackColor = PALE_YELLOW
        txtPAPMTREVCode.BackColor = PALE_YELLOW
        txtPFDEPTRANSCode.BackColor = PALE_YELLOW
        txtPFDEPREVTRANSCode.BackColor = PALE_YELLOW
        txtStart.BackColor = PALE_YELLOW
        txtLength.BackColor = PALE_YELLOW
        
        'enable the controls that are modifiable during an add
        txtFUNBCode.Enabled = True
        txtFUNBCode.Locked = False
        txtDescription.Enabled = True
        txtDescription.Locked = False
        txtPACode.Enabled = True
        txtPACode.Locked = False
        txtPAPMTCode.Enabled = True
        txtPAPMTCode.Locked = False
        txtPAPMTREVCode.Enabled = True
        txtPAPMTREVCode.Locked = False
        txtPFDEPTRANSCode.Enabled = True
        txtPFDEPTRANSCode.Locked = False
        txtPFDEPREVTRANSCode.Enabled = True
        txtPFDEPREVTRANSCode.Locked = False
        txtNCASAccount.Enabled = True
        txtNCASAccount.Locked = False
        txtStart.Enabled = True
        txtStart.Locked = False
        txtLength.Enabled = True
        txtLength.Locked = False
        
        'set focus to first control
        txtFUNBCode.SetFocus
        'enable all buttons available during an add
        cmdOK.Enabled = True
        cmdApply.Enabled = True
        'disable the controls that cannot be changed during an add
        lstIncomeSource.Enabled = False
        'disable all buttons not available during an add
        cmdNew.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        
    Case EDIT_MODE
        'change the background of the controls that are mandatory for an add
        txtFUNBCode.BackColor = PALE_YELLOW
        txtDescription.BackColor = PALE_YELLOW
        txtPACode.BackColor = PALE_YELLOW
        txtPAPMTCode.BackColor = PALE_YELLOW
        txtPAPMTREVCode.BackColor = PALE_YELLOW
        txtPFDEPTRANSCode.BackColor = PALE_YELLOW
        txtPFDEPREVTRANSCode.BackColor = PALE_YELLOW
        txtStart.BackColor = PALE_YELLOW
        txtLength.BackColor = PALE_YELLOW
        
        'enable the controls that are modifiable during an add
        txtDescription.Enabled = True
        txtDescription.Locked = False
        txtPACode.Enabled = True
        txtPACode.Locked = False
        txtPAPMTCode.Enabled = True
        txtPAPMTCode.Locked = False
        txtPAPMTREVCode.Enabled = True
        txtPAPMTREVCode.Locked = False
        txtPFDEPTRANSCode.Enabled = True
        txtPFDEPTRANSCode.Locked = False
        txtPFDEPREVTRANSCode.Enabled = True
        txtPFDEPREVTRANSCode.Locked = False
        txtNCASAccount.Enabled = True
        txtNCASAccount.Locked = False
        txtStart.Enabled = True
        txtStart.Locked = False
        txtLength.Enabled = True
        txtLength.Locked = False
        
        'set focus to first control
        txtDescription.SetFocus
        'enable all buttons available during an edit
        cmdOK.Enabled = True
        cmdApply.Enabled = True
        'disable the controls that cannot be changed during an add
        lstIncomeSource.Enabled = False
        'disable all buttons not available during an edit
        cmdNew.Enabled = False
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    
    End Select
    

Xit:
    Hourglass False
    Exit Sub

ChangeScreenModeErr:
    ShowUnexpectedError MODULE + "ChangeScreenMode", Err
    Resume Xit


End Sub
Private Sub Form_Unload(Cancel As Integer)
'********************************************************************************
'* Name: Form_Unload
'* Description:
'* Created: 5/11/1999 10:11:15 AM
'********************************************************************************
On Error Resume Next
rsIncomeSourceType.Close
Set rsIncomeSourceType = Nothing
Set cmd = Nothing


End Sub

'This function is to Check if it is Null value for the field
Public Function ConvertNull(ByVal vValue As Variant) As Variant
    If IsNull(vValue) Then
        ConvertNull = vbNullString
    Else
        ConvertNull = vValue
    End If
End Function
Private Sub UpdateList()

On Error GoTo UpdateListErr

'********************************************************************************
'* Name: UpdateList
'* Description:
'* Created: 6/02/99 4:17:53 PM
'********************************************************************************
    Hourglass True
    lstIncomeSource.Clear
    lstIncomeSource.Enabled = True
    
    Set cmd.ActiveConnection = gcnDDS
        
    With cmd
    .CommandType = adCmdText
    .CommandText = "Select * FROM DD_INCOME_SOURCE_TYPE WHERE RECORD_STATUS = " & "'" & "A'" & " ORDER BY FUNB_INCOME_SRC_TYPE"
    Set rsIncomeSourceType = .Execute
    End With
    With rsIncomeSourceType
    Do Until .EOF
        lstIncomeSource.AddItem ConvertNull(!FUNB_INCOME_SRC_TYPE)
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
'* Created: 5/11/99 11:02:51 PM
'********************************************************************************
'    Hourglass True
    With rsIncomeSourceType
       
        txtFUNBCode = !FUNB_INCOME_SRC_TYPE
        txtDescription = !INCOME_SRC_TYPE_DESCR
        txtPACode = !PA_INCOME_SRC_TYPE
        txtPAPMTCode = ConvertNull(!PA_PMT_CODE)
        txtPAPMTREVCode = ConvertNull(!PA_PMT_REV_CODE)
        txtPFDEPTRANSCode = ConvertNull(!PF_DEP_TRANS_CODE)
        txtPFDEPREVTRANSCode = ConvertNull(!PF_DEP_REV_TRANS_CODE)
        txtNCASAccount = ConvertNull(!NCAS_ACCOUNT)
        txtStart = ConvertNull(!START_POS)
        txtLength = ConvertNull(!Length)
        
        txtCreatedBy = !CREATED_BY & " on " & Format(!CREATED_DATETIME, "MM/DD/YYYY")
        If IsNull(!LAST_MOD_BY) Then
            txtModifiedBy = ""
        Else
            txtModifiedBy = ConvertNull(!LAST_MOD_BY) & " on " & Format((ConvertNull(!LAST_MOD_DATETIME)), "MM/DD/YYYY")
        End If
        dblIncomeSourceTypeID = !INCOME_SOURCE_TYPE_ID
    End With

Xit:
    Hourglass False
    Exit Sub

UpdateFieldsErr:
    ShowUnexpectedError MODULE + "UpdateFields", Err
    Resume Xit


End Sub
Private Sub lstIncomeSource_Click()

On Error GoTo lstIncomeSource_ClickErr

'********************************************************************************
'* Name: lstIncomeSource_Click
'* Description:
'* Created: 5/11/99 11:04:04 PM
'********************************************************************************
'    Hourglass True
    With rsIncomeSourceType
        .MoveFirst
        While .Fields("FUNB_INCOME_SRC_TYPE") <> lstIncomeSource.Text
            .MoveNext
        Wend
    End With
    Call UpdateFields
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True

Xit:
    Hourglass False
    Exit Sub

lstIncomeSource_ClickErr:
    ShowUnexpectedError MODULE + "lstIncomeSources_Click", Err
    Resume Xit


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
        'Hourglass True
        With rsIncomeSourceType
            bExists = False
            .MoveFirst
            Do Until .EOF
                If .Fields("FUNB_INCOME_SRC_TYPE") <> Trim(UCase(txtFUNBCode.Text)) Then
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
                    txtFUNBCode.SetFocus
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
Private Sub iuIncomeSourceType()

On Error GoTo iuIncomeSourceTypeErr


'********************************************************************************
'* Name: iuIncomeSourceType
'* Description:Calling the Procedure to Update the Income Source Type Screen Options
'* the stored procedure callled "up_iu_Income_Source_Type"
'* Created: 6/02/99 3:39
'********************************************************************************

    Dim cmd As New ADODB.Command
    'Replace gcnDDS----------------
'    Dim conn As New ADODB.Connection
'    Set cmd.ActiveConnection = conn
    '-------------
     Hourglass True
    Set cmd.ActiveConnection = gcnDDS
    If gStoredProcs("up_iu_Income_Source_Type").GetStoredProcCommand(cmd) = True Then
        cmd.Parameters("FUNB_INCOME_SRC_TYPE") = Trim(UCase(txtFUNBCode.Text))
        cmd.Parameters("INCOME_SRC_TYPE_DESCR") = Trim(txtDescription.Text)
        cmd.Parameters("PA_INCOME_SRC_TYPE") = Trim(txtPACode.Text)
        cmd.Parameters("PA_PMT_CODE") = Trim(txtPAPMTCode.Text)
        cmd.Parameters("PA_PMT_REV_CODE") = Trim(txtPAPMTREVCode.Text)
        cmd.Parameters("PF_DEP_TRANS_CODE") = Trim(txtPFDEPTRANSCode.Text)
        cmd.Parameters("PF_DEP_REV_TRANS_CODE") = Trim(txtPFDEPREVTRANSCode.Text)
        cmd.Parameters("NCAS_ACCOUNT") = Trim(txtNCASAccount.Text)
        cmd.Parameters("user_id") = gobjLoginInfo.UserId
        If IsNumeric(Trim(txtStart.Text)) Then
            cmd.Parameters("start_pos") = CInt(Trim(txtStart.Text))
        Else
            cmd.Parameters("start_pos") = Null
        End If
        If IsNumeric(Trim(txtLength.Text)) Then
            cmd.Parameters("length") = CInt(Trim(txtLength.Text))
        Else
            cmd.Parameters("length") = Null
        End If
        cmd.Parameters("update_mode") = strUpdateMode
        cmd.Parameters("called_from_another_proc") = "N"
        cmd.Parameters("INCOME_SOURCE_TYPE_ID") = dblIncomeSourceTypeID
        cmd.Execute
        If cmd.Parameters("RETURN_VALUE") <> 0 Then
            GetServerErrorMsg cmd.Parameters("RETURN_VALUE"), "Error occurred adding or updating the State record."
        End If
    Else
        MsgBox "Error creating the Insert/Update Income Source Type Stored Procedure.", vbCritical
        Set cmd = Nothing
        ExitApp
    End If
    Set cmd = Nothing

Xit:
    Hourglass False
    Exit Sub

iuIncomeSourceTypeErr:
    ShowUnexpectedError MODULE + "iuIncomeSourceType", Err
    Resume Xit

End Sub
Private Sub delIncomeSourceType()

On Error GoTo delIncomeSourceTypeErr

'********************************************************************************
'* Name: delIncomeSourceType
'* Description:
'* Created: 1/22/99 2:40:05 PM
'********************************************************************************
    
    'Hourglass True

    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = gcnDDS
    If gStoredProcs("up_d_Income_Source_Type").GetStoredProcCommand(cmd) = True Then
        cmd.Parameters("INCOME_SOURCE_TYPE_ID") = dblIncomeSourceTypeID
        cmd.Parameters("user_id") = gobjLoginInfo.UserId
        cmd.Execute
        If cmd.Parameters("RETURN_VALUE") <> 0 Then
            GetServerErrorMsg cmd.Parameters("RETURN_VALUE"), "Error occurred deleting the Income Source Type record."
        End If
    Else
        MsgBox "Error creating the Delete Income Source Type Stored Procedure.", vbCritical
        Set cmd = Nothing
        ExitApp
    End If
    Set cmd = Nothing

Xit:
    Hourglass False
    Exit Sub

delIncomeSourceTypeErr:
    ShowUnexpectedError MODULE + "delIncomeSourceType", Err
    Resume Xit

End Sub

Private Function DataValidation() As Boolean

On Error GoTo DataValidationErr

'********************************************************************************
'* Name: DataValidation
'* Description:
'* Created: 6/2/99 6:04:
'********************************************************************************
    Dim strEmptyFields As String
    Style = vbOKOnly + vbExclamation + vbApplicationModal
    strMessage = "The following data is required:    " & vbCrLf & vbCrLf
    strTitle = "Invalid Data"
    
    If Trim(txtFUNBCode.Text) = vbNullString Then
        strEmptyFields = " Code"
        txtFUNBCode.SetFocus
    End If
    
    If Trim(txtDescription) = vbNullString And strEmptyFields <> vbNullString Then
        strEmptyFields = strEmptyFields & vbCrLf & " Description"
    Else
        If Trim(txtDescription) = vbNullString Then
        strEmptyFields = strEmptyFields & " Description"
        txtDescription.SetFocus
        End If
    End If
    
    If IsNumeric(Trim(txtStart)) = False Then
        If strEmptyFields <> vbNullString Then
            strEmptyFields = strEmptyFields & vbCrLf & " DD Number Starting Position must be numeric"
        Else
            strEmptyFields = strEmptyFields & vbCrLf & " DD Number Starting Position must be numeric"
        End If
    End If
    
    If IsNumeric(Trim(txtLength)) = False Then
        If strEmptyFields <> vbNullString Then
            strEmptyFields = strEmptyFields & vbCrLf & " DD Number Length must be numeric"
        Else
            strEmptyFields = strEmptyFields & vbCrLf & " DD Number Length must be numeric"
        End If
    End If
    
    
    If strEmptyFields <> vbNullString Then
        strMessage = strMessage & strEmptyFields & vbCrLf
        MsgBox strMessage, Style, strTitle
        DataValidation = False
    Else
        DataValidation = True
    End If


Xit:
    Exit Function

DataValidationErr:
    ShowUnexpectedError MODULE + "DataValidation", Err
    Resume Xit


End Function





Private Sub OutlookTitle1_IconClick()
    If cmdCancel.Enabled = True Then
        Unload Me
    End If

End Sub

Private Sub txtFUNBCode_GotFocus()
'********************************************************************************
'* Name: txtFUNBCode_GotFocus()
'* Description: To focus on FUNB CODE
'* Created: 7/23/99 12:04 pm
'********************************************************************************
'
Call SetSelected

End Sub


Private Sub txtDescription_GotFocus()
'********************************************************************************
'* Name: txtDescription_GotFocus()
'* Description: To focus on Description Code
'* Created: 7/23/99 12:05 pm
'********************************************************************************
'
Call SetSelected

End Sub




Private Sub txtLength_GotFocus()
Call SetSelected

End Sub

Private Sub txtNCASAccount_GotFocus()

Call SetSelected

End Sub

Private Sub txtPACode_GotFocus()

Call SetSelected

End Sub


Private Sub txtPAPMTCode_GotFocus()
Call SetSelected

End Sub



Private Sub txtPAPMTREVCode_GotFocus()
Call SetSelected

End Sub


Private Sub txtPFDEPREVTRANSCode_GotFocus()
Call SetSelected

End Sub

Private Sub txtPFDEPTRANSCode_GotFocus()
Call SetSelected

End Sub

Private Sub txtStart_GotFocus()
Call SetSelected

End Sub
