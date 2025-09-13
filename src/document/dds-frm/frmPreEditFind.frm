VERSION 5.00
Begin VB.Form frmPreEditFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   4200
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6135
   Icon            =   "frmPreEditFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFindFirst 
      Caption         =   "Find First"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.ListBox lstColumns 
      Height          =   2595
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txtSearchString 
      Height          =   375
      Left            =   120
      MaxLength       =   255
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Criteria"
      Height          =   2775
      Left            =   2880
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
      Begin VB.CheckBox chkMatchCase 
         Caption         =   "Match case"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox chkExactString 
         Caption         =   "Exact string"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      Caption         =   "In column:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Find what:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmPreEditFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************************************************
' * Description: Get user's search option and call the calling forms search
' *              method
' *
' *
' *
' *
' *
' * Revisions:
' *  8/9/99 yw Added comments.
' *
' *
' ********************************************************************************


' Mod CONSTANTS
Private Const MODULE As String = "Find (Pre-Edit)"

' Mod ENUMS


' Mod TYPES


' Mod DECLARES


' Mod VARIABLES


Option Explicit


Private Sub cmdClose_Click()
    'Close form w/o doing anything
    Unload Me
End Sub

Private Sub cmdFindFirst_Click()
    'Call find using findfirst option
    Find (True)
End Sub

Private Sub cmdFindNext_Click()
    'Call find using findnext option
    Find (False)
End Sub

Private Sub Find(ByVal bFindFirst As Boolean)

On Error GoTo FindErr

    Hourglass True
    'Call the frmPreEditMain's search method
    If frmPreEditMain.Search(txtSearchString.Text, lstColumns.Text, chkMatchCase.Value, chkExactString.Value, bFindFirst) = False Then
        'Not found
        Hourglass False
        Beep
        MsgBox "The specified search string cannot be found in the column.", vbExclamation
    Else
        'found
        Hourglass False
        gsSearchString = txtSearchString.Text
        Unload Me
        'If the fewer than 2 rows selected, select the current one
        If frmPreEditMain.dbgPreEdit.SelBookmarks.Count < 2 Then
            frmPreEditMain.SetRowSelected
        End If
    End If

Xit:
    Exit Sub

FindErr:
    ShowUnexpectedError MODULE + "Find", Err
    Resume Xit


End Sub


Private Sub Form_Activate()
    
    fMainForm.SetMainToolbar True

End Sub

Private Sub Form_Deactivate()
    
    fMainForm.SetMainToolbar False

End Sub

Private Sub Form_Load()
    CenterMe Me
    'Get the last search string
    txtSearchString.Text = gsSearchString
    'Initializing the list box
    InitList
End Sub

Private Sub InitList()

On Error GoTo InitListErr

'********************************************************************************
'* Name: InitList
'*
'* Description: Initialize the list box based upon the names of the columns
'*              of dbgPreEdit
'* Parameters:
'* Created: 8/9/99 11:21:24 AM
'********************************************************************************
    Dim Grid As SSOleDBGrid
    Dim i As Long
    
    Set Grid = frmPreEditMain.dbgPreEdit
    'Get the column's names and add them to the list box
    For i = 0 To Grid.Cols - 1
        lstColumns.AddItem Grid.Columns(i).Name
    Next i
    'Set the listindex according to the previous one
    lstColumns.ListIndex = glListIndex

Xit:
    Exit Sub

InitListErr:
    ShowUnexpectedError MODULE + "InitList", Err
    Resume Xit


End Sub

Private Sub lstColumns_Click()
    'Save the listbox index
    glListIndex = lstColumns.ListIndex
End Sub

Private Sub txtSearchString_Change()

On Error GoTo txtSearchString_ChangeErr

    'Enable buttons if there are something in the textbox
    EnableFindButtons IIf(Trim(txtSearchString.Text) <> vbNullString, True, False)

Xit:
    Exit Sub

txtSearchString_ChangeErr:
    ShowUnexpectedError MODULE + "txtSearchString_Change", Err
    Resume Xit


End Sub

Private Sub EnableFindButtons(ByVal bEnable As Boolean)
    cmdFindNext.Enabled = bEnable
    cmdFindFirst.Enabled = bEnable
End Sub
