VERSION 5.00
Begin VB.Form frmBrowserFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find & Replace"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFind 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Top             =   0
         Width           =   3615
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Default         =   -1  'True
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox chkCase 
         Caption         =   "Match case."
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Top             =   480
         Width           =   3615
      End
      Begin VB.CheckBox chkWord 
         Caption         =   "Whole words only."
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   3615
      End
      Begin VB.CommandButton cmdFindAll 
         Caption         =   "Find All"
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Search for phrase:"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmBrowserFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objRange As Object

Private Sub InitFind()
Set objRange = frmBrowser.webMain.Document.body.createTextRange
End Sub

Private Function GetFlags() As Integer
Dim intFlags As Integer

If chkCase.Value = 1 Then
    intFlags = intFlags + 4
End If

If chkWord.Value = 1 Then
    intFlags = intFlags + 2
End If

GetFlags = intFlags
End Function

Private Function Find(Word As String) As Boolean
On Error GoTo errHandle

If objRange.FindText(Word, 1, GetFlags) = True Then
    objRange.Select
    objRange.setEndPoint "StartToEnd", objRange
    Find = True
Else
    MsgBox "'" & Word & "' was no longer found in the document.", vbInformation, "Finished Searching"
    InitFind
    Find = False
End If

Exit Function

errHandle:
If Err.Number = 91 Then
    InitFind
    Find Word
End If
End Function

Private Sub FindAll(Word As String)
InitFind

Do While objRange.FindText(Word) = True
    If Find(Word) = True Then
        frmBrowser.webMain.Document.execCommand "Bold"
        frmBrowser.webMain.Document.Selection.createRange.parentElement.Style.backgroundColor = "yellow"
    End If
Loop

frmBrowser.webMain.Document.Selection.empty
End Sub

Private Sub cmdFind_Click()
If txtFind.Text <> "" Then Find txtFind.Text
End Sub

Private Sub cmdFindAll_Click()
If txtFind.Text <> "" Then FindAll txtFind.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objRange = Nothing
End Sub
