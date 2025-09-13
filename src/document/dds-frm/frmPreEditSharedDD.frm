VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "SSDW3BO.OCX"
Begin VB.Form frmPreEditSharedDD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select a MRUN"
   ClientHeight    =   3240
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6000
   Icon            =   "frmPreEditSharedDD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid dbgSharedDDNum 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5760
      _Version        =   196616
      AllowUpdate     =   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowColumnSwapping=   0
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      ForeColorEven   =   0
      BackColorOdd    =   8454143
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   1773
      Columns(0).Caption=   "MRUN"
      Columns(0).Name =   "MRUN"
      Columns(0).DataField=   "MRUN"
      Columns(0).DataType=   3
      Columns(0).NumberFormat=   "0-00-00-00"
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   1614
      Columns(1).Caption=   "Institution"
      Columns(1).Name =   "Institution"
      Columns(1).DataField=   "Institution"
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   4313
      Columns(2).Caption=   "Name"
      Columns(2).Name =   "Name"
      Columns(2).DataField=   "Name"
      Columns(2).DataType=   17
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   1561
      Columns(3).Caption=   "Deceased"
      Columns(3).Name =   "Deceased"
      Columns(3).DataField=   "Deceased"
      Columns(3).DataType=   17
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(3).Style=   2
      _ExtentX        =   10160
      _ExtentY        =   3836
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4425
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Please select a MRUN"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmPreEditSharedDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private rsSharedDDNum As New ADODB.Recordset
Private sSharedDDNumSql As String
Private mdInvalidRecordID As Double
Private mnVBMsgBoxResult As VbMsgBoxResult
Private msMRUN As String
Private msInstitution As String
Private msName As String
Private mbDeceasedInd As Boolean

Public Property Get MRUN() As String
    MRUN = msMRUN
End Property

Public Property Get DeceasedInd() As Boolean
    DeceasedInd = mbDeceasedInd
End Property

Public Property Get Institution() As String
    Institution = msInstitution
End Property

Public Property Get PatientName() As String
    PatientName = msName
End Property

Public Property Get MsgBoxResult() As VbMsgBoxResult
    MsgBoxResult = mnVBMsgBoxResult
End Property

Public Property Let InvalidRecordID(vData As Double)
    mdInvalidRecordID = vData
End Property

Private Sub cmdCancel_Click()
    mnVBMsgBoxResult = vbCancel
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    mnVBMsgBoxResult = vbOK
    With rsSharedDDNum
        msMRUN = .Fields("MRUN")
        msInstitution = .Fields("Institution")
        msName = .Fields("Name")
        If .Fields("Deceased") = vbChecked Then
            mbDeceasedInd = True
        Else
            mbDeceasedInd = False
        End If
        
    End With
    Me.Hide
End Sub

Private Sub dbgSharedDDNum_Click()
    EnableUI
End Sub
Private Sub EnableUI()
    If dbgSharedDDNum.SelBookmarks.Count = 1 Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If

End Sub
Private Sub dbgSharedDDNum_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
    EnableUI
End Sub

Private Sub Form_Load()
    CenterMe Me
    Dim stempsql As String
    sSharedDDNumSql = "SELECT MEDICAL_RECORD_NUM AS MRUN, INSTITUTION_CODE AS 'Institution', PATIENT_NAME AS 'Name', CASE DECEASED_IND WHEN 'Y' THEN 1 ELSE 0 END AS Deceased FROM DD_SHARED_DD_NUM "
    stempsql = sSharedDDNumSql & " WHERE INVALID_RECORD_ID = " & mdInvalidRecordID
    RefreshGrid (stempsql)

End Sub

Private Sub RefreshGrid(ByRef sSql As String)
    Hourglass True
    With dbgSharedDDNum
        .Redraw = False
        Set .DataSource = Nothing
        If rsSharedDDNum.State = adStateOpen Then
            rsSharedDDNum.Close
        End If
        rsSharedDDNum.Open sSql, gcnDDS, adOpenForwardOnly, adLockOptimistic
        Set .DataSource = rsSharedDDNum
        .Refresh
        .Redraw = True
    End With
    Hourglass False
End Sub

