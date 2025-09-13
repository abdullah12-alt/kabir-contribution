VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{8CD222DF-7752-11D3-9D1E-00105A19BCF2}#1.0#0"; "OAOTBar.ocx"
Begin VB.Form frmPreEditMain 
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   9990
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab sstPreEdit 
      Height          =   5895
      Left            =   120
      TabIndex        =   1
      Top             =   690
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   917
      TabCaption(0)   =   "Edit/Move/Hide"
      TabPicture(0)   =   "frmPreEditMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPreEditRecs"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dbgPreEdit"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSelectMRUN"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdRecoup"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDelete"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdMove"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdEdit"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdResolveMulti"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdViewDetails"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdOverride"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Unhide"
      TabPicture(1)   =   "frmPreEditMain.frx":0452
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "dbgDeletedRecords"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdPurge"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdUndelete"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdOverride 
         Caption         =   "&Override"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         TabIndex        =   24
         Top             =   5280
         Width           =   1155
      End
      Begin VB.CommandButton cmdViewDetails 
         Caption         =   "View Details"
         Height          =   375
         Left            =   135
         TabIndex        =   23
         Top             =   5280
         Width           =   1080
      End
      Begin VB.CommandButton cmdResolveMulti 
         Caption         =   "Auto Resolve"
         Height          =   375
         Left            =   1260
         TabIndex        =   22
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Errors"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74760
         TabIndex        =   19
         Top             =   3600
         Width           =   9210
         Begin VB.TextBox txtDeletedErrors 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   240
            Width           =   8895
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Errors"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   240
         TabIndex        =   17
         Top             =   3885
         Width           =   9210
         Begin VB.TextBox txtErrors 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   285
            Width           =   8895
         End
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   7185
         TabIndex        =   16
         Top             =   5280
         Width           =   1155
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "&Move"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5985
         TabIndex        =   4
         Top             =   5280
         Width           =   1155
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Hide"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8385
         TabIndex        =   5
         Top             =   5280
         Width           =   1155
      End
      Begin VB.CommandButton cmdUndelete 
         Caption         =   "&Unhide"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66720
         TabIndex        =   7
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton cmdPurge 
         Caption         =   "&Purge"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68040
         TabIndex        =   13
         Top             =   5280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdRecoup 
         Caption         =   "&Recoup"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3615
         TabIndex        =   3
         Top             =   5280
         Width           =   1155
      End
      Begin VB.CommandButton cmdSelectMRUN 
         Caption         =   "&Select MRUN"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2415
         TabIndex        =   2
         Top             =   5280
         Width           =   1155
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid dbgPreEdit 
         Height          =   2790
         Left            =   240
         TabIndex        =   0
         Top             =   705
         Width           =   9210
         _Version        =   196616
         DataMode        =   1
         UseGroups       =   -1  'True
         AllowUpdate     =   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeRow   =   3
         SelectByCell    =   -1  'True
         MaxSelectedRows =   25
         ForeColorEven   =   0
         BackColorOdd    =   8454143
         Levels          =   2
         RowHeight       =   847
         Groups(0).Width =   15214
         Groups(0).Caption=   "Invalid Records"
         Groups(0).Columns.Count=   14
         Groups(0).Columns(0).Width=   1984
         Groups(0).Columns(0).Caption=   "DD #"
         Groups(0).Columns(0).Name=   "DD #"
         Groups(0).Columns(0).DataField=   "Column 0"
         Groups(0).Columns(0).FieldLen=   11
         Groups(0).Columns(0).Locked=   -1  'True
         Groups(0).Columns(0).HasHeadForeColor=   -1  'True
         Groups(0).Columns(1).Width=   2196
         Groups(0).Columns(1).Caption=   "Income Source"
         Groups(0).Columns(1).Name=   "Income Source"
         Groups(0).Columns(1).DataField=   "Column 1"
         Groups(0).Columns(1).FieldLen=   10
         Groups(0).Columns(1).Locked=   -1  'True
         Groups(0).Columns(1).HasHeadForeColor=   -1  'True
         Groups(0).Columns(2).Width=   2487
         Groups(0).Columns(2).Caption=   "Name"
         Groups(0).Columns(2).Name=   "Name"
         Groups(0).Columns(2).DataField=   "Column 2"
         Groups(0).Columns(2).FieldLen=   256
         Groups(0).Columns(2).Locked=   -1  'True
         Groups(0).Columns(3).Width=   1640
         Groups(0).Columns(3).Caption=   "MRUN"
         Groups(0).Columns(3).Name=   "MRUN"
         Groups(0).Columns(3).DataField=   "Column 3"
         Groups(0).Columns(3).DataType=   3
         Groups(0).Columns(3).NumberFormat=   "0-00-00-00"
         Groups(0).Columns(3).FieldLen=   256
         Groups(0).Columns(3).Locked=   -1  'True
         Groups(0).Columns(3).HasHeadForeColor=   -1  'True
         Groups(0).Columns(4).Width=   2328
         Groups(0).Columns(4).Caption=   "FUNB As Of Date"
         Groups(0).Columns(4).Name=   "FUNB As Of Date"
         Groups(0).Columns(4).DataField=   "Column 4"
         Groups(0).Columns(4).DataType=   7
         Groups(0).Columns(4).NumberFormat=   "MM/DD/YYYY"
         Groups(0).Columns(4).FieldLen=   256
         Groups(0).Columns(4).Locked=   -1  'True
         Groups(0).Columns(5).Width=   1931
         Groups(0).Columns(5).Caption=   "Created Date"
         Groups(0).Columns(5).Name=   "Created Date"
         Groups(0).Columns(5).DataField=   "Column 5"
         Groups(0).Columns(5).DataType=   7
         Groups(0).Columns(5).NumberFormat=   "MM/DD/YYYY"
         Groups(0).Columns(5).FieldLen=   256
         Groups(0).Columns(5).Locked=   -1  'True
         Groups(0).Columns(6).Width=   2646
         Groups(0).Columns(6).Caption=   "FUNB Amount"
         Groups(0).Columns(6).Name=   "FUNB Amount"
         Groups(0).Columns(6).Alignment=   1
         Groups(0).Columns(6).DataField=   "Column 6"
         Groups(0).Columns(6).DataType=   6
         Groups(0).Columns(6).NumberFormat=   "$#,##0.00;($#,##0.00)"
         Groups(0).Columns(6).FieldLen=   256
         Groups(0).Columns(6).Locked=   -1  'True
         Groups(0).Columns(7).Width=   4683
         Groups(0).Columns(7).Caption=   "Memo"
         Groups(0).Columns(7).Name=   "Memo"
         Groups(0).Columns(7).DataField=   "Column 7"
         Groups(0).Columns(7).Level=   1
         Groups(0).Columns(7).FieldLen=   255
         Groups(0).Columns(7).Locked=   -1  'True
         Groups(0).Columns(7).HasHeadForeColor=   -1  'True
         Groups(0).Columns(8).Width=   1984
         Groups(0).Columns(8).Caption=   "Account #"
         Groups(0).Columns(8).Name=   "Account #"
         Groups(0).Columns(8).DataField=   "Column 8"
         Groups(0).Columns(8).DataType=   5
         Groups(0).Columns(8).Level=   1
         Groups(0).Columns(8).NumberFormat=   "00000000"
         Groups(0).Columns(8).FieldLen=   256
         Groups(0).Columns(8).Locked=   -1  'True
         Groups(0).Columns(9).Width=   1640
         Groups(0).Columns(9).Caption=   "Institution"
         Groups(0).Columns(9).Name=   "Institution"
         Groups(0).Columns(9).DataField=   "Column 9"
         Groups(0).Columns(9).Level=   1
         Groups(0).Columns(9).FieldLen=   256
         Groups(0).Columns(9).Locked=   -1  'True
         Groups(0).Columns(10).Width=   1402
         Groups(0).Columns(10).Caption=   "Type"
         Groups(0).Columns(10).Name=   "Type"
         Groups(0).Columns(10).DataField=   "Column 10"
         Groups(0).Columns(10).Level=   1
         Groups(0).Columns(10).FieldLen=   256
         Groups(0).Columns(10).Locked=   -1  'True
         Groups(0).Columns(11).Width=   1561
         Groups(0).Columns(11).Caption=   "Deceased"
         Groups(0).Columns(11).Name=   "Deceased"
         Groups(0).Columns(11).DataField=   "Column 11"
         Groups(0).Columns(11).Level=   1
         Groups(0).Columns(11).FieldLen=   256
         Groups(0).Columns(11).Locked=   -1  'True
         Groups(0).Columns(11).Style=   2
         Groups(0).Columns(12).Width=   1826
         Groups(0).Columns(12).Caption=   "Posting Error"
         Groups(0).Columns(12).Name=   "Posting Error"
         Groups(0).Columns(12).DataField=   "Column 12"
         Groups(0).Columns(12).Level=   1
         Groups(0).Columns(12).FieldLen=   256
         Groups(0).Columns(12).Locked=   -1  'True
         Groups(0).Columns(12).Style=   2
         Groups(0).Columns(13).Width=   2117
         Groups(0).Columns(13).Caption=   "Multi MRUN"
         Groups(0).Columns(13).Name=   "Multi MRUN"
         Groups(0).Columns(13).DataField=   "Column 13"
         Groups(0).Columns(13).Level=   1
         Groups(0).Columns(13).FieldLen=   256
         Groups(0).Columns(13).Locked=   -1  'True
         Groups(0).Columns(13).Style=   2
         _ExtentX        =   16245
         _ExtentY        =   4921
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid dbgDeletedRecords 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   6
         Top             =   720
         Width           =   9210
         _Version        =   196616
         DataMode        =   1
         UseGroups       =   -1  'True
         AllowUpdate     =   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowColumnShrinking=   0   'False
         SelectTypeRow   =   1
         SelectByCell    =   -1  'True
         MaxSelectedRows =   2
         ForeColorEven   =   0
         BackColorOdd    =   12648384
         Levels          =   2
         RowHeight       =   847
         Groups(0).Width =   15214
         Groups(0).Caption=   "Deleted Invalid Records"
         Groups(0).Columns.Count=   14
         Groups(0).Columns(0).Width=   1826
         Groups(0).Columns(0).Caption=   "DD #"
         Groups(0).Columns(0).Name=   "DD #"
         Groups(0).Columns(0).DataField=   "Column 0"
         Groups(0).Columns(0).FieldLen=   11
         Groups(0).Columns(0).Locked=   -1  'True
         Groups(0).Columns(0).HasHeadForeColor=   -1  'True
         Groups(0).Columns(1).Width=   2249
         Groups(0).Columns(1).Caption=   "Income Source"
         Groups(0).Columns(1).Name=   "Income Source"
         Groups(0).Columns(1).DataField=   "Column 1"
         Groups(0).Columns(1).FieldLen=   10
         Groups(0).Columns(1).Locked=   -1  'True
         Groups(0).Columns(1).HasHeadForeColor=   -1  'True
         Groups(0).Columns(2).Width=   2593
         Groups(0).Columns(2).Caption=   "Name"
         Groups(0).Columns(2).Name=   "Name"
         Groups(0).Columns(2).DataField=   "Column 2"
         Groups(0).Columns(2).FieldLen=   256
         Groups(0).Columns(2).Locked=   -1  'True
         Groups(0).Columns(3).Width=   1588
         Groups(0).Columns(3).Caption=   "MRUN"
         Groups(0).Columns(3).Name=   "MRUN"
         Groups(0).Columns(3).DataField=   "Column 3"
         Groups(0).Columns(3).DataType=   3
         Groups(0).Columns(3).NumberFormat=   "0-00-00-00"
         Groups(0).Columns(3).FieldLen=   256
         Groups(0).Columns(3).Locked=   -1  'True
         Groups(0).Columns(3).HasHeadForeColor=   -1  'True
         Groups(0).Columns(4).Width=   2408
         Groups(0).Columns(4).Caption=   "FUNB As Of Date"
         Groups(0).Columns(4).Name=   "FUNB As Of Date"
         Groups(0).Columns(4).DataField=   "Column 4"
         Groups(0).Columns(4).DataType=   7
         Groups(0).Columns(4).NumberFormat=   "MM/DD/YYYY"
         Groups(0).Columns(4).FieldLen=   256
         Groups(0).Columns(4).Locked=   -1  'True
         Groups(0).Columns(5).Width=   1931
         Groups(0).Columns(5).Caption=   "Created Date"
         Groups(0).Columns(5).Name=   "Created Date"
         Groups(0).Columns(5).DataField=   "Column 5"
         Groups(0).Columns(5).DataType=   7
         Groups(0).Columns(5).NumberFormat=   "MM/DD/YYYY"
         Groups(0).Columns(5).FieldLen=   256
         Groups(0).Columns(5).Locked=   -1  'True
         Groups(0).Columns(6).Width=   2619
         Groups(0).Columns(6).Caption=   "FUNB Amount"
         Groups(0).Columns(6).Name=   "FUNB Amount"
         Groups(0).Columns(6).Alignment=   1
         Groups(0).Columns(6).DataField=   "Column 6"
         Groups(0).Columns(6).DataType=   6
         Groups(0).Columns(6).NumberFormat=   "$#,##0.00;($#,##0.00)"
         Groups(0).Columns(6).FieldLen=   256
         Groups(0).Columns(6).Locked=   -1  'True
         Groups(0).Columns(7).Width=   4683
         Groups(0).Columns(7).Caption=   "Memo"
         Groups(0).Columns(7).Name=   "Memo"
         Groups(0).Columns(7).DataField=   "Column 7"
         Groups(0).Columns(7).Level=   1
         Groups(0).Columns(7).FieldLen=   255
         Groups(0).Columns(7).Locked=   -1  'True
         Groups(0).Columns(7).HasHeadForeColor=   -1  'True
         Groups(0).Columns(8).Width=   1984
         Groups(0).Columns(8).Caption=   "Account #"
         Groups(0).Columns(8).Name=   "Account #"
         Groups(0).Columns(8).DataField=   "Column 8"
         Groups(0).Columns(8).DataType=   5
         Groups(0).Columns(8).Level=   1
         Groups(0).Columns(8).NumberFormat=   "00000000"
         Groups(0).Columns(8).FieldLen=   256
         Groups(0).Columns(8).Locked=   -1  'True
         Groups(0).Columns(9).Width=   1588
         Groups(0).Columns(9).Caption=   "Institution"
         Groups(0).Columns(9).Name=   "Institution"
         Groups(0).Columns(9).DataField=   "Column 9"
         Groups(0).Columns(9).Level=   1
         Groups(0).Columns(9).FieldLen=   256
         Groups(0).Columns(9).Locked=   -1  'True
         Groups(0).Columns(10).Width=   1402
         Groups(0).Columns(10).Caption=   "Type"
         Groups(0).Columns(10).Name=   "Type"
         Groups(0).Columns(10).DataField=   "Column 10"
         Groups(0).Columns(10).Level=   1
         Groups(0).Columns(10).FieldLen=   256
         Groups(0).Columns(10).Locked=   -1  'True
         Groups(0).Columns(11).Width=   1561
         Groups(0).Columns(11).Caption=   "Deceased"
         Groups(0).Columns(11).Name=   "Deceased"
         Groups(0).Columns(11).DataField=   "Column 11"
         Groups(0).Columns(11).Level=   1
         Groups(0).Columns(11).FieldLen=   256
         Groups(0).Columns(11).Locked=   -1  'True
         Groups(0).Columns(11).Style=   2
         Groups(0).Columns(12).Width=   1879
         Groups(0).Columns(12).Caption=   "Posting Error"
         Groups(0).Columns(12).Name=   "Posting Error"
         Groups(0).Columns(12).DataField=   "Column 12"
         Groups(0).Columns(12).Level=   1
         Groups(0).Columns(12).FieldLen=   256
         Groups(0).Columns(12).Locked=   -1  'True
         Groups(0).Columns(12).Style=   2
         Groups(0).Columns(13).Width=   2117
         Groups(0).Columns(13).Caption=   "Multi MRUN"
         Groups(0).Columns(13).Name=   "Multi MRUN"
         Groups(0).Columns(13).DataField=   "Column 13"
         Groups(0).Columns(13).Level=   1
         Groups(0).Columns(13).FieldLen=   256
         Groups(0).Columns(13).Locked=   -1  'True
         Groups(0).Columns(13).Style=   2
         _ExtentX        =   16245
         _ExtentY        =   4895
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
      Begin VB.Label lblPreEditRecs 
         Caption         =   "Pre-Edit Records: "
         Height          =   270
         Left            =   300
         TabIndex        =   21
         Top             =   3570
         Width           =   3225
      End
      Begin VB.Label Label2 
         Caption         =   "Select a record for purging or unhiding."
         Height          =   375
         Left            =   -74760
         TabIndex        =   14
         Top             =   5280
         Width           =   4695
      End
   End
   Begin OAOTitleBar.OutlookTitleBar outTitle 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   979
      ForeColor       =   16777215
      Caption         =   "Pre-Edit File"
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "&Validate"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find..."
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeselectAll 
      Caption         =   "&Deselect All"
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   12
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Select All"
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   11
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      BorderStyle     =   6  'Inside Solid
      X1              =   120
      X2              =   9720
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   9720
      Y1              =   6720
      Y2              =   6720
   End
End
Attribute VB_Name = "frmPreEditMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ********************************************************************************
' * Description:
' *
' *
' *
' *
' *
' *
' * Revisions:
' *  9/17/99    yw Added comments.
' *  10/6/99    yw Change Delete to Hide
' *
' ********************************************************************************


' Mod CONSTANTS
Private Const MODULE = "Pre-Edit File"

' Mod ENUMS
Private Enum enPreEditMode
    MoveEditDelete
    Undelete
End Enum

' Mod TYPES
Private Type udtSortedColumnFlag
    ColIndex As Integer
    Ascending As Boolean
End Type


' Mod DECLARES


' Mod VARIABLES



Private sBaseSql As String
Private sPreEditSql As String
Private sSortedSql As String
Private sDeletedRecordsSql As String
Private sSourceTypeSql As String
Private sValidationErrorsSql As String
Private PreEditMode As enPreEditMode
Private cmdPreEdit As New ADODB.Command
Private rsPreEdit As New ADODB.Recordset
Private cmdDeletedRecords As New ADODB.Command
Private rsDeletedRecords As New ADODB.Recordset
Private rsSourceType As New ADODB.Recordset
Private rsValidationErrors As New ADODB.Recordset



Private Sub EnablePreEditUI()

On Error GoTo EnablePreEditUIErr

'********************************************************************************
'* Name: EnablePreEditUI
'*
'* Description: Enable/disable buttons according to user's input
'* Parameters:
'* Created: 9/17/99 10:13:40 AM
'********************************************************************************
    Dim bEnableDelete As Boolean
    Dim bEnableMove As Boolean
    Dim bEnableRecoup As Boolean
    Dim bEnableOverride As Boolean
    Dim bEnableSelectMRUN As Boolean
    Dim bEnableEdit As Boolean
    Dim iCount As Integer
    iCount = dbgPreEdit.SelBookmarks.Count
    
    'One selection: check for Move, MultiMRUN
    If iCount = 1 Then
        'Enable delete button
        bEnableEdit = True
        bEnableDelete = True
        bEnableOverride = True
        'Enable Edit button
        Dim bmk As Variant
        'Save the bookmark and disable the redraw
'        bmk = rsPreEdit.Bookmark
        dbgPreEdit.Redraw = False
        rsPreEdit.Move 0, dbgPreEdit.SelBookmarks(0)
        
        'Check for Deceased
        If rsPreEdit("Deceased") = 1 And rsPreEdit("Type") = "CR" Then
            bEnableMove = True
        End If
        'Check for Multi MRUN
        If rsPreEdit("Multi MRUN") = 1 Then
            bEnableSelectMRUN = True
        End If
        'Move the recordset to the saved position
'        rsPreEdit.Move 0, bmk
        dbgPreEdit.Redraw = True
    ElseIf iCount = 2 Then
        'Check recoup
        bEnableRecoup = CheckRecoup
    End If
    'If there is no records, disable the Find button
    If dbgPreEdit.Rows = 0 Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
    
    'Disable/Enable buttons accordingly
    cmdMove.Enabled = bEnableMove
    cmdDelete.Enabled = bEnableDelete
    cmdRecoup.Enabled = bEnableRecoup
    cmdSelectMRUN.Enabled = bEnableSelectMRUN
    cmdEdit.Enabled = bEnableEdit
    cmdOverride.Enabled = bEnableOverride

Xit:
    Exit Sub

EnablePreEditUIErr:
    ShowUnexpectedError MODULE + " EnablePreEditUI", Err
    Resume Xit


End Sub


Private Sub EnableDeletedUI()

On Error GoTo EnableDeletedUIErr

'********************************************************************************
'* Name: EnableDeletedUI
'*
'* Description:
'* Parameters:
'* Created: 9/17/99 10:34:22 AM
'********************************************************************************
    Dim bEnable As Boolean
    If dbgDeletedRecords.SelBookmarks.Count > 0 Then
        bEnable = True
    End If
    If dbgDeletedRecords.Rows = 0 Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
    cmdPurge.Enabled = bEnable
    cmdUndelete.Enabled = bEnable


Xit:
    Exit Sub

EnableDeletedUIErr:
    ShowUnexpectedError MODULE + " EnableDeletedUI", Err
    Resume Xit


End Sub


Private Sub cmdClose_Click()
    Unload Me

End Sub



Private Sub cmdDelete_Click()

On Error GoTo cmdDelete_ClickErr

    If MsgBox("Are you sure you want to hide this record?", vbYesNo + vbExclamation) = vbYes Then
        DeleteInvalidRecords
        RefreshPreEditGrid (sSortedSql)
'        RefreshDeletedRecordsGrid (sDeletedRecordsSql)
        'don't forget to call this after refresh the grids
        EnablePreEditUI
        EnableDeletedUI
    End If

Xit:
    Exit Sub

cmdDelete_ClickErr:
    ShowUnexpectedError MODULE + " cmdDelete_Click", Err
    Resume Xit


End Sub


Private Sub cmdEdit_Click()
    Dim fPreEditEdit As frmPreEditEdit
    Dim oPatientInfo As New clsPatientPreEditInfo
    Dim iX As Integer
On Error GoTo cmdEdit_ClickErr

    
    Set fPreEditEdit = New frmPreEditEdit
    rsPreEdit.Move 0, dbgPreEdit.SelBookmarks(0)
    With oPatientInfo
        .AccountNum = IIf(IsNull(rsPreEdit.Fields("Account #")), "", rsPreEdit.Fields("Account #"))
        .DDNum = IIf(IsNull(rsPreEdit.Fields("DD #")), "", rsPreEdit.Fields("DD #"))
        .MRUN = IIf(IsNull(rsPreEdit.Fields("MRUN")), "", rsPreEdit.Fields("MRUN"))
        .IncomeSource = IIf(IsNull(rsPreEdit.Fields("Income Source")), "", rsPreEdit.Fields("Income Source"))
        .InstitutionCode = IIf(IsNull(rsPreEdit.Fields("Institution")), "", rsPreEdit.Fields("Institution"))
        .InvalidRecordID = rsPreEdit.Fields("InvalidRecordID")
        .Memo = IIf(IsNull(rsPreEdit.Fields("Memo")), "", rsPreEdit.Fields("Memo"))
        .Name = IIf(IsNull(rsPreEdit.Fields("Name")), "", rsPreEdit.Fields("Name"))
    End With
    fPreEditEdit.setPatientPreEditInfo oPatientInfo
    fPreEditEdit.Show vbModal
    If fPreEditEdit.MsgBoxResult = vbOK Then
        Hourglass True
        Dim lReturnValue As Long
        Dim dInvalidRecordID As Double
        'save the invalid record id
        dInvalidRecordID = rsPreEdit.Fields("InvalidRecordID")
        'update the database
        Dim cmdUpdate As New ADODB.Command
        Set cmdUpdate.ActiveConnection = gcnDDS
        If gStoredProcs("up_u_PreEdit").GetStoredProcCommand(cmdUpdate) = True Then
            For iX = 0 To cmdUpdate.Parameters.Count - 1
                cmdUpdate.Parameters(iX) = Null
            Next iX

            cmdUpdate.Parameters("dInvalidRecordID") = dInvalidRecordID
            cmdUpdate.Parameters("sDDNum") = oPatientInfo.DDNum
            cmdUpdate.Parameters("sIncomeSource") = oPatientInfo.IncomeSource
            cmdUpdate.Parameters("sMemo") = oPatientInfo.Memo
            cmdUpdate.Parameters("sUserID") = gobjLoginInfo.UserId
            cmdUpdate.Execute
            lReturnValue = cmdUpdate.Parameters("RETURN_VALUE").value
            If lReturnValue <> 0 Then
                GetServerErrorMsg lReturnValue, "Error occurs when updating invalid record (ID: " & dInvalidRecordID & " ) " _
                & "error message from server follows: "
            Else
                'refresh the database
                RefreshPreEditGrid (sSortedSql)
                'don't forget to call this after refresh the grids
                EnablePreEditUI
                'search the recordset using the previously saved invalid record id
                dbgPreEdit.Redraw = False
                rsPreEdit.Find "InvalidRecordID = " & CStr(dInvalidRecordID)
                dbgPreEdit.Redraw = True
                Hourglass False
            End If
        End If
        Set cmdUpdate = Nothing
    End If
    Unload fPreEditEdit
    Set fPreEditEdit = Nothing
    Set oPatientInfo = Nothing

Xit:
    Exit Sub

cmdEdit_ClickErr:
    ShowUnexpectedError MODULE + " cmdEdit_Click", Err
    Resume Xit


End Sub

Private Sub cmdFind_Click()
    Dim fFindForm As New frmPreEditFind

On Error GoTo cmdFind_ClickErr

    fFindForm.Show vbModal
    Set fFindForm = Nothing
    

Xit:
    Exit Sub

cmdFind_ClickErr:
    ShowUnexpectedError MODULE + " cmdFind_Click", Err
    Resume Xit


End Sub

Public Sub SetRowSelected()
'********************************************************************************
'* Name: SetRowSelected
'*
'* Description: Add the row to the selected collection and enable/disable buttons
'* Parameters:
'* Created: 9/17/99 10:39:44 AM
'********************************************************************************
    If PreEditMode = MoveEditDelete Then
        dbgPreEdit.SelBookmarks.Add dbgPreEdit.Bookmark
        EnablePreEditUI
    ElseIf PreEditMode = Undelete Then
        dbgDeletedRecords.SelBookmarks.Add dbgDeletedRecords.Bookmark
        EnableDeletedUI
    End If
End Sub

Private Sub cmdMove_Click()

On Error GoTo cmdMove_ClickErr

    Beep
    If MsgBox("Are you sure you want to move this record for posting?", vbYesNo + vbExclamation) = vbYes Then
        If MoveInvalidRecords = True Then
            RefreshPreEditGrid (sSortedSql)
            EnablePreEditUI
        End If
    End If

Xit:
    Exit Sub

cmdMove_ClickErr:
    ShowUnexpectedError MODULE + " cmdMove_Click", Err
    Resume Xit


End Sub
Private Function MoveInvalidRecords() As Boolean
    Dim cmdMove As New ADODB.Command
    Dim lInvalidRecordID As Long
    Dim i As Long
    Dim lReturnValue As Long
    Dim iX As Integer
On Error GoTo MoveInvalidRecordsErr

    
    MoveInvalidRecords = False
    Set cmdMove.ActiveConnection = gcnDDS
    If gStoredProcs("up_id_Move_Invalid_Rec").GetStoredProcCommand(cmdMove) = True Then
        
        'The SelBookmarks should have only one selection after the design change
        For i = 0 To dbgPreEdit.SelBookmarks.Count - 1
            rsPreEdit.Move 0, (dbgPreEdit.SelBookmarks(i))
            lInvalidRecordID = rsPreEdit.Fields("InvalidRecordID")
            For iX = 0 To cmdMove.Parameters.Count - 1
                cmdMove.Parameters(iX) = Null
            Next iX
            
            cmdMove.Parameters("dInvalidRecordID") = lInvalidRecordID
            cmdMove.Parameters("sPostingMode") = "Deposit"
            cmdMove.Parameters("sUserID") = gobjLoginInfo.UserId
            cmdMove.Execute
            lReturnValue = cmdMove.Parameters("RETURN_VALUE").value
            If lReturnValue <> 0 Then
                GetServerErrorMsg lReturnValue, "Error occurs when processing invalid record (ID: " & lInvalidRecordID & " ) " _
                & "error message from server follows: " & vbCrLf
            Else
                MoveInvalidRecords = True
            End If
            
        Next i
    End If
    MoveInvalidRecords = True

Xit:
    Exit Function

MoveInvalidRecordsErr:
    ShowUnexpectedError MODULE + " MoveInvalidRecords", Err
    Resume Xit


End Function
Private Sub DeleteInvalidRecords()
    Dim i As Long
    Dim bkmrk As Variant
    Dim cmdDelete As New ADODB.Command
    Dim dInvalidRecordID As Double
    Dim lReturnValue As Long
    Dim iX As Integer
On Error GoTo DeleteInvalidRecordsErr

    Set cmdDelete.ActiveConnection = gcnDDS
    Hourglass True
    dbgPreEdit.Redraw = False
    For i = 0 To dbgPreEdit.SelBookmarks.Count - 1
        'get the bookmark for the selbookmarks collection
        'and use it to move to that record in the recordset
        rsPreEdit.Move 0, (dbgPreEdit.SelBookmarks(i))
        If gStoredProcs("up_u_Hide_PreEdit").GetStoredProcCommand(cmdDelete) = True Then
            'The SelBookmarks should have only one selection after the design change
            dInvalidRecordID = rsPreEdit.Fields("InvalidRecordID")
            For iX = 0 To cmdDelete.Parameters.Count - 1
                cmdDelete.Parameters(iX) = Null
            Next iX
            
            cmdDelete.Parameters("dInvalidRecordID") = dInvalidRecordID
            cmdDelete.Parameters("sRecordStatus") = "I"
            cmdDelete.Parameters("sUserID") = gobjLoginInfo.UserId
            cmdDelete.Execute
            lReturnValue = cmdDelete.Parameters("RETURN_VALUE").value
            If lReturnValue <> 0 Then
                GetServerErrorMsg lReturnValue, "Error occurs when processing invalid record (ID: " & dInvalidRecordID & " ) " _
                & "error message from server follows: " & vbCrLf
            End If
        End If
    Next i
    dbgPreEdit.Redraw = True
    Hourglass False
    Set cmdDelete = Nothing


Xit:
    Exit Sub

DeleteInvalidRecordsErr:
    ShowUnexpectedError MODULE + " DeleteInvalidRecords", Err
    Resume Xit


End Sub

Private Sub UndeleteInvalidRecords()
    Dim i As Long
    Dim bkmrk As Variant
    Dim cmdUndelete As New ADODB.Command
    Dim dInvalidRecordID As Double
    Dim lReturnValue As Long
    Dim iX As Integer
On Error GoTo UndeleteInvalidRecordsErr

    Set cmdUndelete.ActiveConnection = gcnDDS
    Hourglass True
    dbgDeletedRecords.Redraw = False
    For i = 0 To dbgDeletedRecords.SelBookmarks.Count - 1
        'get the bookmark for the selbookmarks collection
        'and use it to move to that record in the recordset
        rsDeletedRecords.Move 0, (dbgDeletedRecords.SelBookmarks(i))
        If gStoredProcs("up_u_Hide_PreEdit").GetStoredProcCommand(cmdUndelete) = True Then
            'The SelBookmarks should have only one selection after the design change
            dInvalidRecordID = rsDeletedRecords.Fields("InvalidRecordID")
            For iX = 0 To cmdUndelete.Parameters.Count - 1
               cmdUndelete.Parameters(iX) = Null
            Next iX
            cmdUndelete.Parameters("dInvalidRecordID") = dInvalidRecordID
            cmdUndelete.Parameters("sRecordStatus") = "A"
            cmdUndelete.Parameters("sUserID") = gobjLoginInfo.UserId
            cmdUndelete.Execute
            lReturnValue = cmdUndelete.Parameters("RETURN_VALUE").value
            If lReturnValue <> 0 Then
                GetServerErrorMsg lReturnValue, "Error occurs when processing invalid record (ID: " & dInvalidRecordID & ") error message from server follows: " & vbCrLf
            End If
        End If
    Next i
    dbgDeletedRecords.Redraw = True
    Hourglass False
    Set cmdUndelete = Nothing


Xit:
    Exit Sub

UndeleteInvalidRecordsErr:
    ShowUnexpectedError MODULE + " UndeleteInvalidRecords", Err
    Resume Xit


End Sub

Private Sub PurgeInvalidRecords()
    Dim i As Long
    Dim bkmrk As Variant
    Dim cmdPurge As New ADODB.Command
    Dim sPurgeSql As String
    Dim lInvalidRecordID As Long
    Dim lReturnValue As Long

On Error GoTo PurgeInvalidRecordsErr

    Set cmdPurge.ActiveConnection = gcnDDS
    Hourglass True
    If gStoredProcs("up_d_Invalid_Rec").GetStoredProcCommand(cmdPurge) = True Then
        For i = 0 To dbgDeletedRecords.SelBookmarks.Count - 1
            'get the bookmark for the selbookmarks collection
            'and use it to move to that record in the recordset
            rsDeletedRecords.Move 0, (dbgDeletedRecords.SelBookmarks(i))
            lInvalidRecordID = rsDeletedRecords.Fields("InvalidRecordID")
            cmdPurge.Parameters("dInvalidRecordID") = lInvalidRecordID
            cmdPurge.Execute
            lReturnValue = cmdPurge.Parameters("RETURN_VALUE").value
            If lReturnValue <> 0 Then
                GetServerErrorMsg lReturnValue, "Error occurs when processing invalid record (ID: " & lInvalidRecordID & " ) " _
                & "error message from server follows: " & vbCrLf
            End If
        Next i
    End If
    Hourglass False
    Set cmdPurge = Nothing


Xit:
    Exit Sub

PurgeInvalidRecordsErr:
    ShowUnexpectedError MODULE + " PurgeInvalidRecords", Err
    Resume Xit


End Sub

Private Sub cmdOverride_Click()

If dbgPreEdit.SelBookmarks.Count <> 1 Then
    MsgBox "You can only override one record at a time", vbInformation
    Exit Sub
End If
    
Load frmOverride
    
'Move to the record you want to overrride
rsPreEdit.Move 0, (dbgPreEdit.SelBookmarks(0))

'Fill the values of the Override Screen
With frmOverride
.lblFields(14).Caption = rsPreEdit.Fields("InvalidRecordId")

End With

frmOverride.Show vbModal
RefreshPreEditGrid (sSortedSql)

End Sub



Private Sub cmdPurge_Click()

On Error GoTo cmdPurge_ClickErr

    Beep
    If MsgBox("Are you sure you want to purge these records?", vbYesNo + vbExclamation) = vbYes Then
        PurgeInvalidRecords
'        RefreshDeletedRecordsGrid (sDeletedRecordsSql)
        'don't forget to call this after refresh the grids
        EnableDeletedUI
    End If

Xit:
    Exit Sub

cmdPurge_ClickErr:
    ShowUnexpectedError MODULE + " cmdPurge_Click", Err
    Resume Xit


End Sub

Private Sub cmdRecoup_Click()
'*********************************************************
'*
'* This function is used to recoup 1 DR to one or many CR
'*********************************************************
    
    Dim cmd As New ADODB.Command
    Dim dCRInvalidRecordID(6) As Double
    Dim dDRInvalidRecordID As Double
    Dim lReturnValue As Long
    Dim dInvalidRecordID As Double
    Dim iDaysInHouse As Integer
    Dim sATP_PML_FLAG As String
    Dim dAffVisitId As Double
    Dim sAffAtpRateId As String
    Dim iX As Integer

On Error GoTo cmdRecoup_ClickErr

'    MsgBox "Recoupment temporarily disabled", vbInformation
'    On Error Resume Next
'    Set cmd = Nothing
'    Exit Sub

    Set cmd.ActiveConnection = gcnDDS
    Hourglass True
    If dbgPreEdit.SelBookmarks.Count > 6 Then
        MsgBox "Currently the program only allows for 6 Credits for 1 Debit", vbInformation
        Exit Sub
    End If
    
    'Validate the first debit record you come across
    For iX = 0 To dbgPreEdit.SelBookmarks.Count - 1
        'Find the debit record in the group
        rsPreEdit.Move 0, (dbgPreEdit.SelBookmarks(iX))
        If rsPreEdit.Fields("Type") = "DR" Then
            If dDRInvalidRecordID > 0 Then
                MsgBox "Only one Debit record allowed for recoupment.", vbInformation
                Exit Sub
            Else
                dDRInvalidRecordID = rsPreEdit.Fields("InvalidRecordID")
            End If
        Else
            'Fill the array of invalid credit records
            If iX = 0 Then
                dCRInvalidRecordID(iX) = rsPreEdit.Fields("InvalidRecordID")
            Else
                If dCRInvalidRecordID(iX - 1) > 0 Then
                    dCRInvalidRecordID(iX) = rsPreEdit.Fields("InvalidRecordID")
                Else
                    dCRInvalidRecordID(iX - 1) = rsPreEdit.Fields("InvalidRecordID")
                End If
            End If
        End If
    Next iX
    
    If dDRInvalidRecordID > 0 Then
        'Validate the debit record.  If the record is valid continue on with the recoupment
        If ValidateRecoupNew(dDRInvalidRecordID, iDaysInHouse, sATP_PML_FLAG, dAffVisitId, sAffAtpRateId) = True Then
            
            'Commented this out temporarily to test
            
            
            'If valid for Recoupment then update the
'            With cmd
'                If gStoredProcs("up_iu_Trans_RECOUP").GetStoredProcCommand(cmd) = True Then
'                    For iX = 0 To 5
'                        If dCRInvalidRecordID(iX) = 0 Then
'                            cmd.Parameters("dCRInvalidRec" & iX + 1) = Null
'                        Else
'                            cmd.Parameters("dCRInvalidRec" & iX + 1) = dCRInvalidRecordID(iX)
'                        End If
'                    Next iX
'
'                    cmd.Parameters("dDRInvalidRecordID") = dDRInvalidRecordID
'                    cmd.Parameters("dTotDaysInhouse") = iDaysInHouse
'                    cmd.Parameters("sAtpPmlFlag") = sATP_PML_FLAG
'                    cmd.Parameters("dAffinityVisitID") = dAffVisitId
'                    cmd.Parameters("sAffinityAtpRateID") = sAffAtpRateId
'                    cmd.Parameters("sUserID") = gobjLoginInfo.UserId
'                    cmd.Execute
'                    lReturnValue = cmd.Parameters("RETURN_VALUE").Value
'                    If lReturnValue <> 0 Then
'                        GetServerErrorMsg lReturnValue, "Error occurs when processing invalid record (ID: " & dDRInvalidRecordID & "error message from server follows: "
'                    Else
'                        MsgBox "The transactions have been posted successfully.", vbInformation
'                        RefreshPreEditGrid (sSortedSql)
'                        EnablePreEditUI
'                    End If
'                End If
'            End With
        Else
            MsgBox "These records are not valid for recoupment.", vbExclamation
        End If
    End If
    
    Hourglass False

Xit:
    Exit Sub

cmdRecoup_ClickErr:

    ShowUnexpectedError MODULE + " cmdRecoup_Click", Err
    Resume Xit


End Sub


Private Sub cmdResolveMulti_Click()
On Error GoTo cmdResolveMultiErr

Dim rsInvalid As New ADODB.Recordset
Dim rsToCorrect As New ADODB.Recordset
Dim sSql As String
Dim vMark As Variant
Dim iComma As Integer
Dim iFound As Integer
Dim sCompareLastName As String
Dim dInvalidRecIdHold As Double
Dim iCount As Integer
Dim bConflict As Boolean
Dim sPatientNameHold(1 To 10) As String
Dim sLastName As String
Dim sLastNameHold As String
Dim dTotAmtHold As Double
Dim sInstIdHold As String

'***************************************************
'* Retreive the records with shared-dd_num_ind = 'Y'
'***************************************************
Hourglass True
sSql = "SELECT DD_INVALID_REC.INVALID_RECORD_ID ,DD_NUM, TOT_FUNB_BENEFIT_AMT, DR_CR_FLAG, DD_SHARED_DD_NUM.INSTITUTION_CODE, DD_SHARED_DD_NUM.MEDICAL_RECORD_NUM, DD_SHARED_DD_NUM.PATIENT_NAME, DD_SHARED_DD_NUM.DECEASED_IND"
sSql = sSql & " FROM DD_INVALID_REC, DD_SHARED_DD_NUM"
sSql = sSql & " WHERE DD_INVALID_REC.INVALID_RECORD_ID = DD_SHARED_DD_NUM.INVALID_RECORD_ID AND DD_INVALID_REC.RECORD_STATUS = 'A' AND DD_INVALID_REC.SHARED_DD_NUM_IND = 'Y'"
sSql = sSql & " ORDER BY TOT_FUNB_BENEFIT_AMT, DD_NUM, DD_INVALID_REC.INVALID_RECORD_ID,DD_SHARED_DD_NUM.PATIENT_NAME"

rsInvalid.Open sSql, gcnDDS, adOpenStatic
'Mark the first record to resolve
If rsInvalid.EOF = False Then
    vMark = rsInvalid.Bookmark
    iComma = InStr(1, rsInvalid!PATIENT_NAME, ",")
    sLastNameHold = Trim$(Left$(rsInvalid!PATIENT_NAME, iComma - 1))
    dInvalidRecIdHold = rsInvalid!INVALID_RECORD_ID
    sPatientNameHold(1) = rsInvalid!PATIENT_NAME
    dTotAmtHold = rsInvalid!TOT_FUNB_BENEFIT_AMT
    sInstIdHold = rsInvalid!INSTITUTION_CODE
    iCount = 1
End If
Do Until rsInvalid.EOF
    'Determine to last name of the patient
    'We cannot resolve a multi mrun situation if the last name is different for any of the records
    'Get the Last Name for the first record
    Do Until bConflict = True
        rsInvalid.MoveNext
        If rsInvalid.EOF Then
            bConflict = True
        Else
            If dInvalidRecIdHold = rsInvalid!INVALID_RECORD_ID And dTotAmtHold = rsInvalid!TOT_FUNB_BENEFIT_AMT Then
                iComma = InStr(1, rsInvalid!PATIENT_NAME, ",")
                sLastName = Trim$(Left$(rsInvalid!PATIENT_NAME, iComma - 1))
                If sLastNameHold = sLastName And sInstIdHold = rsInvalid!INSTITUTION_CODE Then
                    iCount = iCount + 1
                    sPatientNameHold(iCount) = rsInvalid!PATIENT_NAME
                    If iCount > 10 Then
                        bConflict = True
                    End If
                Else
                    bConflict = True
                End If
            Else
                If iCount = 1 Then
                    bConflict = True
                Else
                    Exit Do
                End If
            End If
        End If
    Loop

    If bConflict = True Then
        'We cannot determine the intent of who receives the money
        'Go to the next invalid record id
        rsInvalid.Bookmark = vMark
        Do Until rsInvalid!INVALID_RECORD_ID <> dInvalidRecIdHold
            rsInvalid.MoveNext
            If rsInvalid.EOF Then
                Exit Do
            End If
        Loop
    Else
        'Go to the starting record
        rsInvalid.Bookmark = vMark
        sSql = "SELECT * FROM DD_INVALID_REC"
        sSql = sSql & " WHERE DD_NUM = '" & rsInvalid!DD_NUM & "' AND TOT_FUNB_BENEFIT_AMT = " & rsInvalid!TOT_FUNB_BENEFIT_AMT & " AND DR_CR_FLAG = '" & rsInvalid!DR_CR_FLAG & "' AND SHARED_DD_NUM_IND = 'Y'"
        sSql = sSql & " ORDER BY INVALID_RECORD_ID"
        rsToCorrect.Open sSql, gcnDDS, adOpenStatic
        If rsToCorrect.EOF = False Then
            rsToCorrect.MoveLast
            If rsToCorrect.RecordCount = iCount Then
                'We Can Resolve this MRUN situation
                rsToCorrect.MoveFirst
                Do Until rsToCorrect.EOF
                    If UpdateMultiMrunRecord(rsToCorrect!INVALID_RECORD_ID, Format$(rsInvalid!MEDICAL_RECORD_NUM, "0000000"), rsInvalid!INSTITUTION_CODE, rsInvalid!PATIENT_NAME, IIf(rsInvalid!DECEASED_IND = "Y", True, False)) = False Then
                        Set rsInvalid = Nothing
                        Set rsToCorrect = Nothing
                        MsgBox "Error updating Multi Mrun Record.  PLease review"
                        GoTo Xit
                    End If
                    rsToCorrect.MoveNext
                    rsInvalid.MoveNext
                Loop
            End If
        End If
        
        rsInvalid.Bookmark = vMark
        rsToCorrect.MoveFirst
        Do Until rsInvalid.EOF
            If rsInvalid!DD_NUM <> rsToCorrect!DD_NUM Then
                Exit Do
            Else
                rsInvalid.MoveNext
            End If
        Loop
    End If

    If rsInvalid.EOF Then
        Exit Do
    End If
    
    Set rsToCorrect = Nothing
    'Mark the new record
    vMark = rsInvalid.Bookmark
    iComma = InStr(1, rsInvalid!PATIENT_NAME, ",")
    sLastNameHold = Trim$(Left$(rsInvalid!PATIENT_NAME, iComma - 1))
    dInvalidRecIdHold = rsInvalid!INVALID_RECORD_ID
    sInstIdHold = rsInvalid!INSTITUTION_CODE
    dTotAmtHold = rsInvalid!TOT_FUNB_BENEFIT_AMT
    iCount = 1
    bConflict = False
Loop

RefreshPreEditGrid (sSortedSql)
 'don't forget to call this after refresh the grids
 EnablePreEditUI
Hourglass False
MsgBox "Finished resolving Multi MRUN situations."

Xit:
Set rsInvalid = Nothing
Set rsToCorrect = Nothing
Hourglass False
Exit Sub
    
cmdResolveMultiErr:
    
    ShowUnexpectedError MODULE + " cmdUndelete_Click", Err
    Resume Xit

End Sub

Private Sub cmdUndelete_Click()

On Error GoTo cmdUndelete_ClickErr

    Beep
    If MsgBox("Are you sure you want to unhide these records?", vbYesNo + vbExclamation) = vbYes Then
        UndeleteInvalidRecords
        'refresh grids
        RefreshPreEditGrid (sSortedSql)
        RefreshDeletedRecordsGrid (sDeletedRecordsSql)
        'don't forget to call this after refresh the grids
        EnablePreEditUI
        EnableDeletedUI
    End If


Xit:
    Exit Sub

cmdUndelete_ClickErr:
    ShowUnexpectedError MODULE + " cmdUndelete_Click", Err
    Resume Xit


End Sub


Private Sub cmdValidate_Click()

On Error GoTo cmdValidate_ClickErr

    giProcess = VALIDATE_ONLY
    ValidateRecords
    RefreshPreEditGrid (sSortedSql)
'    RefreshDeletedRecordsGrid (sDeletedRecordsSql)
    'don't forget to call this after refresh the grids
    EnablePreEditUI
    EnableDeletedUI

Xit:
    Exit Sub

cmdValidate_ClickErr:
    ShowUnexpectedError MODULE + " cmdValidate_Click", Err
    Resume Xit


End Sub





Private Sub cmdSelectMRUN_Click()
    Dim fSharedDD As New frmPreEditSharedDD
    Dim dInvalidRecordID As Double
    Dim bRet As Boolean

On Error GoTo cmdSelectMRUN_ClickErr

    rsPreEdit.Move 0, dbgPreEdit.SelBookmarks(0)
    fSharedDD.InvalidRecordID = rsPreEdit.Fields("InvalidRecordID")
    fSharedDD.Show vbModal
    If fSharedDD.MsgBoxResult = vbOK Then
        Hourglass True
        Dim sMRUN As String
        Dim sInstitutionCode As String
        Dim sPatientName As String
        Dim bDeceasedInd As Boolean
        Dim lReturnValue As Double
        sMRUN = fSharedDD.MRUN
        sInstitutionCode = fSharedDD.Institution
        sPatientName = fSharedDD.PatientName
        bDeceasedInd = fSharedDD.DeceasedInd
        'save the invalid record id
        dInvalidRecordID = rsPreEdit.Fields("InvalidRecordID")
        bRet = UpdateMultiMrunRecord(dInvalidRecordID, sMRUN, sInstitutionCode, sPatientName, bDeceasedInd)
        'update the database
        If bRet = True Then
            'refresh the database
            RefreshPreEditGrid (sSortedSql)
            'don't forget to call this after refresh the grids
            EnablePreEditUI
            'search the recordset using the previously saved invalid record id
            dbgPreEdit.Redraw = False
            rsPreEdit.Find "InvalidRecordID = " & CStr(dInvalidRecordID)
            dbgPreEdit.Redraw = True
            Hourglass False
        End If
    End If
    Unload fSharedDD
    Set fSharedDD = Nothing


Xit:
    Exit Sub

cmdSelectMRUN_ClickErr:
    ShowUnexpectedError MODULE + " cmdSelectMRUN_Click", Err
    Resume Xit


End Sub

Private Function UpdateMultiMrunRecord(ByVal dInvalidRecordID As Double, ByVal sMRUN As String, ByVal sInstitutionCode As String, ByVal sPatientName As String, ByVal bDeceasedInd As String) As Boolean
On Error GoTo UpdateMultiMrunRecordErr

    Dim cmdUpdate As New ADODB.Command
    Dim lReturnValue As Long
    Dim iX As Integer
    Set cmdUpdate.ActiveConnection = gcnDDS
    If gStoredProcs("up_u_Multi_MRUN").GetStoredProcCommand(cmdUpdate) = True Then
        For iX = 0 To cmdUpdate.Parameters.Count - 1
            cmdUpdate.Parameters(iX) = Null
        Next iX
        
        cmdUpdate.Parameters("dInvalidRecordID") = dInvalidRecordID
        cmdUpdate.Parameters("sMRUN") = sMRUN
        cmdUpdate.Parameters("sInstitutionCode") = sInstitutionCode
        cmdUpdate.Parameters("sPatientName") = sPatientName
        If bDeceasedInd = True Then
            cmdUpdate.Parameters("sDeceasedInd") = "Y"
        Else
            cmdUpdate.Parameters("sDeceasedInd") = "N"
        End If
        cmdUpdate.Parameters("sUserID") = gobjLoginInfo.UserId
        cmdUpdate.Execute
        lReturnValue = cmdUpdate.Parameters("RETURN_VALUE").value
        If lReturnValue <> 0 Then
            GetServerErrorMsg lReturnValue, "Error occurs when updating invalid record (ID: " & dInvalidRecordID & " ) " _
            & "error message from server follows: "
            UpdateMultiMrunRecord = False
        Else
            UpdateMultiMrunRecord = True
        End If
    End If
    
    Set cmdUpdate = Nothing

Exit Function

UpdateMultiMrunRecordErr:
    
    Set cmdUpdate = Nothing
    UpdateMultiMrunRecord = False

End Function


Private Sub dbdValidationErrors_Click()

On Error GoTo dbdValidationErrors_ClickErr

    If PreEditMode = MoveEditDelete Then
        dbgPreEdit.CancelUpdate
    Else
        dbgDeletedRecords.CancelUpdate
    End If

Xit:
    Exit Sub

dbdValidationErrors_ClickErr:
    ShowUnexpectedError MODULE + " dbdValidationErrors_Click", Err
    Resume Xit


End Sub

Private Sub cmdViewDetails_Click()

Dim cr As CRAXDRT.Report
Dim crSub As CRAXDRT.Report
Dim frm As Form1
Dim sDDNum As String
Dim sMRUNString As String
Dim sTemp As String
On Error GoTo CERR

If IsEmpty(dbgPreEdit.SelBookmarks(0)) Then
    MsgBox "No records are selected"
    Exit Sub
End If

Hourglass True
rsPreEdit.Move 0, dbgPreEdit.SelBookmarks(0)
If IsNull(rsPreEdit.Fields("DD #")) Then
    Err.Raise 23455, , "Details are not allowed on blank DD Number"
End If
sDDNum = rsPreEdit.Fields("DD #")

Set cr = gcrApp.OpenReport(gsDataPath & "\viewinfo.rpt", crOpenReportByTempCopy)

Set frm = New Form1
Dim sSql As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rsADOTemp As New ADODB.Recordset

    Screen.MousePointer = vbHourglass
    sSql = "SELECT * FROM DD_ATP_PML_INFO"
    sSql = sSql & " WHERE DD_NUMBER Like '" & sDDNum & "%'" ' AND MRUN = '" & sMRUN & "'  AND INSTITUTION_CODE = '" & sInstCode & "'"
    sSql = sSql & " ORDER BY FROM_DATE ASC, THRU_DATE ASC"
    rs.Open sSql, gcnDDS, adOpenStatic
    Do Until rs.EOF
        If InStr(1, sMRUNString, rs!MRUN, vbTextCompare) = 0 Then
            sMRUNString = sMRUNString & "'" & rs!MRUN & "',"
        End If
        rs.MoveNext
    Loop
    
    If rs.RecordCount > 0 Then
        sMRUNString = Left$(sMRUNString, Len(sMRUNString) - 1)
        rs.MoveFirst
    Else
        MsgBox "There are no ATP/PML records matching DD Number selected"
        GoTo Xit
    End If
    cr.ParameterFields.Item(1).AddCurrentValue sDDNum
    
    If Len(sMRUNString) > 0 And InStr(1, sMRUNString, ",", vbTextCompare) = 0 Then
        sSql = "SELECT CURRENT_POPULATION.*, PF_INSTITUTION.INSTITUTION_NAME "
        sSql = sSql & " FROM CURRENT_POPULATION, pfs.dbo.PF_INSTITUTION PF_INSTITUTION"
        sSql = sSql & " WHERE CURRENT_POPULATION.INSTITUTION_CODE = PF_INSTITUTION.INSTITUTION_CODE"
        sSql = sSql & " AND PATIENT_MRUN = " & sMRUNString
        rsADOTemp.Open sSql, gcnDDS
        If Not rsADOTemp.EOF Then
            cr.ParameterFields.Item(2).AddCurrentValue "Note: Patient currently active at " & rsADOTemp!INSTITUTION_NAME
        Else
            If rs!DEATH_FLAG = "Y" Then
                cr.ParameterFields.Item(2).AddCurrentValue "Note: Patient is deceased"
            Else
                cr.ParameterFields.Item(2).AddCurrentValue "Note: Patient is discharged"
            End If
        End If
        rsADOTemp.Close
    Else
        cr.ParameterFields.Item(2).AddCurrentValue ""
    End If
     
    sTemp = rsPreEdit.Fields("Income Source")
    cr.ParameterFields.Item(3).AddCurrentValue sTemp
    sTemp = rsPreEdit!Memo
    cr.ParameterFields.Item(4).AddCurrentValue sTemp
    sTemp = rsPreEdit.Fields("FUNB As Of Date")
    cr.ParameterFields.Item(5).AddCurrentValue sTemp
    cr.ParameterFields.Item(6).AddCurrentValue CDbl(rsPreEdit.Fields("FUNB Amount"))
    cr.Database.SetDataSource rs
    'cr.ReadRecords
     
    Set crSub = cr.OpenSubreport("h:\dds\viewinfo2.rpt")
    sSql = "SELECT INSTITUTION_CODE,MRUN,ACCOUNT_NUMBER,ADMIT_ARRIVE_DATE,DISCHARGE_DISPOSITION_DATE,PATIENT_SERVICE_CODE FROM DD_VISIT_INFO"
    sSql = sSql & " WHERE MRUN IN ( " & sMRUNString & ")"
    rs2.Open sSql, gcnDDS, adOpenStatic
    crSub.Database.SetDataSource rs2
    'crSub.ReadRecords
    
    frm.CRViewer1.ReportSource = cr
    frm.CRViewer1.ViewReport
    frm.Caption = "Direct Deposit Review"
    frm.WindowState = vbMaximized
    Set frm.oReport = cr
    frm.Show
Xit:
    Screen.MousePointer = vbDefault
    Set rs = Nothing
    Set rs2 = Nothing
    Set rsADOTemp = Nothing
    Set cr = Nothing
    Set crSub = Nothing
    Set frm = Nothing
Exit Sub

CERR:
    MsgBox Error
    Resume Xit
    
    
End Sub

Private Sub dbgDeletedRecords_Click()
    EnableDeletedUI

End Sub

Private Sub dbgDeletedRecords_HeadClick(ByVal ColIndex As Integer)

On Error GoTo dbgDeletedRecords_HeadClickErr

    Static DeletedRecordsColumnFlag As udtSortedColumnFlag
    Hourglass True
    If DeletedRecordsColumnFlag.ColIndex = ColIndex And DeletedRecordsColumnFlag.Ascending = True Then
        sDeletedRecordsSql = "[" & dbgDeletedRecords.Columns(ColIndex).Name & "] Desc,[FUNB As Of Date] Asc"
        DeletedRecordsColumnFlag.Ascending = False
    Else
        sDeletedRecordsSql = "[" & dbgDeletedRecords.Columns(ColIndex).Name & "] Asc,[FUNB As Of Date] Asc"
        DeletedRecordsColumnFlag.Ascending = True
    End If
    DeletedRecordsColumnFlag.ColIndex = ColIndex
    RefreshDeletedRecordsGrid (sDeletedRecordsSql)
    Hourglass False

Xit:
    Exit Sub

dbgDeletedRecords_HeadClickErr:
    ShowUnexpectedError MODULE + " dbgDeletedRecords_HeadClick", Err
    Resume Xit


End Sub




Private Sub dbgDeletedRecords_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    
    ChangeDeletedError

End Sub
Private Sub ChangeDeletedError()
On Error GoTo ChangeDeletedErr
    Dim sSql As String
    Dim rsErrors As New ADODB.Recordset
    
    If rsDeletedRecords.EOF Then
        txtDeletedErrors.Text = ""
    Else
        sSql = "SELECT INVALID_REC_ERR_MSG FROM DD_INVALID_REC_ERROR"
        sSql = sSql & " WHERE INVALID_RECORD_ID = " & rsDeletedRecords!InvalidRecordID
        rsErrors.Open sSql, gcnDDS, adOpenForwardOnly
        txtDeletedErrors.Text = ""
        Do Until rsErrors.EOF
            txtDeletedErrors = txtDeletedErrors & rsErrors!INVALID_REC_ERR_MSG & vbCrLf
            rsErrors.MoveNext
        Loop
    End If
    
    On Error Resume Next
    rsErrors.Close
Exit Sub
    
ChangeDeletedErr:
    MsgBox Error, vbInformation


End Sub

Private Sub dbgDeletedRecords_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
    EnableDeletedUI

End Sub


Private Sub dbgDeletedRecords_UnboundReadData(ByVal RowBuf As SSDataWidgets_B_OLEDB.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim r, i, j As Integer
Dim ct As Integer
ct = 14

If rsDeletedRecords Is Nothing Then
    Exit Sub
End If
    
If rsDeletedRecords.RecordCount = 0 Then
    Exit Sub
End If
    
    If IsNull(StartLocation) Then
    If ReadPriorRows Then
        rsDeletedRecords.MoveLast
    Else
        rsDeletedRecords.MoveFirst
    End If

Else
    rsDeletedRecords.Bookmark = StartLocation
    If ReadPriorRows Then
        rsDeletedRecords.MovePrevious
    Else
        rsDeletedRecords.MoveNext
    End If

End If

For i = 0 To RowBuf.RowCount - 1
    If rsDeletedRecords.BOF Or rsDeletedRecords.EOF Then Exit For

    For j = 0 To ct - 1
        RowBuf.value(i, j) = rsDeletedRecords(j)
    Next j

    RowBuf.Bookmark(i) = rsDeletedRecords.Bookmark

If ReadPriorRows Then
        rsDeletedRecords.MovePrevious
    Else
        rsDeletedRecords.MoveNext
    End If

    r = r + 1

Next i

RowBuf.RowCount = r


End Sub

Private Sub dbgPreEdit_Click()
    EnablePreEditUI
End Sub

Private Sub dbgPreEdit_DblClick()
    If cmdEdit.Enabled = True Then
        cmdEdit.value = True
    End If
End Sub

Private Sub dbgPreEdit_HeadClick(ByVal ColIndex As Integer)

On Error GoTo dbgPreEdit_HeadClickErr

    Static PreEditColumnFlag As udtSortedColumnFlag
    Hourglass True
    If dbgPreEdit.Columns(ColIndex).Name <> "FUNB As Of Date" Then
        If PreEditColumnFlag.ColIndex = ColIndex And PreEditColumnFlag.Ascending = True Then
            sSortedSql = "[" & dbgPreEdit.Columns(ColIndex).Name & "] Desc,[FUNB As Of Date] Asc"
            PreEditColumnFlag.Ascending = False
        Else
            sSortedSql = "[" & dbgPreEdit.Columns(ColIndex).Name & "] Asc,[FUNB As Of Date] Asc"
            PreEditColumnFlag.Ascending = True
        End If
    Else
        If PreEditColumnFlag.ColIndex = ColIndex And PreEditColumnFlag.Ascending = True Then
            sSortedSql = "[" & dbgPreEdit.Columns(ColIndex).Name & "] Desc"
            PreEditColumnFlag.Ascending = False
        Else
            sSortedSql = "[" & dbgPreEdit.Columns(ColIndex).Name & "] Asc"
            PreEditColumnFlag.Ascending = True
        End If
    End If
    PreEditColumnFlag.ColIndex = ColIndex
    
    RefreshPreEditGrid (sSortedSql)
    Hourglass False

Xit:
    Exit Sub

dbgPreEdit_HeadClickErr:
    ShowUnexpectedError MODULE + " dbgPreEdit_HeadClick", Err
    Resume Xit


End Sub




Private Sub dbgPreEdit_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    
    ChangePreEditError

End Sub
Private Sub ChangePreEditError()
On Error GoTo ChangePreEditErr
    Dim sSql As String
    Dim rsErrors As New ADODB.Recordset
    
    If rsPreEdit.EOF Then
        txtErrors.Text = ""
    Else
        sSql = "SELECT INVALID_REC_ERR_MSG FROM DD_INVALID_REC_ERROR"
        sSql = sSql & " WHERE INVALID_RECORD_ID = " & rsPreEdit!InvalidRecordID
        rsErrors.Open sSql, gcnDDS, adOpenForwardOnly
        txtErrors.Text = ""
        Do Until rsErrors.EOF
            txtErrors = txtErrors & rsErrors!INVALID_REC_ERR_MSG & vbCrLf
            rsErrors.MoveNext
        Loop
    
        If Not IsNull(rsPreEdit.Fields("Message Control ID")) Then
            txtErrors = txtErrors & "****Posting Error*****"
            txtErrors = txtErrors & " Message Control Id: " & rsPreEdit.Fields("Message Control Id")
            txtErrors = txtErrors & " PA Posting Status: " & rsPreEdit.Fields("PA Posting Status")
            txtErrors = txtErrors & " PF Posting Status: " & rsPreEdit.Fields("PF Posting Status")
            txtErrors = txtErrors & " Posting Errors Created Date: " & rsPreEdit.Fields("Posting Errors Created Date")
            txtErrors = txtErrors & " PA Error Code: " & rsPreEdit.Fields("PA Error Code")
            txtErrors = txtErrors & " PF Error Code: " & rsPreEdit.Fields("PF Error Code")
        End If
    End If
    
    On Error Resume Next
    rsErrors.Close
Exit Sub

    
ChangePreEditErr:
    MsgBox Error, vbInformation

End Sub

Private Sub dbgPreEdit_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
    
    EnablePreEditUI

End Sub

Private Sub dbgPreEdit_UnboundReadData(ByVal RowBuf As SSDataWidgets_B_OLEDB.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim r, i, j As Integer
Dim ct As Integer
ct = 14

If rsPreEdit Is Nothing Then
    Exit Sub
End If
    
If rsPreEdit.RecordCount = 0 Then
    Exit Sub
End If
    
    If IsNull(StartLocation) Then
    If ReadPriorRows Then
        rsPreEdit.MoveLast
    Else
        rsPreEdit.MoveFirst
    End If

Else
    rsPreEdit.Bookmark = StartLocation
    If ReadPriorRows Then
        rsPreEdit.MovePrevious
    Else
        rsPreEdit.MoveNext
    End If

End If

For i = 0 To RowBuf.RowCount - 1
    If rsPreEdit.BOF Or rsPreEdit.EOF Then Exit For

    For j = 0 To ct - 1
        RowBuf.value(i, j) = rsPreEdit(j)
    Next j

    RowBuf.Bookmark(i) = rsPreEdit.Bookmark

If ReadPriorRows Then
        rsPreEdit.MovePrevious
    Else
        rsPreEdit.MoveNext
    End If

    r = r + 1

Next i

RowBuf.RowCount = r


End Sub

Private Sub Form_Activate()
    
    fMainForm.SetMainToolbar False



End Sub

Private Sub Form_Deactivate()
    fMainForm.SetMainToolbar False

End Sub

Private Sub Form_Load()

On Error GoTo Form_LoadErr

    Set outTitle.Picture = fMainForm.imlToolbarIcons.ListImages("Pre-Edit File").Picture
    
    Set cmdPreEdit.ActiveConnection = gcnDDS
    If gStoredProcs("up_s_PreEditMain").GetStoredProcCommand(cmdPreEdit) = False Then
        Err.Raise 2345, , "PreEditMain Stored Procedure failed"
    End If
    
    InitAllGrid
    SetMode

Xit:
    Exit Sub

Form_LoadErr:
    ShowUnexpectedError MODULE + " Form_Load", Err
    Resume Xit


End Sub


Private Sub InitAllGrid()

On Error GoTo InitAllGridErr

    sSortedSql = "[Income Source] Asc"
    RefreshPreEditGrid (sSortedSql)
    sDeletedRecordsSql = ""
    '    RefreshDeletedRecordsGrid (sDeletedRecordsSql)
    sSourceTypeSql = "SELECT FUNB_INCOME_SRC_TYPE AS 'Source Type', INCOME_SRC_TYPE_DESCR  AS Description " _
                            & "FROM DD_INCOME_SOURCE_TYPE ORDER BY FUNB_INCOME_SRC_TYPE"
    'RefreshSourceTypeCombo (sSourceTypeSql)
    sValidationErrorsSql = "SELECT INVALID_REC_ERR_MSG AS 'Validation Errors' FROM DD_INVALID_REC_ERROR "
    Dim stempsql As String
    stempsql = sValidationErrorsSql
    'RefreshValidationErrorsCombo (stempsql)
    

Xit:
    Exit Sub

InitAllGridErr:
    ShowUnexpectedError MODULE + " InitAllGrid", Err
    Resume Xit

End Sub

Private Sub RefreshPreEditGrid(ByRef sSort As String)

On Error GoTo RefreshPreEditGridErr

    Hourglass True
    With dbgPreEdit
        .Redraw = False
        If rsPreEdit.State = adStateOpen Then
            rsPreEdit.CancelUpdate
            rsPreEdit.Close
        End If
        Set .DataSource = Nothing
        
        cmdPreEdit.Parameters("activeStatus") = "A"
        Set rsPreEdit = cmdPreEdit.Execute
        If sSort <> "" Then
            rsPreEdit.Sort = sSort
        End If
        .Rebind
        
'        Set .DataSource = rsPreEdit
'        .Refresh
        .Redraw = True
        rsPreEdit.MoveFirst
    End With
    lblPreEditRecs = "Total Pre-Edit Records: " & rsPreEdit.RecordCount
    Hourglass False
    

Xit:
    Exit Sub

RefreshPreEditGridErr:
    ShowUnexpectedError MODULE + " RefreshPreEditGrid", Err
    Resume Xit


End Sub

Private Sub RefreshDeletedRecordsGrid(ByRef sSort As String)

On Error GoTo RefreshDeletedRecordsGridErr

    Hourglass True
    With dbgDeletedRecords
        .Redraw = False
        Set .DataSource = Nothing
        If rsDeletedRecords.State = adStateOpen Then
            rsDeletedRecords.Close
        End If
        cmdPreEdit.Parameters("activeStatus") = "I"
        Set rsDeletedRecords = cmdPreEdit.Execute
        If sSort <> "" Then
            rsDeletedRecords.Sort = sSort
        End If
        'rsDeletedRecords.Open sSql, gcnDDS, adOpenDynamic, adLockOptimistic
        .Rebind
'        Set .DataSource = rsDeletedRecords
'        .Refresh
        .Redraw = True
        rsDeletedRecords.MoveFirst
    End With
    Hourglass False


Xit:
    Exit Sub

RefreshDeletedRecordsGridErr:
    ShowUnexpectedError MODULE + " RefreshDeletedRecordsGrid", Err
    Resume Xit


End Sub

Public Function Search(ByVal sSearchString As String, ByVal sColIndex As String, ByVal bMatchCase As Boolean, ByVal bExactString As Boolean, ByVal bFindFirst As Boolean) As Boolean
    Dim Grid As SSOleDBGrid
    Dim rs As ADODB.Recordset
    Dim vOldBookmark As Variant
    Dim bFound As Boolean

On Error GoTo SearchErr

    bFound = False
    Search = False
    Dim i As Long
    If PreEditMode = MoveEditDelete Then
        Set Grid = dbgPreEdit
        Set rs = rsPreEdit
    ElseIf PreEditMode = Undelete Then
        Set Grid = dbgDeletedRecords
        Set rs = rsDeletedRecords
    End If
    Grid.Redraw = False
    vOldBookmark = rs.Bookmark
    
    Dim sTemp As String
    If Not bMatchCase Then
        sSearchString = UCase(sSearchString)
    End If
    With rs
        If bFindFirst Then
            .MoveFirst
        Else
            .MoveNext
        End If
        Do While Not .EOF
            sTemp = IIf(IsNull(rs.Fields(sColIndex).value), "", rs.Fields(sColIndex).value)
            If Not bMatchCase Then
                sTemp = UCase(sTemp)
            End If
            If bExactString Then
                If sTemp = sSearchString Then
                    bFound = True
                    Exit Do
                End If
            Else
                If InStr(1, sTemp, sSearchString) > 0 Then
                    bFound = True
                    Exit Do
                End If
            End If
            .MoveNext
        Loop
    End With
        
    If bFound Then
        Search = True
    Else
        Search = False
        rs.Move 0, vOldBookmark
    End If
    Grid.Redraw = True
            
    


Xit:
    Set Grid = Nothing
    Set rs = Nothing
    Exit Function

SearchErr:
    ShowUnexpectedError MODULE + "Search", Err
    Resume Xit


End Function




Private Sub SetMode()

On Error GoTo SetModeErr

    Select Case sstPreEdit.Tab
        Case 0
            PreEditMode = MoveEditDelete
            EnablePreEditUI
        Case 1
            'MsgBox "Refreshing deleted rows. Please wait ", vbInformation
            Hourglass True
            RefreshDeletedRecordsGrid (sDeletedRecordsSql)
            PreEditMode = Undelete
            EnableDeletedUI
            Hourglass False
    End Select

Xit:
    Exit Sub

SetModeErr:
    ShowUnexpectedError MODULE + " SetMode", Err
    Resume Xit


End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo Form_UnloadErr

    'doing clean up
    'Release the reference first, otherwise it could bring the system down
    Set dbgPreEdit.DataSource = Nothing
    If rsPreEdit.State = adStateOpen Then
        rsPreEdit.Close
    End If
    Set rsPreEdit = Nothing
    Set dbgDeletedRecords.DataSource = Nothing
    If rsDeletedRecords.State = adStateOpen Then
        rsDeletedRecords.Close
    End If
    Set rsDeletedRecords = Nothing
    If rsSourceType.State = adStateOpen Then
        rsSourceType.Close
    End If
    Set rsSourceType = Nothing
    If rsValidationErrors.State = adStateOpen Then
        rsValidationErrors.Close
    End If
    Set cmdPreEdit = Nothing
    
    Set rsValidationErrors = Nothing
    

Xit:
    Exit Sub

Form_UnloadErr:
    ShowUnexpectedError MODULE + "Form_Unload", Err
    Resume Xit


End Sub



Private Sub outTitle_IconClick()
    If cmdClose.Enabled = True Then
        Unload Me
    End If

End Sub

Private Sub sstPreEdit_Click(PreviousTab As Integer)
    SetMode
    
End Sub

Private Function CheckRecoup() As Boolean

On Error GoTo CheckRecoupErr

    CheckRecoup = False
    'turn the grid's redraw off
    dbgPreEdit.Redraw = False
    If dbgPreEdit.SelBookmarks.Count < 2 Then
        dbgPreEdit.Redraw = True
        Exit Function
    End If
    Dim bmkCurrent As Variant
    'save the data for the current row of recordset
    Dim sDDNum As String
    Dim dFUNBAmount As Double
    Dim sIncomeSource As String
    Dim sMRUN As String
    Dim sInstitution As String
    Dim sType As String
    Dim iLastBookmark As Integer
    Dim iX As Integer
    Dim iDRCount As Integer
    Dim iCRCount As Integer
    Dim dDRTotal As Double
    Dim dCRTotal As Double
    
    With rsPreEdit
        sDDNum = .Fields("DD #")
        dFUNBAmount = .Fields("FUNB Amount")
        sIncomeSource = .Fields("Income Source")
        sType = .Fields("Type")
    End With
    'Loop through the bookmarks and verify these records
    iLastBookmark = dbgPreEdit.SelBookmarks.Count - 1
    For iX = 0 To iLastBookmark
        rsPreEdit.Move 0, dbgPreEdit.SelBookmarks(iX)
        'doing comparison
        If rsPreEdit.Fields("DD #") <> sDDNum Then
            'move to the original row
            rsPreEdit.Move 0, dbgPreEdit.SelBookmarks(iLastBookmark)
            dbgPreEdit.Redraw = True
            Exit Function
        End If
        
        
        If rsPreEdit.Fields("FUNB Amount") <> dFUNBAmount Then
            'move to the original row
            rsPreEdit.Move 0, dbgPreEdit.SelBookmarks(iLastBookmark)
            dbgPreEdit.Redraw = True
            Exit Function
        End If
        If rsPreEdit.Fields("Income Source") <> sIncomeSource Then
            'move to the original row
            rsPreEdit.Move 0, dbgPreEdit.SelBookmarks(iLastBookmark)
            dbgPreEdit.Redraw = True
            Exit Function
        End If
    
        If rsPreEdit.Fields("Type") = "DR" Then
            iDRCount = iDRCount + 1
            dDRTotal = dDRTotal + rsPreEdit.Fields("FUNB Amount")
        End If
    
        If iDRCount > 1 Then
            'move to the original row
            rsPreEdit.Move 0, dbgPreEdit.SelBookmarks(iLastBookmark)
            dbgPreEdit.Redraw = True
            Exit Function
        End If
        
        If rsPreEdit.Fields("Type") = "CR" Then
            iCRCount = iCRCount + 1
            dCRTotal = dCRTotal + rsPreEdit.Fields("FUNB Amount")
        End If
    
    Next iX
    
    If dCRTotal <> dDRTotal Then
        'move to the original row
        rsPreEdit.Move 0, dbgPreEdit.SelBookmarks(iLastBookmark)
        dbgPreEdit.Redraw = True
        Exit Function
    End If
    
    dbgPreEdit.Redraw = True
    CheckRecoup = True
        

Xit:

Exit Function

CheckRecoupErr:
'   If anything goes wrong, just returns false
    Resume Xit


End Function
Private Function CheckMove() As Boolean

On Error GoTo CheckMoveErr

    CheckMove = False
    'turn the grid's redraw off
    dbgPreEdit.Redraw = False
    'there should be exactly 2 selected rows to begin with
    If dbgPreEdit.SelBookmarks.Count <> 1 Then
        dbgPreEdit.Redraw = True
        Exit Function
    End If
    Dim bmkCurrent As Variant
    'save the data for the current row of recordset
    bmkCurrent = rsPreEdit.Bookmark
    rsPreEdit.Move 0, dbgPreEdit.SelBookmarks(0)
    If rsPreEdit.Fields("Deceased") = 1 And rsPreEdit.Fields("Type") = "CR" Then
        CheckMove = True
    End If

Xit:
    rsPreEdit.Move 0, bmkCurrent
    dbgPreEdit.Redraw = True
    Exit Function

CheckMoveErr:
'   If anything goes wrong, just returns false
    Resume Xit


End Function

