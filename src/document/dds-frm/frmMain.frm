VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{26807351-DE4B-11D2-9C83-00105A19BCF2}#1.0#0"; "VertMenu.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{8CEC1091-A679-11D3-9BDE-00105A19BCF2}#1.0#0"; "Subclas.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "HEARTS Direct Deposit Subsystem"
   ClientHeight    =   7275
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   8880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock WinsockEmail 
      Left            =   3360
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5280
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0896
            Key             =   "Face03"
         EndProperty
      EndProperty
   End
   Begin SubclassCtl.Subclass SizeSubclass 
      Left            =   5100
      Top             =   1065
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Height          =   6615
      Left            =   0
      ScaleHeight     =   6555
      ScaleWidth      =   1980
      TabIndex        =   1
      Top             =   390
      Width           =   2040
      Begin VertMenu.VerticalMenu VerticalMenu1 
         Height          =   9150
         Left            =   15
         TabIndex        =   2
         Top             =   -45
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   16140
         MenusMax        =   0
      End
      Begin MSAdodcLib.Adodc adcQuickSearch 
         Height          =   540
         Left            =   0
         Top             =   5430
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   953
         ConnectMode     =   1
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   1
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "QSearch"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin ComCtl3.CoolBar CoolBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   688
      BandCount       =   1
      VariantHeight   =   0   'False
      _CBWidth        =   8880
      _CBHeight       =   390
      _Version        =   "6.7.9816"
      Child1          =   "tbrMain"
      MinWidth1       =   4800
      MinHeight1      =   330
      Width1          =   4800
      Key1            =   "Main"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   330
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlToolBarSmallIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Login"
               Object.ToolTipText     =   "Login as new user"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "StepAway"
               Object.ToolTipText     =   "Step Away from Computer"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "S1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "Paste"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "S2"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Calculator"
               Object.ToolTipText     =   "Calculator"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7005
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2143
            MinWidth        =   1411
            Text            =   "User: None       "
            TextSave        =   "User: None       "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "3:41 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2640
      Top             =   1065
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   3600
      Top             =   1770
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CEA
            Key             =   "Validate Transactions"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":113C
            Key             =   "Pre-Edit File"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1456
            Key             =   "Post Transactions"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2628
            Key             =   "State Treasurer"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2942
            Key             =   "Load FUNB File"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D94
            Key             =   "Balance Direct Deposit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34A0
            Key             =   "Reports"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38F2
            Key             =   "Transaction"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D44
            Key             =   "View Transaction Group"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4196
            Key             =   "Stop Transaction Group"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":45E8
            Key             =   "Search Patient Account"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A3A
            Key             =   "Print Reports"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":50B4
            Key             =   "Patient Account Info"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53CE
            Key             =   "Bulk Interest Account"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":56E8
            Key             =   "Modify Allowances"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A02
            Key             =   "Std Allowance Timeframes"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":62DC
            Key             =   "Institutions"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6BB6
            Key             =   "Purchase Types"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7490
            Key             =   "Security"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7D6A
            Key             =   "State Codes"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8644
            Key             =   "Transaction Categories"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8F1E
            Key             =   "Deposit Sequence No1"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9AF0
            Key             =   "Deposit Sequence No"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9C10
            Key             =   "NCAS Accounting"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A4EA
            Key             =   "Vendor ID Number2"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B76C
            Key             =   "Income Source Types"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BA86
            Key             =   "DDS Configuration"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BF12
            Key             =   "Regions"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolBarSmallIcons 
      Left            =   2385
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C366
            Key             =   "Login"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C7B8
            Key             =   "Print_Reports"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CCFA
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D23C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D77E
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DCC0
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E202
            Key             =   "Transaction"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E654
            Key             =   "View_Transaction_Group"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EAA6
            Key             =   "Stop_Transaction_Group"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EEF8
            Key             =   "Search_Patient_Account"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F34A
            Key             =   "Calculator"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F79C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FAB6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4185
      Top             =   3015
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   197
      ImageHeight     =   170
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FDD0
            Key             =   "CkCard"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14094
            Key             =   "Help02"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":150AD
            Key             =   "Death"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15569
            Key             =   "Cashbook"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15687
            Key             =   "Eye"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19E8C
            Key             =   "Devil"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A2B4
            Key             =   "Skull"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A735
            Key             =   "Send"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":244BA
            Key             =   "Email"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25509
            Key             =   "Moneybut"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":256C9
            Key             =   "Safe"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25827
            Key             =   "Signboo"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2599A
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25BE4
            Key             =   "Yes"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25D65
            Key             =   "Direct-Deposit"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29875
            Key             =   "monBag"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C7B8
            Key             =   "Money"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F73A
            Key             =   "rotateskull"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":319A2
            Key             =   "Signat"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":348B7
            Key             =   "Notes"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34D13
            Key             =   "Yes2b"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34E1E
            Key             =   "Email3D"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B09E
            Key             =   "Comments1"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B243
            Key             =   "DSN"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B363
            Key             =   "Monitor1"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B4C9
            Key             =   "Happy"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuChangePassword 
         Caption         =   "&Change Password..."
      End
      Begin VB.Menu mnuStepAway 
         Caption         =   "&Step Away from Computer"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "Vertical &Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuMainGrp 
      Caption         =   "&Main Functions"
      Visible         =   0   'False
      Begin VB.Menu mnuMain 
         Caption         =   "Patient Account..."
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Bulk Interest Account..."
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Modify Patient Allowance..."
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Account Inquiry..."
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Enter Transactions..."
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Reports..."
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Reports"
         Index           =   6
      End
      Begin VB.Menu mnuMainBlank 
         Caption         =   ""
      End
   End
   Begin VB.Menu mnuMaintGrp 
      Caption         =   "Mai&ntenance"
      Visible         =   0   'False
      Begin VB.Menu mnuMaint 
         Caption         =   "Standard Allowance Timeframe..."
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaint 
         Caption         =   "Institution..."
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaint 
         Caption         =   "NCAS Accounting Data..."
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaint 
         Caption         =   "Purchase Types..."
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaint 
         Caption         =   "Transaction Category..."
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaint 
         Caption         =   "Transaction Codes..."
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaint 
         Caption         =   "Threshold Amounts..."
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaint 
         Caption         =   "State Codes..."
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaint 
         Caption         =   "&Security..."
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaintBlank 
         Caption         =   ""
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About... "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ********************************************************************************
' * Description: This is the Main MDI Form  this is the driver for all
' *              functions in the program
' *
' *
' *
' * Revisions:
' *  2/24/99    fml Added comments.
' *  3/23/99    aschulte - Added error handling to Login As New User.  The function was not
' *             performing properly when selecting Login As New user from a Maintenance screen.
' *  3/23/99    aschulte - When you login as a new user, the institution drop down list was not refreshed
' *  3/29/99    aschulte - Changed the behavior of the institution drop down when changing users
'********************************************************************************

Public WithEvents objDOS As DosOutputs
Attribute objDOS.VB_VarHelpID = -1
Public msDosOutput As String

' Mod CONSTANTS
Private Const MODULE As String = "Main Screen - "

' Mod ENUMS
Private Enum SMTP_State
    MAIL_CONNECT
    MAIL_HELO
    MAIL_FROM
    MAIL_RCPTTO
    MAIL_DATA
    MAIL_DOT
    MAIL_QUIT
    MAIL_ERROR
End Enum


' Mod TYPES


' Mod DECLARES


' Mod VARIABLES

Private m_State As SMTP_State
Private moEmailMsg As clsEmailMessage

Dim bViewToolbarText As Boolean


Private Sub MDIForm_Activate()
    If Command() = "/d" Then
        MsgBox "Activating Main Form"
    End If
    Call MDIForm_Resize
End Sub

Private Sub MDIForm_Load()
'********************************************************************************
'* Name: MDIForm_Load
'*
'* Description: Sets up the toolbars, security, menus, outlook bar and status bar for use
'* Parameters:
'* Created: 2/26/99 9:51:33 AM
'********************************************************************************
   
    On Error GoTo MDIFormLoadErr
    If Command() = "/d" Then
        MsgBox "Loading Dos Outputs"
    End If
    
    Set objDOS = New DosOutputs
    Hourglass True
    If Command() = "/d" Then
        MsgBox "Sending message to Size SubClass"
    End If
    
    SizeSubclass.hwnd = Me.hwnd
    SizeSubclass.Messages(WM_GETMINMAXINFO) = True
    
    'Remove All the toolbars to start from scratch
    RemoveToolbar "Main"
    If Command() = "/d" Then
        MsgBox "Adding Main Toolbar"
    End If
    'Add the main toolbar
    AddToolbar "Main"
    
    'Get the position of the main screen from the registry
    Me.Left = GetSetting(App.Title, "Startup", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Startup", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Startup", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Startup", "MainHeight", 6500)
    Me.WindowState = GetSetting(App.Title, "Startup", "MainState", 2)
    
    'Dislay the state seal in the middle of the screen
    frmSeal.Show
    
    
    If Command() = "/d" Then
        MsgBox "Filling Menu Options"
    End If
    
    'Fill the outlook bar an menu bar depending on the users permissions
    FillMenuOptions
   
    If Command() = "/d" Then
        MsgBox "SetMainToolBar"
    End If
   
     
    SetMainToolbar False
    If Command() = "/d" Then
        MsgBox "Finished Loading Main Form"
    End If
    
     
 
    Exit Sub
MDIFormLoadErr:
    MsgBox Err.LastDllError & " - " & ErrorDescriptionDLL(Err.LastDllError)
    MsgBox Error, vbCritical
    
    ExitApp
    
End Sub



Private Sub FillMenuOptions()

On Error GoTo FillMenuOptionsErr

'********************************************************************************
'* Name: FillMenuOptions
'*
'* Description: This function checks the permissions of viewing the main forms
'*              and adds to the outlook bar and to the menu bar
'* Parameters:
'* Created: 2/26/99 11:14:06 AM
'********************************************************************************

    Picture1.Visible = False
    
    AddVerticalMenuItem "frmLoadFUNB", "Load FUNB File", "&Load FUNB File", 1
    AddVerticalMenuItem "frmEditDeposits", "Validate Transactions", "&Validate Transactions", 1
    AddVerticalMenuItem "frmPreEditMain", "Pre-Edit File", "Pre-&Edit File", 1
    AddVerticalMenuItem "frmPostDeposits", "Post Transactions", "&Post Transactions", 1
    AddVerticalMenuItem "frmBalance", "Balance Direct Deposit", "&Balance Direct Deposit", 1
    AddVerticalMenuItem "frmStateTreasurer", "State Treasurer", "&State Treasurer", 1
    AddVerticalMenuItem "frmMainReports", "Reports", "&Reports", 1

    If mnuMainGrp.Visible = True Then
        mnuMainBlank.Visible = False
    End If
    

    AddVerticalMenuItem "frmInstitutions", "Institutions", "&Institutions", 2
    AddVerticalMenuItem "frmDDSUsers", "Security", "&Security", 2
    AddVerticalMenuItem "frmIncomeSourceTypes", "Income Source Types", "&Income Source Types", 2
    AddVerticalMenuItem "frmRegions", "Regions", "&Regions", 2
    AddVerticalMenuItem "frmDDConfiguration", "DDS Configuration", "&DDS Configuration", 2

    If mnuMaintGrp.Visible = True Then
        mnuMaintBlank.Visible = False
    End If

    If VerticalMenu1.MenusMax > 0 Then
        VerticalMenu1.MenuCur = 1
    End If
    Picture1.Visible = True

Xit:
    Exit Sub

FillMenuOptionsErr:
    ShowUnexpectedError MODULE + "FillMenuOptions", Err
    Resume Xit

End Sub
Private Sub AddVerticalMenuItem(ByVal sForm As String, ByVal sIconKey As String, ByVal sMenuCaption As String, ByVal iMenu As Integer)

On Error GoTo AddVerticalMenuItemErr

'********************************************************************************
'* Name: AddVerticalMenuItem
'*
'* Description: Adds one item to the menu bar and to the vertical outlook menu bar
'* Parameters:
'* Created: 2/26/99 11:15:57 AM
'********************************************************************************

    If iMenu = 1 Then
        If mnuMainGrp.Visible = False Then
            'Need to add the Main Functions Group
            mnuMainGrp.Visible = True
            VerticalMenu1.MenusMax = VerticalMenu1.MenusMax + 1
            VerticalMenu1.MenuCur = iMenu
            VerticalMenu1.MenuCaption = "Main Functions"
        Else
            VerticalMenu1.MenuCur = iMenu
            VerticalMenu1.MenuItemsMax = VerticalMenu1.MenuItemsMax + 1
        End If
        mnuMain(VerticalMenu1.MenuItemsMax - 1).Caption = sMenuCaption
        mnuMain(VerticalMenu1.MenuItemsMax - 1).Visible = True
    End If
    
    If iMenu = 2 Then
        If mnuMaintGrp.Visible = False Then
            If mnuMainGrp.Visible = False Then
                iMenu = 1
                VerticalMenu1.MenusMax = 1
            Else
                VerticalMenu1.MenusMax = VerticalMenu1.MenusMax + 1
            End If
            mnuMaintGrp.Visible = True
            VerticalMenu1.MenuCur = iMenu
            VerticalMenu1.MenuCaption = "Maintenance"
        Else
            If mnuMainGrp.Visible = False Then
                iMenu = 1
            End If
            VerticalMenu1.MenuCur = iMenu
            VerticalMenu1.MenuItemsMax = VerticalMenu1.MenuItemsMax + 1
        End If
        mnuMaint(VerticalMenu1.MenuItemsMax - 1).Caption = sMenuCaption
        mnuMaint(VerticalMenu1.MenuItemsMax - 1).Visible = True
    End If
    
    VerticalMenu1.MenuItemCur = VerticalMenu1.MenuItemsMax
    VerticalMenu1.MenuItemCaption = sIconKey
    VerticalMenu1.MenuItemKey = sForm
    Set VerticalMenu1.MenuItemIcon = imlToolbarIcons.ListImages(sIconKey).Picture


Xit:
    Exit Sub

AddVerticalMenuItemErr:
    ShowUnexpectedError MODULE + "AddVerticalMenuItem", Err
    Resume Xit


End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Unload Me
    Cancel = True
    Set objDOS = Nothing
    ExitApp
End Sub

Private Sub MDIForm_Resize()
' Keep the vertical menu in proportion w/ MDI Form
    On Error Resume Next
    
    VerticalMenu1.Height = Me.ScaleHeight - 150
'There is a problem with the control when for resizes
'By changing the visible property off and on we can
'bypass the fault in the ocx
    VerticalMenu1.Visible = False
    VerticalMenu1.Visible = True
    If Picture1.Visible = True Then
        frmSeal.Left = ((Me.Width - Picture1.Width) / 2) - (frmSeal.Width / 2)
    Else
        frmSeal.Left = ((Me.Width / 2) - (frmSeal.Width / 2))
    End If
    
    frmSeal.Top = ((Me.Height - 1500) / 2) - (frmSeal.Height / 2)
    
    On Error GoTo 0

End Sub


Private Sub mnuChangePassword_Click()

    Dim fLogin As New frmLogin
    fLogin.iMode = ChangePassword
    fLogin.Show vbModal
    Unload fLogin
    Set fLogin = Nothing

End Sub

Private Sub mnuFileLogin_Click()
    
    Call LoginAsNewUser
    
End Sub

Private Sub mnuMain_Click(Index As Integer)
    
    Call VerticalMenu1_MenuItemClick(1, CLng(Index + 1))

End Sub



Private Sub mnuMaint_Click(Index As Integer)

    If VerticalMenu1.MenusMax = 1 Then
        'We do not have the Main screen functions up so call menu 1
        Call VerticalMenu1_MenuItemClick(1, CLng(Index + 1))
    Else
        Call VerticalMenu1_MenuItemClick(2, CLng(Index + 1))
    End If
     
End Sub

Private Sub mnuStepAway_Click()
    Call StepAway
End Sub


Private Sub SizeSubclass_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)
    Dim MinMax As MINMAXINFO

    If Msg = WM_GETMINMAXINFO Then
        'Copy to our local MinMax variable
        CopyMemory MinMax, ByVal lParam, Len(MinMax)
        'Set minimum/maximum tracking size
        MinMax.ptMinTrackSize.X = 802
        MinMax.ptMinTrackSize.Y = 572
        'Copy data back to Windows
        CopyMemory ByVal lParam, MinMax, Len(MinMax)
        Result = 0
    End If

End Sub






Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub







Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    Call MDIForm_Resize
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    Picture1.Visible = mnuViewToolbar.Checked
    Call MDIForm_Resize
End Sub

Private Sub mnuEditCopy_Click()

On Error GoTo mnuEditCopy_ClickErr

   
   Clipboard.Clear
   If TypeOf Screen.ActiveControl Is ComboBox Then
      Clipboard.SetText Screen.ActiveControl.Text
   ElseIf TypeOf Screen.ActiveControl Is DateSelector Then
      Clipboard.SetText Screen.ActiveControl.Text
   ElseIf TypeOf Screen.ActiveControl Is ListBox Then
      Clipboard.SetText Screen.ActiveControl.Text
   ElseIf TypeOf Screen.ActiveControl Is ListView Then
      Clipboard.SetText Screen.ActiveControl.SelectedItem.Text
   Else
      On Error Resume Next
      Clipboard.SetText Screen.ActiveControl.SelText
   End If

Xit:
    Exit Sub

mnuEditCopy_ClickErr:
    ShowUnexpectedError MODULE + "mnuEditCopy_Click", Err
    Resume Xit


End Sub

Private Sub mnuEditCut_Click()

On Error GoTo mnuEditCut_ClickErr

   ' First do the same as a copy.
   mnuEditCopy_Click
   ' Now clear contents of active control.
   If TypeOf Screen.ActiveControl Is ComboBox Then
      Screen.ActiveControl.Text = ""
   ElseIf TypeOf Screen.ActiveControl Is DateSelector Then
      Screen.ActiveControl.Text = ""
   ElseIf TypeOf Screen.ActiveControl Is ListView Then
      'Do Nothing
   Else
      On Error Resume Next
      Screen.ActiveControl.SelText = ""
   End If

Xit:
    Exit Sub

mnuEditCut_ClickErr:
    ShowUnexpectedError MODULE + "mnuEditCut_Click", Err
    Resume Xit


End Sub

Private Sub mnuEditPaste_Click()

On Error GoTo mnuEditPaste_ClickErr

   
   If TypeOf Screen.ActiveControl Is ComboBox Then
      Screen.ActiveControl.Text = Clipboard.GetText()
   ElseIf TypeOf Screen.ActiveControl Is DateSelector Then
      Screen.ActiveControl.Text = Clipboard.GetText()
   Else
      On Error Resume Next
      Screen.ActiveControl.SelText = Clipboard.GetText()
   End If

Xit:
    Exit Sub

mnuEditPaste_ClickErr:
    ShowUnexpectedError MODULE + "mnuEditPaste_Click", Err
    Resume Xit


End Sub





Private Sub mnuFileExit_Click()
    'unload the form
    ExitApp

End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    fMainForm.ActiveForm.PrintReport

End Sub


Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .Flags = cdlPDPrintSetup
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    
    Select Case Button.key
    
    Case "Login"
        fMainForm.SetFocus
        Call LoginAsNewUser
    Case "StepAway"
        Call StepAway
        
    Case "Cut"
        mnuEditCut_Click
    Case "Copy"
        mnuEditCopy_Click
    Case "Paste"
        mnuEditPaste_Click
    Case "Calculator"
        Shell "calc.exe", vbNormalFocus
    End Select
End Sub

Private Sub VerticalMenu1_MenuItemClick(MenuNumber As Long, MenuItem As Long)
    
    VerticalMenu1.MenuCur = MenuNumber
    VerticalMenu1.MenuItemCur = MenuItem
    
    If fMainForm.ActiveForm.Name = VerticalMenu1.MenuItemKey Then
        ' User clicked the same Screen that is currently opened
        Exit Sub
    End If
    
    Hourglass True
    
    Select Case VerticalMenu1.MenuCaption
    Case "Main Functions"
        Select Case VerticalMenu1.MenuItemKey
        Case "frmLoadFUNB"
            frmLoadFUNB.Show
            frmLoadFUNB.SetFocus
        Case "frmPreEditMain"
            frmPreEditMain.Show
            frmPreEditMain.SetFocus
        Case "frmBalance"
            frmBalance.Show
            frmBalance.SetFocus
        Case "frmEditDeposits"
            frmEditDeposits.Show
            frmEditDeposits.SetFocus
        Case "frmPostDeposits"
            frmPostDeposits.Show
            frmPostDeposits.SetFocus
        Case "frmMainReports"
            frmMainReports.Show
            frmMainReports.SetFocus
        Case "frmStateTreasurer"
            frmStateTreasurer.Show
            frmStateTreasurer.SetFocus
        End Select
    Case "Maintenance"
        Select Case VerticalMenu1.MenuItemKey
        Case "frmInstitutions"
            frmInstitutions.Show
            frmInstitutions.SetFocus
        
        Case "frmDDSUsers"
            frmDDSUsers.Show
            frmDDSUsers.SetFocus
            
       Case "frmIncomeSourceTypes"
            frmIncomeSourceTypes.Show
            frmIncomeSourceTypes.SetFocus
            
       Case "frmRegions"
            frmRegions.Show
            frmRegions.SetFocus
       
       Case "frmDDConfiguration"
           frmDDConfiguration.Show
           frmDDConfiguration.SetFocus
            
        End Select
    End Select
    Hourglass False

End Sub

Public Sub RemoveToolbar(sToolbar As String)
    On Error Resume Next
    CoolBar.Bands.Remove sToolbar

End Sub

Public Sub AddToolbar(ByVal sToolbar As String)
    On Error Resume Next
    Select Case sToolbar
    Case "Main"
        CoolBar.Bands.Add CoolBar.Bands.Count + 1, "Main", "", , False, tbrMain, True
        CoolBar.Bands(sToolbar).Width = tbrMain.Width + 750
    End Select
End Sub

Public Sub SetMainToolbar(ByVal bEditMenuEnabled As Boolean)

    'Set the apropriate buttons for the main toolbar when no screens are displayed
    tbrMain.Buttons("Cut").Enabled = bEditMenuEnabled
    mnuEditCut.Enabled = bEditMenuEnabled
    
    tbrMain.Buttons("Copy").Enabled = bEditMenuEnabled
    mnuEditCopy.Enabled = bEditMenuEnabled
    
    tbrMain.Buttons("Paste").Enabled = bEditMenuEnabled
    mnuEditPaste.Enabled = bEditMenuEnabled
    
End Sub

Private Sub LoginAsNewUser()

'********************************************************************************
'* Name: LoginAsNewUser
'*
'* Description: Logic to login user as a different user
'* Parameters:
'* Created: 3/23/99 9:57:37 AM
'********************************************************************************

    On Error Resume Next
    
    Dim frm As Form
    For Each frm In Forms
        Select Case frm.Name
        Case "frmMain", "frmSeal"
            'Do Nothing
        Case Else
            If frm.CheckIfOkToUnload = True Then
                Unload frm
            Else
                Exit Sub
            End If
        End Select
    Next
    

    On Error GoTo LoginAsNewUserErr
    
    gcnDDS.Close
    
    Dim fLogin As New frmLogin
    Dim ix As Integer
    'Save which institution the user was working on
    If gobjLoginInfo.InstitutionID <> vbNullString Then
        SaveSetting App.Title, gobjLoginInfo.UserId, "Institution", gobjLoginInfo.InstitutionID
    End If

    fLogin.iMode = LOGINNEWUSER
    DoEvents
    fLogin.Show vbModal
    If Not fLogin.OK Then
        'Login using old settings
        gcnDDS.Open gobjLoginInfo.ConnectString
    Else
        'Login successfully. Reset the Trans Group Number
        gobjLoginInfo.TransGroupNum = 0
        gobjLoginInfo.InstitutionID = ""
        gobjLoginInfo.InstitutionName = ""
        VerticalMenu1.MenuCur = 1
        VerticalMenu1.MenuItemsMax = 1
        VerticalMenu1.MenuCur = 2
        VerticalMenu1.MenuItemsMax = 1
        VerticalMenu1.MenusMax = 0
        VerticalMenu1.Refresh
        mnuMainBlank.Visible = True
        mnuMaintBlank.Visible = True
        For ix = 0 To mnuMain.Count - 1
            mnuMain(ix).Visible = False
        Next ix
        For ix = 0 To mnuMaint.Count - 1
            mnuMaint(ix).Visible = False
        Next ix
        mnuMainGrp.Visible = False
        mnuMaintGrp.Visible = False
        FillMenuOptions
        If VerticalMenu1.MenusMax = 1 And VerticalMenu1.MenuCaption = "Maintenance" Then
            DoEvents
            SetMainToolbar False
        End If
    End If
    
    Unload fLogin
    
    fMainForm.sbStatusBar.Panels(1).Text = "User: " & gobjLoginInfo.UserId & "      "
    fMainForm.sbStatusBar.Panels(3).Text = Format$(Now, "MM/dd/yyyy")
    gobjLoginInfo.MRN = vbNullString

Xit:
    Exit Sub

LoginAsNewUserErr:
    ShowUnexpectedError MODULE + "LoginAsNewUser", Err
    Resume Xit


End Sub

Private Sub StepAway()

On Error GoTo StepAwayErr

'********************************************************************************
'* Name: StepAway
'*
'* Description:
'* Parameters:
'* Created: 2/25/99 1:10:09 PM
'********************************************************************************

    
    Dim ctl As Control
    Dim frm As Form
    Dim ix As Integer
    If tbrMain.Buttons("StepAway").ToolTipText = "Step Away from Computer" Then
        mnuStepAway.Caption = "Resume Work"
        tbrMain.Buttons("StepAway").ToolTipText = "Resume work"
        For Each ctl In fMainForm.Controls
            ctl.Enabled = False
        Next
        
        CoolBar.Enabled = True
        tbrMain.Enabled = True
        For ix = 1 To tbrMain.Buttons.Count
            tbrMain.Buttons(ix).Enabled = False
        Next ix
        tbrMain.Buttons("StepAway").Enabled = True
        mnuFile.Enabled = True
        mnuStepAway.Enabled = True
        mnuFileExit.Enabled = True
        
        Picture1.Visible = False
        Call MDIForm_Resize
        For Each frm In Forms
            Select Case frm.Name
            Case "frmMain", "frmSeal"
                'Do Nothing
            Case Else
                frm.Hide
            End Select
        Next
    Else
        'Try to restore the screen
        Dim fLogin As New frmLogin
        fLogin.iMode = STEPAWAYMODE
        fLogin.Show vbModal
        Hourglass False
        If Not fLogin.OK Then
            'Login Failed so exit app
            MsgBox " Incorrect password.", vbExclamation
            Unload fLogin
            Exit Sub
        End If
        Unload fLogin
        
        tbrMain.Buttons("StepAway").ToolTipText = "Step Away from Computer"
        mnuStepAway.Caption = "Step Away from Computer"
        Picture1.Visible = True
        
        For Each ctl In fMainForm.Controls
            ctl.Enabled = True
        Next
    
        For ix = 1 To tbrMain.Buttons.Count
            tbrMain.Buttons(ix).Enabled = True
        Next ix
        
        For Each frm In Forms
            Select Case frm.Name
            Case "frmMain", "frmSeal"
                'Do Nothing
            Case Else
                frm.Show
            End Select
        Next
    End If

        SetMainToolbar False
        

Xit:
    Exit Sub

StepAwayErr:
    If Err = 438 Or Err = 387 Then
        Resume Next
    End If
    
    ShowUnexpectedError MODULE + "StepAway", Err
    Resume Xit


End Sub

Public Function SendEmail(ByRef oEmailMsgPassed As clsEmailMessage) As Boolean
    
    Dim i As Integer
    Set moEmailMsg = Nothing
    Set moEmailMsg = oEmailMsgPassed
    
    If WinsockEmail.SocketHandle <> -1 Then
        WinsockEmail.Close
    End If
    m_State = MAIL_CONNECT
    WinsockEmail.Connect Trim$(moEmailMsg.SMTPServer), 25
    
    'Wait here until the email has been sent
    Do Until m_State = MAIL_QUIT Or m_State = MAIL_ERROR
        DoEvents
    Loop
    
    If m_State = MAIL_QUIT Then
        SendEmail = True
    Else
        SendEmail = False
    End If
   
    
End Function

Private Sub WinsockEmail_DataArrival(ByVal bytesTotal As Long)

    Dim strServerResponse   As String
    Dim strResponseCode     As String
    Dim strDataToSend       As String
    Dim strMessage As String
    '
    'Retrive data from winsock buffer
    '
    WinsockEmail.GetData strServerResponse
    '
    Debug.Print strServerResponse
    '
    'Get server response code (first three symbols)
    '
    strResponseCode = Left(strServerResponse, 3)
    '
    'Only these three codes tell us that previous
    'command accepted successfully and we can go on
    '
    If strResponseCode = "250" Or _
       strResponseCode = "220" Or _
       strResponseCode = "354" Then
       
        Select Case m_State
            Case MAIL_CONNECT
                'Change current state of the session
                m_State = MAIL_HELO
                '
                'Remove blank spaces
                strDataToSend = Trim$(moEmailMsg.From) & vbCrLf
                '
                'Retrieve mailbox name from e-mail address
                strDataToSend = Left$(strDataToSend, _
                                InStr(1, strDataToSend, "@") - 1)
                'Send HELO command to the server
                WinsockEmail.SendData "HELO " & strDataToSend & vbCrLf
                '
                Debug.Print "HELO " & strDataToSend
                '
            Case MAIL_HELO
                '
                'Change current state of the session
                m_State = MAIL_FROM
                '
                'Send MAIL FROM command to the server
                WinsockEmail.SendData "MAIL FROM:<" & Trim$(moEmailMsg.From) & ">" & vbCrLf
                '
                Debug.Print "MAIL FROM:" & Trim$(moEmailMsg.From)
                '
            Case MAIL_FROM
                '
                moEmailMsg.RecipientToSend = moEmailMsg.RecipientToSend + 1
                If moEmailMsg.RecipientToSend = moEmailMsg.Recipients.Count Then
                    'Change current state of the session
                    m_State = MAIL_RCPTTO
                End If
                '
                'Send RCPT TO command to the server
                WinsockEmail.SendData "RCPT TO:<" & moEmailMsg.Recipients(moEmailMsg.RecipientToSend) & ">" & vbCrLf
                '
                'Debug.Print "RCPT TO:" & Trim$(txtRecipient)
                '
            Case MAIL_RCPTTO
                '
                'Change current state of the session
                m_State = MAIL_DATA
                '
                'Send DATA command to the server
                WinsockEmail.SendData "DATA" & vbCrLf
                '
                Debug.Print "DATA"
                '
            Case MAIL_DATA
                '
                'Change current state of the session
                m_State = MAIL_DOT
                '
                'So now we are sending a message body
                'Each line of text must be completed with
                'linefeed symbol (Chr$(10) or vbLf) not with vbCrLf
                '
                'Send Subject line
                WinsockEmail.SendData "From:" & Trim$(moEmailMsg.From) & vbCrLf
                WinsockEmail.SendData "To:" & Trim$(moEmailMsg.ToRecipient) & vbCrLf
                If moEmailMsg.CCRecipient <> "" Then
                    WinsockEmail.SendData "Cc:" & moEmailMsg.CCRecipient & vbCrLf
                End If
                If moEmailMsg.BCCRecipient <> "" Then
                    WinsockEmail.SendData "Bcc:" & moEmailMsg.BCCRecipient & vbCrLf
                End If
                
                WinsockEmail.SendData "Subject:" & moEmailMsg.Subject & vbCrLf
                WinsockEmail.SendData "MIME-Version: 1.0" & vbCrLf
                If moEmailMsg.AttachmentCount = 0 Then
                    WinsockEmail.SendData "Content-Type: text/plain; charset=us-ascii" & vbCrLf
                    WinsockEmail.SendData "Content-Transfer-Encoding: 7bit" & vbCrLf & vbLf
                    strMessage = moEmailMsg.Message & vbCrLf & vbCrLf
                    WinsockEmail.SendData strMessage
                Else
                    WinsockEmail.SendData "Content-Type: MULTIPART/MIXED; BOUNDARY=" & """" & moEmailMsg.Boundary & """" & vbCrLf & vbLf
                    WinsockEmail.SendData "This is a multi-part message in MIME format." & vbCrLf
                    WinsockEmail.SendData "--" & moEmailMsg.Boundary & vbCrLf
                    WinsockEmail.SendData "Content-Type: text/plain; charset=us-ascii" & vbCrLf
                    WinsockEmail.SendData "Content-Transfer-Encoding: 7bit" & vbCrLf & vbLf
                    WinsockEmail.SendData moEmailMsg.Message & vbCrLf & vbLf
                    strMessage = moEmailMsg.GetEncodedAttachments
             '
                    Dim varLines    As Variant
                    Dim varLine     As Variant
                '
                    'Parse message to get lines (for VB6 only)
                    varLines = Split(strMessage, vbCrLf)
                    'clear memory
                    strMessage = ""
                    '
                    'Send each line of the message
                    For Each varLine In varLines
                        WinsockEmail.SendData CStr(varLine) & vbLf
                    '
                        Debug.Print CStr(varLine)
                    Next
                End If
                '
                'Send a dot symbol to inform server
                'that sending of message comleted
                WinsockEmail.SendData vbCrLf & "." & vbCrLf
                '
                Debug.Print "."
                '
            Case MAIL_DOT
                'Change current state of the session
                m_State = MAIL_QUIT
                '
                'Send QUIT command to the server
                WinsockEmail.SendData "QUIT" & vbCrLf
                '
                Debug.Print "QUIT"
            Case MAIL_QUIT
                '
                'Close connection
                WinsockEmail.Close
                Set moEmailMsg = Nothing
                '
        End Select
       
    Else
        '
        'If we are here server replied with
        'unacceptable respose code therefore we need
        'close connection and inform user about problem
        '
        WinsockEmail.Close
        Set moEmailMsg = Nothing
        '
        If Not m_State = MAIL_QUIT Then
            MsgBox "SMTP Error: " & strServerResponse, _
                    vbInformation, "SMTP Error"
            m_State = MAIL_ERROR
        End If
        '
    End If
    
End Sub

Private Sub WinsockEmail_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    MsgBox "Winsock Error number " & Number & vbCrLf & _
            Description, vbExclamation, "Winsock Error"

End Sub


