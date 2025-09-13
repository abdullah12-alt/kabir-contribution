VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBrowser 
   Caption         =   "Acquire FUNB File"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar barMain 
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   7200
      Visible         =   0   'False
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   7080
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   847
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reload"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "preview"
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "preview"
                  Text            =   "Print Preview"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "page"
                  Text            =   "Page Setup"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "popup"
            Object.ToolTipText     =   "Popup Blocker"
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "find"
            Object.ToolTipText     =   "Find && Replace"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save File To Disk"
            ImageIndex      =   16
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   10005
      TabIndex        =   1
      Top             =   480
      Width           =   10005
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go!"
         Default         =   -1  'True
         Height          =   255
         Left            =   9240
         TabIndex        =   3
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtURL 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   9015
      End
      Begin MSComctlLib.ImageList imlToolbarIcons 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   16
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":0FDC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":16EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":1E00
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":2512
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":2C24
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":3336
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":3C10
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":428A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":4904
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":4FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":58D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":61B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":68AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":7186
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBrowser.frx":7880
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin SHDocVwCtl.WebBrowser webMain 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   9975
      ExtentX         =   17595
      ExtentY         =   10398
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Sample code for 'Mastering Internet Explorer: Using the DHTML Document Object Model'

Dim WithEvents objDocument As HTMLDocument
Attribute objDocument.VB_VarHelpID = -1

Private Enum BrowserNavConstants
    navOpenInNewWindow = &H1
    navNoHistory = &H2
    navNoReadFromCache = &H4
    navNoWriteToCache = &H8
    navAllowAutosearch = &H10
    navBrowserBar = &H20
    navHyperlink = &H40
End Enum

Public FileName As String


Private Sub WebNavigate(URL As String, Optional Flags, Optional Target, Optional PostData, Optional Headers)
webMain.Navigate URL, Flags, Target, PostData, Headers
End Sub

Private Function CheckPopup() As Boolean
On Error GoTo errHandle

If webMain.Document.activeElement.tagName = "BODY" Or _
    webMain.Document.activeElement.tagName = "IFRAME" Then
    CheckPopup = False
Else
    CheckPopup = True
End If

Exit Function

errHandle:
If Err.Number = 91 Then
    CheckPopup = False
End If
End Function

Private Function CheckLock() As Boolean
On Error Resume Next
Dim strRestrict As Variant, i As Integer, intCount As Integer, objRange As Object
strRestrict = Array("amateur", "cock", "penis", "tit", "mature", "anal", "oral", "vaginal", "swallow", "blowjob", "sex", "porn", "hot", "babe")

Set objRange = webMain.Document.body.createTextRange

'This error is thrown when no page has been loaded yet and there is no
'body to use the createTextRange method on.
If Err.Number = 91 Then
    CheckLock = True
    Exit Function
End If

'Searching for text in the page
For i = 0 To UBound(strRestrict)
    If objRange.findText(strRestrict(i)) = True Then
        intCount = intCount + 1
    End If
Next i

If intCount > 4 Then
    CheckLock = False
Else
    CheckLock = True
End If

Set objRange = Nothing
End Function

Private Function CheckVBWM(ByVal URL As String) As Boolean
If InStr(1, Split(URL, "/")(2), "vbwm.com") = 0 Then
    CheckVBWM = False
Else
    CheckVBWM = True
End If
End Function

Private Sub CheckCommands()
If webMain.QueryStatusWB(OLECMDID_PRINT) = 0 Then
    tlbMain.Buttons.Item("print").Enabled = False
Else
    tlbMain.Buttons.Item("print").Enabled = True
End If

If webMain.QueryStatusWB(OLECMDID_PRINTPREVIEW) = 0 Then
    tlbMain.Buttons.Item("preview").Enabled = False
    tlbMain.Buttons.Item("preview").ButtonMenus.Item("preview").Enabled = False
Else
    tlbMain.Buttons.Item("preview").Enabled = True
    tlbMain.Buttons.Item("preview").ButtonMenus.Item("preview").Enabled = True
End If

If webMain.QueryStatusWB(OLECMDID_PAGESETUP) = 0 Then
    tlbMain.Buttons.Item("preview").ButtonMenus.Item("page").Enabled = False
Else
    tlbMain.Buttons.Item("preview").ButtonMenus.Item("page").Enabled = True
End If

If webMain.Document Is Nothing Then
    tlbMain.Buttons.Item("find").Enabled = False
Else
    tlbMain.Buttons.Item("find").Enabled = True
End If
End Sub

Private Sub cmdGo_Click()
WebNavigate txtURL.Text
End Sub

Private Sub Form_Load()
Dim sText As String
sText = ReadIniFile(App.Path & "\" & App.EXEName & ".ini", "Startup", "InitialAddress")
txtURL.Text = sText
WebNavigate sText
End Sub

Private Sub Form_Resize()
webMain.Left = 0
webMain.Top = picMain.Height + tlbMain.Height
webMain.Width = Me.Width - 100
webMain.Height = Me.Height - picMain.Height - tlbMain.Height - stbMain.Height - 400
barMain.Top = Me.Height - stbMain.Height - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objDocument = Nothing
End Sub



Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim ts As clsTextFile

Select Case Button.key
    Case "back"
        webMain.GoBack
    Case "forward"
        webMain.GoForward
    Case "stop"
        webMain.Stop
        Button.Enabled = False
    Case "reload"
        webMain.Refresh
    Case "home"
        webMain.GoHome
    Case "search"
        webMain.GoSearch
    Case "preview"
        webMain.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
    Case "print"
        webMain.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
    Case "lock"
        If Button.value = tbrPressed Then
            tlbMain.Buttons.Item("vbwm").value = tbrUnpressed
        Else
            tlbMain.Buttons.Item("vbwm").value = tbrUnpressed
        End If
    Case "vbwm"
        If Button.value = tbrPressed Then
            tlbMain.Buttons.Item("lock").value = tbrUnpressed
        Else
            tlbMain.Buttons.Item("lock").value = tbrUnpressed
        End If
    Case "find"
        frmBrowserFind.Show 1
    Case "save"
        With fMainForm.dlgCommonDialog
        .DialogTitle = "Save FUNB Detail File"
        .InitDir = GetSetting(App.EXEName, "Settings", "LoadInitDir", "")
        .FileName = FileName
        .CancelError = True
        On Error Resume Next
        .ShowSave
        If Err = 0 Then
            'AS - 2/16/2014 Replaced FileSystemObject with clsTextFile
            'Dim fso As New FileSystemObject
            'Dim ts As TextStream
            If FileExists(.FileName) = True Then
                MsgBox "File already exists"
            Else
                Set ts = New clsTextFile
                ts.OpenFile .FileName, OUTPUT_NEW
                ts.WriteLine webMain.Document.body.innerText
                ts.CloseFile
            End If
            Set ts = Nothing
            'Set fso = Nothing
        End If
        End With
    End Select
End Sub

Private Sub tlbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.key
    Case "preview"
        webMain.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
    Case "page"
        webMain.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
End Select
End Sub

Private Sub txtURL_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then WebNavigate txtURL.Text
End Sub

Private Sub webMain_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
tlbMain.Buttons.Item("stop").Enabled = True
CheckCommands

End Sub

Private Sub webMain_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
Select Case Command
    Case 1 'Forward
        tlbMain.Buttons.Item("forward").Enabled = Enable
    Case 2 'Back
        tlbMain.Buttons.Item("back").Enabled = Enable
End Select
End Sub

Private Sub webMain_DocumentComplete(ByVal pDisp As Object, URL As Variant)
tlbMain.Buttons.Item("stop").Enabled = False
If InStr(1, URL, "about:blank") = 0 Then stbMain.Panels.Item(2).Text = webMain.LocationName
CheckCommands
End Sub

Private Sub webMain_DownloadBegin()
barMain.Visible = True
End Sub

Private Sub webMain_DownloadComplete()
barMain.Visible = False
End Sub

Private Sub webMain_NewWindow2(ppDisp As Object, Cancel As Boolean)
If tlbMain.Buttons.Item("popup").value = tbrPressed Then
    If CheckPopup = False Then
        Cancel = True
        stbMain.Panels.Item(3).Text = "A pop-up window has been blocked."
    End If
End If
End Sub

Private Sub webMain_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
barMain.Max = ProgressMax
barMain.value = Progress
End Sub
