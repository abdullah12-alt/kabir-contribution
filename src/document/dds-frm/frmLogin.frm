VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login - HEARTS Direct Deposit"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1410
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1020
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   2265
      TabIndex        =   9
      Tag             =   "Cancel"
      Top             =   1920
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   660
      TabIndex        =   8
      Tag             =   "OK"
      Top             =   1920
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   615
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   225
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Confirm Password:"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   6
      Tag             =   "&Password:"
      Top             =   1440
      Width           =   1365
   End
   Begin VB.Label lblLabels 
      Caption         =   "&New Password:"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Tag             =   "&Password:"
      Top             =   1050
      Width           =   1365
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Tag             =   "&Password:"
      Top             =   660
      Width           =   1365
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Tag             =   "&User Name:"
      Top             =   255
      Width           =   1365
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************************************************
' * Description:
' * This form performs four functions
' *  1) Initial Logon to application
' *  2) Change the current user
' *  3) Change user password
' *  4) restore screen from sleep mode
' *
' *
' * Revisions:
' *  3/23/99    For changing password, I eliminated supplying the password
' *
' *
' ********************************************************************************


' Mod CONSTANTS
Private Const MODULE As String = "Login - "

' Mod ENUMS
Public Enum LoginMode
    LOGIN
    LOGINNEWUSER
    ChangePassword
    STEPAWAYMODE
End Enum
' Mod VARIABLES
Public mcmd As New ADODB.Command

Private Const EXTENDSTR As String = "P#Ssa(fC"
' Mod VARIABLES

Public OK As Boolean
Public iMode As LoginMode

Private Sub Form_Activate()
    Select Case iMode
    Case LOGIN
        If Len(txtUserName) > 0 Then
            txtPassword(0).SetFocus
        Else
            txtUserName.SetFocus
            SetSelected
        End If
    Case LOGINNEWUSER
        txtUserName.SetFocus
        SetSelected
        fMainForm.SetMainToolbar True
            
    Case ChangePassword
        txtPassword(0).SetFocus
        fMainForm.SetMainToolbar True
    Case STEPAWAYMODE
        If Len(txtUserName) > 0 Then
            txtPassword(0).SetFocus
        Else
            txtUserName.SetFocus
        End If
        fMainForm.SetMainToolbar True
    
    End Select
End Sub

Private Sub Form_Deactivate()

    Select Case iMode
    Case LOGINNEWUSER, ChangePassword, STEPAWAYMODE
        fMainForm.SetMainToolbar False
    End Select


End Sub

Private Sub Form_Load()
    
    Select Case iMode
    Case LOGIN
        txtUserName.Text = GetSetting(App.Title, "Login", "UserName", "")
        Me.Caption = "Login - HEARTS Direct Deposit"
        txtPassword(1).Visible = False
        txtPassword(2).Visible = False
        cmdOK.Top = txtPassword(1).Top + 100
        cmdCancel.Top = txtPassword(1).Top + 100
        Me.Height = cmdOK.Top + cmdOK.Height + 700
        lblLabels(2).Visible = False
        lblLabels(3).Visible = False
    Case LOGINNEWUSER
        txtUserName.Text = gobjLoginInfo.UserId
        Me.Caption = "Login as New User"
        txtPassword(1).Visible = False
        txtPassword(2).Visible = False
        cmdOK.Top = txtPassword(1).Top + 100
        cmdCancel.Top = txtPassword(1).Top + 100
        Me.Height = cmdOK.Top + cmdOK.Height + 700
        lblLabels(2).Visible = False
        lblLabels(3).Visible = False
    Case ChangePassword
        Me.Caption = "Change Password"
        txtUserName.Text = gobjLoginInfo.UserId
        txtUserName.Locked = True
    Case STEPAWAYMODE
        Me.Caption = "Resume Work"
        txtUserName.Text = gobjLoginInfo.UserId
        txtPassword(1).Visible = False
        txtPassword(2).Visible = False
        cmdOK.Top = txtPassword(1).Top + 100
        cmdCancel.Top = txtPassword(1).Top + 100
        Me.Height = cmdOK.Top + cmdOK.Height + 700
        lblLabels(2).Visible = False
        lblLabels(3).Visible = False
            
    End Select
    
    '3/15/2009 - AS Added code to
    Set mcmd = Nothing
    Set mcmd.ActiveConnection = gcnDDS
    If gStoredProcs("up_Login_Stats").GetStoredProcCommand(mcmd) = False Then
        Err.Raise -123456
    End If

End Sub



Private Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    Dim sSql As String
    Dim oSHA256 As cSHA256
    Set oSHA256 = New cSHA256
    Dim rsCheck As New ADODB.Recordset

On Error GoTo cmdOKErr

Hourglass True
Static iLoginAttempts As Integer
Dim sTempConnectString As String
With gobjLoginInfo
    .UserId = txtUserName
    .UserPassword = txtPassword(0)
'    If .DBDriver = "" Or .DBName = "" Or .DDSServer = "" Then
'        MsgBox "One or more settings in the initialization file is not set.  Exiting application.", vbCritical
'        ExitApp
'    End If

    Select Case iMode
    Case LOGIN
        iLoginAttempts = iLoginAttempts + 1
        sSql = "SELECT * FROM DD_USER WHERE USER_ID = '" & .UserId & "'"
        rsCheck.Open sSql, gcnDDS, adOpenStatic
        If rsCheck.EOF Then
            'User Does Not Exist
            Err.Raise -2147467259
        Else
            If rsCheck!RECORD_STATUS = "I" Then
                Err.Raise -1234568
            End If
            '2/5/2009 - AS - Changed the code to no longer check user credential with sybase
            If rsCheck!PSWTEXT = oSHA256.SHA256(.UserId & .UserPassword & EXTENDSTR) Then
                'User logged in successfully
                mcmd.Parameters("user_id") = .UserId
                mcmd.Parameters("misc_text") = ""
                mcmd.Parameters("version") = App.Major & "." & App.Minor
                mcmd.Parameters("update_mode") = "L"
                mcmd.Execute
                If mcmd.Parameters("RETURN_VALUE").value <> 0 Then
                    Err.Raise -12345679
                End If
            Else
                'Password does not match what is on file
                mcmd.Parameters("user_id") = .UserId
                mcmd.Parameters("misc_text") = ""
                mcmd.Parameters("version") = App.Major & "." & App.Minor
                mcmd.Parameters("update_mode") = "F"
                mcmd.Execute
                Err.Raise -2147467259
            End If
        End If
        On Error GoTo cmdOKErr
        
        iLoginAttempts = 0
        OK = True
        Me.Hide
    Case LOGINNEWUSER
         iLoginAttempts = iLoginAttempts + 1
        
        sSql = "SELECT * FROM DD_USER WHERE USER_ID = '" & .UserId & "'"
        rsCheck.Open sSql, gcnDDS, adOpenStatic
        If rsCheck.EOF Then
            'User Does Not Exist
            Err.Raise -2147467259
        Else
            If rsCheck!RECORD_STATUS = "I" Then
                Err.Raise -1234568
            End If
            If IsNull(rsCheck!PSWTEXT) Then
            '2/5/2009 - AS - Changed the code to no longer check user credential with sybase
                MsgBox "Login credentials expired.  Please contact administrator to reset password. Exiting Application.", vbCritical
                ExitApp
            Else
                If rsCheck!PSWTEXT = oSHA256.SHA256(.UserId & .UserPassword & EXTENDSTR) Then
                    'User logged in successfully
                    mcmd.Parameters("user_id") = .UserId
                      mcmd.Parameters("misc_text") = ""
                    mcmd.Parameters("version") = App.Major & "." & App.Minor
                    mcmd.Parameters("update_mode") = "L"
                    mcmd.Execute
                    If mcmd.Parameters("RETURN_VALUE").value <> 0 Then
                        Err.Raise -12345679
                    End If
                Else
                    'Password does not match what is on file
                    mcmd.Parameters("user_id") = .UserId
                    mcmd.Parameters("misc_text") = ""
                    mcmd.Parameters("version") = App.Major & "." & App.Minor
                    mcmd.Parameters("update_mode") = "F"
                    mcmd.Execute
                    Err.Raise -2147467259
                End If
            End If
        End If
        
        
        iLoginAttempts = 0
        Hourglass False
        OK = True
        Me.Hide
    Case ChangePassword
        
        Dim cmd As New ADODB.Command
        
        If txtPassword(0) <> gobjLoginInfo.UserPassword Then
            MsgBox "The old password supplied is incorrect. Try again.", vbInformation
            Hourglass False
            txtPassword(0).SetFocus
            Exit Sub
        End If
        
        If Len(txtPassword(1)) < 6 Then
            MsgBox "New password must be greater than 6 alphanumeric characters.", vbInformation
            Hourglass False
            txtPassword(1).SetFocus
            Exit Sub
        Else
            If txtPassword(1) <> txtPassword(2) Then
                MsgBox "The confirmed password does not match the new password.", vbInformation
                Hourglass False
                txtPassword(2).SetFocus
                Exit Sub
            Else
                If txtPassword(0) = txtPassword(1) Then
                    MsgBox "The new password cannot be the same as your old password.", vbInformation
                    Hourglass False
                    txtPassword(1).SetFocus
                    Exit Sub
                End If
            End If
        End If
        mcmd.Parameters("user_id") = txtUserName
        mcmd.Parameters("update_mode") = "P"
        mcmd.Parameters("misc_text") = oSHA256.SHA256(gobjLoginInfo.UserId & txtPassword(1).Text & EXTENDSTR)
        mcmd.Parameters("version") = App.Major & "." & App.Minor
        mcmd.Execute
        If mcmd.Parameters("RETURN_VALUE").value <> 0 Then
            Err.Raise -1237533
        End If
'        With gobjLoginInfo
'            .UserPassword = txtPassword(1)
'            .ConnectString = "Provider=MSDASQL.1;DRIVER={" & .DBDriver & "};UID=" & .UserId & ";PWD=" & .UserPassword & ";database=" & .DBName & ";SRVR=" & .DDSServer
'        End With
        iLoginAttepts = 0
        Hourglass False
        MsgBox "Password was successfully changed.", vbInformation, "Change Password"
        OK = True
        Me.Hide
    Case STEPAWAYMODE
        If txtPassword(0) = gobjLoginInfo.UserPassword Then
            OK = True
            Me.Hide
        Else
            OK = False
            Me.Hide
        End If
    End Select
End With
Set rsCheck = Nothing
Exit Sub

cmdOKErr:

    Select Case Err
    Case -1234567
        MsgBox "Password must be reset.  Please contact your Personal Funds Security Officer"
        Set rsCheck = Nothing
        ExitApp
    Case -1234568
        MsgBox "Your user id has been deactivated.  Please contact your Personal Funds Security Officer"
        Set rsCheck = Nothing
        ExitApp
    Case -2147467259
        Set rsCheck = Nothing
        Hourglass False
        MsgBox "The user id or password provided is not valid. Please try again.", vbInformation
        txtUserName.SetFocus
        txtUserName.SelStart = 0
        txtUserName.SelLength = Len(txtUserName.Text)
    Case -1237533
        Set rsCheck = Nothing
        Hourglass False
        MsgBox "The password was not changed correctly. Try again.", vbInformation
        txtUserName.SetFocus
        txtUserName.SelStart = 0
        txtUserName.SelLength = Len(txtUserName.Text)
    Case Else
        MsgBox Err.Number + ", " + Err.Description + ", " + Err.Source + ", The database is unavailable at the present time, Please try again at a later time.", , "Login"
        Set rsCheck = Nothing
        Resume
        ExitApp
    End Select

    If iLoginAttempts = 3 Then
        If iMode = LOGIN Then
            ExitApp
        Else
            Unload Me
        End If
    End If
    
End Sub


Private Sub txtPassword_Change(Index As Integer)
    
    ValidateCtrls

End Sub

Private Sub txtPassword_GotFocus(Index As Integer)

    SetSelected
    
End Sub


Private Sub txtUserName_Change()
    
    ValidateCtrls

End Sub

Private Sub txtUserName_GotFocus()
    SetSelected
    
End Sub

Private Sub ValidateCtrls()

    Select Case iMode
    Case LOGIN, LOGINNEWUSER, STEPAWAYMODE
        If Len(txtUserName) > 0 And Len(txtPassword(0)) > 5 Then
            cmdOK.Enabled = True
        Else
            cmdOK.Enabled = False
        End If
    Case ChangePassword
        If Len(txtUserName) > 0 And Len(txtPassword(0)) > 5 And Len(txtPassword(1)) > 5 And Len(txtPassword(2)) > 5 Then
            cmdOK.Enabled = True
        Else
            cmdOK.Enabled = False
        End If
    End Select
    
End Sub
