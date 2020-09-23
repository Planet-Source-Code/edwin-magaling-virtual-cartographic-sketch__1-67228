VERSION 5.00
Begin VB.Form frmadduser 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adding Username and Password"
   ClientHeight    =   2925
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4395
   Icon            =   "frmadduser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1728.187
   ScaleMode       =   0  'User
   ScaleWidth      =   4126.667
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Unmask Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   90
      TabIndex        =   11
      Top             =   2475
      Width           =   1905
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "User-level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   405
      TabIndex        =   8
      Top             =   1530
      Width           =   3840
      Begin VB.OptionButton optother 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "User"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2385
         TabIndex        =   10
         Top             =   270
         Width           =   1185
      End
      Begin VB.OptionButton optadmin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Administrator"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   315
         TabIndex        =   9
         Top             =   270
         Width           =   1590
      End
   End
   Begin VB.TextBox txtVerifyPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   990
      Width           =   2445
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   0
      Top             =   180
      Width           =   2445
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2025
      TabIndex        =   3
      Top             =   2430
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3150
      TabIndex        =   4
      Top             =   2430
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   585
      Width           =   2445
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Verify Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   1080
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   675
      TabIndex        =   6
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   630
      TabIndex        =   5
      Top             =   270
      Width           =   975
   End
End
Attribute VB_Name = "frmadduser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        txtPassword.PasswordChar = ""
        txtVerifyPassword.PasswordChar = ""
    Else
        txtPassword.PasswordChar = "*"
        txtVerifyPassword.PasswordChar = "*"
    End If
End Sub

Private Sub cmdCancel_Click()
With frmusers
    .List1.Text = ""
    .cmdDelete.Enabled = False
    .cmdchange.Enabled = False
End With
    Unload Me
End Sub

Private Sub cmdSave_Click()
If txtUserName = Empty Then
    MsgBox "Input Username!", vbOKOnly + vbExclamation, "Information"
    txtUserName.SetFocus
    Exit Sub
End If
If txtPassword = Empty Then
    MsgBox "Input Password!", vbOKOnly + vbExclamation, "Information"
    txtPassword.SetFocus
    Exit Sub
End If
If txtVerifyPassword = Empty Then
    MsgBox "Verify your Password!", vbOKOnly + vbExclamation, "Information"
    txtVerifyPassword.SetFocus
    Exit Sub
End If
If optadmin.Value = False And optother.Value = False Then
    MsgBox "Select User-level!", vbOKOnly + vbExclamation, "Information"
    Exit Sub
End If
If MsgBox("Save New Username and Password?", vbYesNo + vbQuestion, "Saving Username and Password") = vbNo Then
    Exit Sub
End If

    
    
    'users
    Set userRS = New ADODB.Recordset
    userStr = "select username from users where username='" & Trim(txtUserName.Text) & "'"
    userRS.Open userStr, userConn, adOpenKeyset, adLockOptimistic
    If userRS.BOF And userRS.EOF Then
        If txtPassword = txtVerifyPassword Then
            Set userRS = New ADODB.Recordset
            userRS.Open "users", userConn, adOpenKeyset, adLockOptimistic
            With userRS
                .AddNew
                !Username = txtUserName.Text
                !Password = txtPassword.Text
                If optadmin = True Then
                    !userlevel = "Administrator"
                Else
                    !userlevel = "User"
                End If
                !background = "\background\default.jpg"
                .Update
                .Close
            End With
    
            txtUserName.Text = ""
            txtPassword.Text = ""
            txtVerifyPassword.Text = ""
            txtUserName.SetFocus
        Else
            MsgBox "Verify Password"
            txtVerifyPassword.Text = ""
            txtVerifyPassword.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Username Already Exists!", vbOKOnly + vbInformation, "Information"
        txtUserName.Text = ""
        txtPassword.Text = ""
        txtVerifyPassword.Text = ""
        txtUserName.SetFocus
        Exit Sub
    End If
       
    frmusers.List1.Clear

    'user
    Set userRS = New ADODB.Recordset
    userRS.Open "users", userConn, adOpenKeyset, adLockReadOnly

    While userRS.EOF <> True
        frmusers.List1.AddItem userRS!Username
        userRS.MoveNext
    Wend


frmusers.cmdDelete.Enabled = False
frmusers.cmdchange.Enabled = False

MsgBox "Username and Password Saved!", vbOKOnly + vbInformation, "Information"
Unload Me

End Sub

Private Sub Form_Load()
  userdbConnect
   txtPassword.PasswordChar = "*"
   txtVerifyPassword.PasswordChar = "*"
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtVerifyPassword.SetFocus
End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPassword.SetFocus
End If
End Sub

Private Sub txtVerifyPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    optadmin.SetFocus
End If
End Sub
