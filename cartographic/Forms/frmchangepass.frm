VERSION 5.00
Begin VB.Form frmchange 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Changing Password"
   ClientHeight    =   2115
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4380
   Icon            =   "frmchangepass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1249.612
   ScaleMode       =   0  'User
   ScaleWidth      =   4112.582
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
      TabIndex        =   8
      Top             =   1620
      Width           =   1905
   End
   Begin VB.TextBox txtoldpassword 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   540
      Width           =   2445
   End
   Begin VB.TextBox txtNewPassword 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   990
      Width           =   2445
   End
   Begin VB.CommandButton cmdchange 
      Caption         =   "C&hange"
      Enabled         =   0   'False
      Height          =   390
      Left            =   2025
      TabIndex        =   2
      Top             =   1575
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   3150
      TabIndex        =   4
      Top             =   1575
      Width           =   1095
   End
   Begin VB.Label lblusername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1890
      TabIndex        =   7
      Top             =   180
      Width           =   75
   End
   Begin VB.Label Username 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username : "
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
      Left            =   540
      TabIndex        =   6
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password :"
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
      Left            =   270
      TabIndex        =   5
      Top             =   630
      Width           =   1290
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password :"
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
      Index           =   2
      Left            =   180
      TabIndex        =   3
      Top             =   1080
      Width           =   1380
   End
End
Attribute VB_Name = "frmchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        txtoldpassword.PasswordChar = ""
        txtNewPassword.PasswordChar = ""
    Else
        txtoldpassword.PasswordChar = "*"
        txtNewPassword.PasswordChar = "*"
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

Private Sub cmdchange_Click()
If txtNewPassword = Empty Then
    MsgBox "Input New Password!", vbOKOnly + vbInformation, "Information"
    Exit Sub
End If

If MsgBox("Change Password?", vbYesNo + vbQuestion, "Changing Password") = vbNo Then
    Exit Sub
End If
    
    'users
    Set userRS = New ADODB.Recordset
    userStr = "select * from users where username='" & Trim(lblusername.Caption) & "'"
    userRS.Open userStr, userConn, adOpenKeyset, adLockOptimistic
    With userRS
        !Password = txtNewPassword
        .Update
        .Close
    End With

With frmusers
    .List1.Text = ""
    .cmdchange.Enabled = False
    .cmdDelete.Enabled = False
End With

MsgBox "Password Successfully Changed!", vbOKOnly + vbInformation, "Information"
Unload Me

End Sub

Private Sub Form_Load()
    userdbConnect
    txtoldpassword.PasswordChar = "*"
    txtNewPassword.PasswordChar = "*"
End Sub

Private Sub txtNewPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdchange.SetFocus
End If
End Sub

Private Sub txtoldpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

'users
Set userRS = New ADODB.Recordset
userStr = "Select * From users Where username = '" & Trim(lblusername.Caption) & "'"
userRS.Open userStr, userConn, adOpenKeyset, adLockReadOnly
    
    If Not userRS.EOF And Not userRS.BOF Then
        If txtoldpassword.Text <> userRS!Password Then
            MsgBox "Invalid Password", vbCritical, Caption
            txtoldpassword.Text = ""
            txtoldpassword.SetFocus
            Exit Sub
        Else
            txtNewPassword.Enabled = True
            txtNewPassword.SetFocus
            txtoldpassword.Enabled = False
            cmdchange.Enabled = True
        End If
    End If

End If
End Sub
