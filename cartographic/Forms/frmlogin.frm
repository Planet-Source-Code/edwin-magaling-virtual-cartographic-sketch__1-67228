VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Cartographic Sketch ver. 1.0"
   ClientHeight    =   1800
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4560
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1063.5
   ScaleMode       =   0  'User
   ScaleWidth      =   4281.593
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
      Left            =   135
      TabIndex        =   6
      Top             =   1305
      Width           =   1905
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
      Left            =   2070
      TabIndex        =   0
      Top             =   135
      Width           =   2310
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
      Left            =   2070
      TabIndex        =   1
      Top             =   630
      Width           =   2310
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&k"
      Height          =   390
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Log-In"
      Top             =   1260
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancel"
      Top             =   1260
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   180
      Picture         =   "frmlogin.frx":0442
      Stretch         =   -1  'True
      Top             =   270
      Width           =   510
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "User Name :"
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
      Index           =   0
      Left            =   900
      TabIndex        =   5
      Top             =   195
      Width           =   1065
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Index           =   1
      Left            =   990
      TabIndex        =   4
      Top             =   720
      Width           =   945
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pwctr As Integer

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        txtPassword.PasswordChar = ""
    Else
        txtPassword.PasswordChar = "*"
    End If
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()

Set userRS = New ADODB.Recordset
If txtUserName.Text <> "" Then
    
    userStr = "Select * From users Where username = '" & Trim(txtUserName.Text) & "'"
    userRS.Open userStr, userConn, adOpenKeyset, adLockReadOnly
    If userRS!userlevel = "Administrator" Then
        frmmain.mnutoolsuser.Enabled = True
    Else
        frmmain.mnutoolsuser.Enabled = False
    End If
    
    If Not userRS.EOF And Not userRS.BOF Then
        If txtPassword.Text <> userRS!Password Then
            pwctr = pwctr + 1
            If pwctr = 1 Then
                MsgBox "Invalid password! You have 2 tries remaining!", vbOKOnly + vbInformation, "Information"
                txtPassword.Text = ""
                txtPassword.SetFocus
            
            ElseIf pwctr = 2 Then
                MsgBox "Invalid password! You only have 1 try remaining!", vbOKOnly + vbInformation, "Information"
                txtPassword.Text = ""
                txtPassword.SetFocus
            Else
               End
            End If
        Else
            Set userRS = New ADODB.Recordset
            userRS.Open "users", userConn, adOpenKeyset, adLockOptimistic
            
            With userRS
            While .EOF <> True
                !Status = 0
                .Update
                .MoveNext
            Wend
                .Close
            End With
            
            Set userRS = New ADODB.Recordset
            userStr = "Select status from users where username='" & (txtUserName.Text) & "'"
            userRS.Open userStr, userConn, adOpenKeyset, adLockOptimistic
            With userRS
                !Status = 1
                .Update
                .Close
            End With
            
            Unload Me
            Set userRS = New ADODB.Recordset
            userStr = "Select background from users where status=" & 1
            userRS.Open userStr, userConn, adOpenKeyset, adLockReadOnly
            frmmain.Picture = LoadPicture(App.Path & userRS!background)
            If userRS!background = "\background\default.jpg" Then
                frmmain.mnuextrasbackdefault.Checked = True
            ElseIf userRS!background = "\background\bg1.jpg" Then
                frmmain.mnuextrasbackf1.Checked = True
            ElseIf userRS!background = "\background\bg2.jpg" Then
                frmmain.mnuextrasbackf2.Checked = True
            End If
            frmmain.Show
        End If
    Else
        MsgBox "Invalid Username!", vbOKOnly + vbExclamation, "Warning.."
        txtUserName.Text = ""
        txtPassword.Text = ""
        txtUserName.SetFocus
    End If
Else
    MsgBox "Invalid Username and Password!", vbOKOnly + vbExclamation, "Warning.."
    txtUserName.SetFocus
End If

End Sub

Private Sub Form_Load()
    userdbConnect
    txtPassword.PasswordChar = "*"
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdok.Enabled = True Then
        cmdok.Value = True
    End If
End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPassword.SetFocus
End If
End Sub
