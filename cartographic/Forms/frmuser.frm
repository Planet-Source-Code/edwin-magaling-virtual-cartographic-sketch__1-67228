VERSION 5.00
Begin VB.Form frmusers 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Management"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frmuser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5100
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "User-Level"
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
      Height          =   1095
      Left            =   2610
      TabIndex        =   6
      Top             =   1125
      Width           =   2085
      Begin VB.Label Label3 
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
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   360
         TabIndex        =   7
         Top             =   450
         Width           =   75
      End
   End
   Begin VB.CommandButton cmdchange 
      Caption         =   "&Change  Password"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1215
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cl&ose"
      Height          =   390
      Left            =   3870
      TabIndex        =   4
      Top             =   3240
      Width           =   1080
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   390
      Left            =   2790
      TabIndex        =   3
      Top             =   3240
      Width           =   1080
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&New"
      Height          =   390
      Left            =   135
      TabIndex        =   1
      Top             =   3240
      Width           =   1080
   End
   Begin VB.ListBox List1 
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
      Height          =   2430
      Left            =   180
      TabIndex        =   0
      Top             =   585
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User List"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   180
      TabIndex        =   2
      Top             =   90
      Width           =   1545
   End
End
Attribute VB_Name = "frmusers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
    frmadduser.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdchange_Click()

'users
Set userRS = New ADODB.Recordset
userStr = "SELECT * FROM users where username='" & List1.Text & "'"
userRS.Open userStr, userConn, adOpenKeyset, adLockOptimistic

frmchange.lblusername.Caption = userRS!Username

frmchange.Show vbModal

End Sub

Private Sub cmdDelete_Click()

If List1.Text = "Administrator" Then
    MsgBox "Default User, access denied!", vbOKOnly + vbExclamation, "Warning..."
    List1.Text = ""
    cmdDelete.Enabled = False
    cmdchange.Enabled = False
    Exit Sub
End If
If MsgBox("Delete Username and Password?", vbYesNo + vbQuestion, "Deleting User") = vbNo Then
    List1.Text = ""
    cmdDelete.Enabled = False
    cmdchange.Enabled = False
    Exit Sub
End If
           
    'users
    Set userCmd = New ADODB.Command
    userStr = "DELETE * FROM users where username ='" & List1.Text & "'"
    With userCmd
        .ActiveConnection = userConn
        .CommandType = adCmdText
        .CommandText = userStr
        .Execute
    End With
List1.Clear
Call listrefresh

cmdDelete.Enabled = False
cmdchange.Enabled = False

MsgBox "Username and Password Deleted!", vbOKOnly + vbInformation, "Information"

End Sub


Private Sub Form_Load()
    userdbConnect
    Call listrefresh
End Sub

Private Sub listrefresh()
'users
Set userRS = New ADODB.Recordset
userRS.Open "users", userConn, adOpenKeyset, adLockReadOnly

While userRS.EOF <> True
   List1.AddItem userRS!Username
    userRS.MoveNext
Wend

End Sub

Private Sub List1_Click()
If List1.Text <> Empty Then
    cmdDelete.Enabled = True
    cmdchange.Enabled = True
End If

Set userRS = New ADODB.Recordset
userStr = "select userlevel from users where username='" & List1.Text & "'"
userRS.Open userStr, userConn, adOpenKeyset, adLockReadOnly
    
    Label3.Caption = userRS!userlevel
    
End Sub
