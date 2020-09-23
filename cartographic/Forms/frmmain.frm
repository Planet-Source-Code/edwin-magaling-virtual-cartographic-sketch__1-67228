VERSION 5.00
Begin VB.Form frmmain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Virtual Cartographic Sketch Version 1.0"
   ClientHeight    =   8280
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11055
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufilesketch 
         Caption         =   "&Cartograohic Sketch"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnufileprofile 
         Caption         =   "&Profile"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnufilesep 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnutoolsuser 
         Caption         =   "User Management"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuextras 
      Caption         =   "E&xtras"
      Begin VB.Menu mnuextraskeyboard 
         Caption         =   "On Screen &Keyboard"
      End
      Begin VB.Menu mnuextrassolitare 
         Caption         =   "Solitare"
      End
      Begin VB.Menu mnuextrassep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuextrasback 
         Caption         =   "Backgrounds"
         Begin VB.Menu mnuextrasbackf1 
            Caption         =   "Cartographic 1"
         End
         Begin VB.Menu mnuextrasbackf2 
            Caption         =   "Cartographic 2"
         End
         Begin VB.Menu mnuextrasbacksep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuextrasbackdefault 
            Caption         =   "Default..."
         End
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelptips 
         Caption         =   "&Tips..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuhelpsystem 
         Caption         =   "About the &System"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuhelpsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuhelpabout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xbackground As String

Private Sub background()
    Me.Picture = LoadPicture(App.Path & xbackground)
    Set userRS = New ADODB.Recordset
    userStr = "Select background from users where status=" & 1
    userRS.Open userStr, userConn, adOpenKeyset, adLockOptimistic
    With userRS
        !background = xbackground
        .Update
        .Close
    End With
End Sub

Private Sub mnuextrasback1_Click()

End Sub

Private Sub mnuextrasback2_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Exit Virtual Cartographic Sketch?", vbQuestion + vbYesNo, "Virtual Cartographic Sketch") = vbYes Then
    End
Else
    Cancel = True
End If
End Sub

Private Sub mnuextrasbackdefault_Click()
mnuextrasbackf1.Checked = False
mnuextrasbackf2.Checked = False
mnuextrasbackdefault.Checked = True
xbackground = "\background\default.jpg"
background
End Sub

Private Sub mnuextrasbackf1_Click()
mnuextrasbackf1.Checked = True
mnuextrasbackf2.Checked = False
mnuextrasbackdefault.Checked = False
xbackground = "\background\bg1.jpg"
background
End Sub

Private Sub mnuextrasbackf2_Click()
mnuextrasbackf1.Checked = False
mnuextrasbackf2.Checked = True
mnuextrasbackdefault.Checked = False
xbackground = "\background\bg2.jpg"
background
End Sub

Private Sub mnuextraskeyboard_Click()
    Shell ("osk.exe"), vbNormalFocus
End Sub

Private Sub mnuextrassolitare_Click()
    Shell ("sol.exe"), vbNormalFocus
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub

Private Sub mnufileprofile_Click()
    frmprofile.Show vbModal
End Sub

Private Sub mnufilesketch_Click()
    Set userRS = New ADODB.Recordset
    userStr = "Select userlevel from users where status=" & 1
    userRS.Open userStr, userConn, adOpenKeyset, adLockReadOnly
    If userRS!userlevel <> "Administrator" Then
        frmcartographic.mnutoolsparts.Enabled = False
    End If
    frmcartographic.Show vbModal
End Sub

Private Sub mnuhelpabout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuhelpsystem_Click()
    frmhelpabout1.Show vbModal
End Sub

Private Sub mnuhelptips_Click()
    frmhelptips1.Show vbModal
End Sub

Private Sub mnutoolsuser_Click()
    frmusers.Show vbModal
End Sub
