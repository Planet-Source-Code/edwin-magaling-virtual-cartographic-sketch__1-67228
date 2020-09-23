VERSION 5.00
Begin VB.Form frmpartmanager 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Head Parts Manager"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   Icon            =   "frmpartmanager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   9600
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sizes"
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
      Height          =   1455
      Left            =   6480
      TabIndex        =   26
      Top             =   1980
      Width           =   2895
      Begin VB.CommandButton cmdsizeleft 
         Height          =   330
         Left            =   1530
         Picture         =   "frmpartmanager.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   585
         Width           =   375
      End
      Begin VB.CommandButton cmdsizedown 
         Height          =   330
         Left            =   1890
         Picture         =   "frmpartmanager.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   855
         Width           =   375
      End
      Begin VB.CommandButton cmdsizeright 
         Height          =   330
         Left            =   2250
         Picture         =   "frmpartmanager.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   585
         Width           =   375
      End
      Begin VB.CommandButton cmdsizeup 
         Height          =   330
         Left            =   1890
         Picture         =   "frmpartmanager.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   315
         Width           =   375
      End
      Begin VB.TextBox txtw 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   855
         TabIndex        =   28
         Top             =   405
         Width           =   555
      End
      Begin VB.TextBox txth 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   855
         TabIndex        =   27
         Top             =   855
         Width           =   555
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Width :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   34
         Top             =   495
         Width           =   510
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Height :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   33
         Top             =   900
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Alignment"
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
      Height          =   1455
      Left            =   6480
      TabIndex        =   17
      Top             =   360
      Width           =   2895
      Begin VB.TextBox txty 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   855
         TabIndex        =   25
         Top             =   855
         Width           =   555
      End
      Begin VB.TextBox txtx 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   855
         TabIndex        =   24
         Top             =   405
         Width           =   555
      End
      Begin VB.CommandButton cmdalignup 
         Height          =   330
         Left            =   1890
         Picture         =   "frmpartmanager.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   315
         Width           =   375
      End
      Begin VB.CommandButton cmdalignright 
         Height          =   330
         Left            =   2250
         Picture         =   "frmpartmanager.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   585
         Width           =   375
      End
      Begin VB.CommandButton cmdaligndown 
         Height          =   330
         Left            =   1890
         Picture         =   "frmpartmanager.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   810
         Width           =   375
      End
      Begin VB.CommandButton cmdalignleft 
         Height          =   330
         Left            =   1530
         Picture         =   "frmpartmanager.frx":2210
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   585
         Width           =   375
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Y-Axis :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   23
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "X-Axis :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   22
         Top             =   495
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save To Database"
      Height          =   420
      Left            =   7065
      TabIndex        =   16
      Top             =   6750
      Width           =   1635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cl&ose"
      Height          =   420
      Left            =   7065
      TabIndex        =   13
      Top             =   7470
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Head Parts Category"
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
      Height          =   2535
      Left            =   6480
      TabIndex        =   2
      Top             =   3645
      Width           =   2895
      Begin VB.OptionButton optcap 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cap"
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1620
         TabIndex        =   12
         Top             =   1980
         Width           =   1050
      End
      Begin VB.OptionButton optbeard 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Beard"
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1620
         TabIndex        =   11
         Top             =   1575
         Width           =   1050
      End
      Begin VB.OptionButton optglass 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Glass"
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1620
         TabIndex        =   10
         Top             =   1170
         Width           =   1050
      End
      Begin VB.OptionButton optlips 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Lips"
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1620
         TabIndex        =   9
         Top             =   765
         Width           =   1050
      End
      Begin VB.OptionButton optnose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nose"
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1620
         TabIndex        =   8
         Top             =   360
         Width           =   1050
      End
      Begin VB.OptionButton opteyes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Eyes"
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   270
         TabIndex        =   7
         Top             =   1980
         Width           =   1050
      End
      Begin VB.OptionButton optbrow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Eyebrow"
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   270
         TabIndex        =   6
         Top             =   1575
         Width           =   1050
      End
      Begin VB.OptionButton optears 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ears"
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   270
         TabIndex        =   5
         Top             =   1170
         Width           =   1050
      End
      Begin VB.OptionButton opthair 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hair"
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   270
         TabIndex        =   4
         Top             =   765
         Width           =   1050
      End
      Begin VB.OptionButton optjaw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Jaw"
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   270
         TabIndex        =   3
         Top             =   360
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load New Sketch"
      Height          =   420
      Left            =   7065
      TabIndex        =   1
      Top             =   6300
      Width           =   1635
   End
   Begin VB.PictureBox picpartout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7530
      Left            =   225
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   0
      Top             =   360
      Width           =   6045
   End
   Begin VB.PictureBox picpart 
      Height          =   465
      Left            =   5805
      ScaleHeight     =   405
      ScaleWidth      =   135
      TabIndex        =   15
      Top             =   7065
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label tmp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1530
      TabIndex        =   14
      Top             =   90
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "frmpartmanager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y, w, h As Integer 'alignment & sizes
Dim message, title, default, myvalue, dest As String

Private Sub optdisable()
    optjaw.Enabled = False
    opthair.Enabled = False
    optears.Enabled = False
    optbrow.Enabled = False
    opteyes.Enabled = False
    optnose.Enabled = False
    optlips.Enabled = False
    optglass.Enabled = False
    optbeard.Enabled = False
    optcap.Enabled = False
End Sub

Private Sub optenable()
    optjaw.Enabled = True
    opthair.Enabled = True
    optears.Enabled = True
    optbrow.Enabled = True
    opteyes.Enabled = True
    optnose.Enabled = True
    optlips.Enabled = True
    optglass.Enabled = True
    optbeard.Enabled = True
    optcap.Enabled = True
End Sub

Private Sub aligndisable()
    txtx.Enabled = False
    txty.Enabled = False
    cmdalignup.Enabled = False
    cmdaligndown.Enabled = False
    cmdalignleft.Enabled = False
    cmdalignright.Enabled = False
End Sub

Private Sub alignenable()
    txtx.Enabled = True
    txty.Enabled = True
    cmdalignup.Enabled = True
    cmdaligndown.Enabled = True
    cmdalignleft.Enabled = True
    cmdalignright.Enabled = True
End Sub

Private Sub sizedisable()
    txtw.Enabled = False
    txtw.Enabled = False
    cmdsizeup.Enabled = False
    cmdsizedown.Enabled = False
    cmdsizeleft.Enabled = False
    cmdsizeright.Enabled = False
End Sub

Private Sub sizeenable()
    txtw.Enabled = True
    txth.Enabled = True
    cmdsizeup.Enabled = True
    cmdsizedown.Enabled = True
    cmdsizeleft.Enabled = True
    cmdsizeright.Enabled = True
End Sub

Private Sub cmdaligndown_Click()
    txty.Text = txty.Text + 1
    picpartout.Cls
    picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
End Sub

Private Sub cmdalignleft_Click()
    txtx.Text = txtx.Text - 1
    picpartout.Cls
    picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
End Sub

Private Sub cmdalignright_Click()
    txtx.Text = txtx.Text + 1
    picpartout.Cls
    picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
End Sub

Private Sub cmdalignup_Click()
      txty.Text = txty.Text - 1
      picpartout.Cls
      picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
End Sub

Private Sub Cmdload_Click()
File = Open_File(Me.hWnd) 'show the open file dlg
If Trim(File) = "" Then Exit Sub ' make sure the file is correct
tmp.Caption = File

x = 10
y = 20
w = 400
h = 450

picpartout.Cls
picpart.Picture = LoadPicture(tmp.Caption)
picpartout.PaintPicture picpart.Picture, x, y, w, h, , , , , vbSrcAnd

txtx = x
txty = y
txtw = w
txth = h

alignenable
sizeenable
optenable

End Sub

Private Sub cmdSave_Click()
start:
    message = "Input Filename"
    title = "Saving Sketch..."
    If optjaw.Value = True Then
        default = "jaw"
    ElseIf opthair.Value = True Then
        default = "Hair"
    ElseIf optears.Value = True Then
        default = "Ears"
    ElseIf optbrow.Value = True Then
        default = "Eyebrow"
    ElseIf opteyes.Value = True Then
        default = "Eyes"
    ElseIf optnose.Value = True Then
        default = "Nose"
    ElseIf optlips.Value = True Then
        default = "Lips"
    ElseIf optglass.Value = True Then
        default = "Glass"
    ElseIf optbeard.Value = True Then
        default = "Beard"
    ElseIf optcap.Value = True Then
        default = "Cap"
    End If
        
    myvalue = InputBox(message, title, default)
    If myvalue = "" Then Exit Sub
    myvalue = myvalue + ".jpg"
    dest = "\pics\" & myvalue
    
    If default = "jaw" Then 'jaw
        Set jawRS = New ADODB.Recordset
        jawStr = "Select pic from jaw where pic='" & dest & "'"
        jawRS.Open jawStr, partsConn, adOpenKeyset, adLockReadOnly
        If Not jawRS.EOF And Not jawRS.BOF Then
            MsgBox "Filename Already Exist!", vbExclamation, "Cartographic"
            GoTo start
        End If
        Set jawRS = New ADODB.Recordset
        jawRS.Open "jaw", partsConn, adOpenKeyset, adLockOptimistic
        With jawRS
            .AddNew
            !pic = dest
            !x = txtx.Text
            !y = txty.Text
            !Width = txtw.Text
            !Height = txth.Text
            .Update
            .Close
        End With
    End If
    
    If default = "Hair" Then 'hair
        Set hairRS = New ADODB.Recordset
        hairStr = "Select pic from hair where pic='" & dest & "'"
        hairRS.Open hairStr, partsConn, adOpenKeyset, adLockReadOnly
        If Not hairRS.EOF And Not hairRS.BOF Then
            MsgBox "Filename Already Exist!", vbExclamation, "Cartographic"
            GoTo start
        End If
        Set hairRS = New ADODB.Recordset
        hairRS.Open "hair", partsConn, adOpenKeyset, adLockOptimistic
        With hairRS
            .AddNew
            !pic = dest
            !x = txtx.Text
            !y = txty.Text
            !Width = txtw.Text
            !Height = txth.Text
            .Update
            .Close
        End With
    End If
    
    If default = "Ears" Then 'Ears
        Set earsRS = New ADODB.Recordset
        earsStr = "Select pic from Ears where pic='" & dest & "'"
        earsRS.Open earsStr, partsConn, adOpenKeyset, adLockReadOnly
        If Not earsRS.EOF And Not earsRS.BOF Then
            MsgBox "Filename Already Exist!", vbExclamation, "Cartographic"
            GoTo start
        End If
        Set earsRS = New ADODB.Recordset
        earsRS.Open "Ears", partsConn, adOpenKeyset, adLockOptimistic
        With earsRS
            .AddNew
            !pic = dest
            !x = txtx.Text
            !y = txty.Text
            !Width = txtw.Text
            !Height = txth.Text
            .Update
            .Close
        End With
    End If
    
    If default = "Eyebrow" Then 'Eyebrow
        Set browRS = New ADODB.Recordset
        browStr = "Select pic from brow where pic='" & dest & "'"
        browRS.Open browStr, partsConn, adOpenKeyset, adLockReadOnly
        If Not browRS.EOF And Not browRS.BOF Then
            MsgBox "Filename Already Exist!", vbExclamation, "Cartographic"
            GoTo start
        End If
        Set browRS = New ADODB.Recordset
        browRS.Open "brow", partsConn, adOpenKeyset, adLockOptimistic
        With browRS
            .AddNew
            !pic = dest
            !x = txtx.Text
            !y = txty.Text
            !Width = txtw.Text
            !Height = txth.Text
            .Update
            .Close
        End With
    End If
    
    If default = "Eyes" Then 'Eyes
        Set eyesRS = New ADODB.Recordset
        eyesStr = "Select pic from Eyes where pic='" & dest & "'"
        eyesRS.Open eyesStr, partsConn, adOpenKeyset, adLockReadOnly
        If Not eyesRS.EOF And Not eyesRS.BOF Then
            MsgBox "Filename Already Exist!", vbExclamation, "Cartographic"
            GoTo start
        End If
        Set eyesRS = New ADODB.Recordset
        eyesRS.Open "Eyes", partsConn, adOpenKeyset, adLockOptimistic
        With eyesRS
            .AddNew
            !pic = dest
            !x = txtx.Text
            !y = txty.Text
            !Width = txtw.Text
            !Height = txth.Text
            .Update
            .Close
        End With
    End If
    
    If default = "Nose" Then 'Nose
        Set noseRS = New ADODB.Recordset
        noseStr = "Select pic from Nose where pic='" & dest & "'"
        noseRS.Open noseStr, partsConn, adOpenKeyset, adLockReadOnly
        If Not noseRS.EOF And Not noseRS.BOF Then
            MsgBox "Filename Already Exist!", vbExclamation, "Cartographic"
            GoTo start
        End If
        Set noseRS = New ADODB.Recordset
        noseRS.Open "Nose", partsConn, adOpenKeyset, adLockOptimistic
        With noseRS
            .AddNew
            !pic = dest
            !x = txtx.Text
            !y = txty.Text
            !Width = txtw.Text
            !Height = txth.Text
            .Update
            .Close
        End With
    End If
        
    If default = "Lips" Then 'Lips
        Set lipsRS = New ADODB.Recordset
        lipsStr = "Select pic from Lips where pic='" & dest & "'"
        lipsRS.Open lipsStr, partsConn, adOpenKeyset, adLockReadOnly
        If Not lipsRS.EOF And Not lipsRS.BOF Then
            MsgBox "Filename Already Exist!", vbExclamation, "Cartographic"
            GoTo start
        End If
        Set lipsRS = New ADODB.Recordset
        lipsRS.Open "Lips", partsConn, adOpenKeyset, adLockOptimistic
        With lipsRS
            .AddNew
            !pic = dest
            !x = txtx.Text
            !y = txty.Text
            !Width = txtw.Text
            !Height = txth.Text
            .Update
            .Close
        End With
    End If
   
    If default = "Glass" Then 'Glass
        Set glassRS = New ADODB.Recordset
        glassStr = "Select pic from Glass where pic='" & dest & "'"
        glassRS.Open glassStr, partsConn, adOpenKeyset, adLockReadOnly
        If Not glassRS.EOF And Not glassRS.BOF Then
            MsgBox "Filename Already Exist!", vbExclamation, "Cartographic"
            GoTo start
        End If
        Set glassRS = New ADODB.Recordset
        glassRS.Open "Glass", partsConn, adOpenKeyset, adLockOptimistic
        With glassRS
            .AddNew
            !pic = dest
            !x = txtx.Text
            !y = txty.Text
            !Width = txtw.Text
            !Height = txth.Text
            .Update
            .Close
        End With
    End If
    
    If default = "Beard" Then 'Beard
        Set beardRS = New ADODB.Recordset
        beardStr = "Select pic from Beard where pic='" & dest & "'"
        beardRS.Open beardStr, partsConn, adOpenKeyset, adLockReadOnly
        If Not beardRS.EOF And Not beardRS.BOF Then
            MsgBox "Filename Already Exist!", vbExclamation, "Cartographic"
            GoTo start
        End If
        Set beardRS = New ADODB.Recordset
        beardRS.Open "Beard", partsConn, adOpenKeyset, adLockOptimistic
        With beardRS
            .AddNew
            !pic = dest
            !x = txtx.Text
            !y = txty.Text
            !Width = txtw.Text
            !Height = txth.Text
            .Update
            .Close
        End With
    End If
    
    If default = "Cap" Then 'Cap
        Set capRS = New ADODB.Recordset
        capStr = "Select pic from Cap where pic='" & dest & "'"
        capRS.Open capStr, partsConn, adOpenKeyset, adLockReadOnly
        If Not capRS.EOF And Not capRS.BOF Then
            MsgBox "Filename Already Exist!", vbExclamation
            GoTo start
        End If
        Set capRS = New ADODB.Recordset
        capRS.Open "Cap", partsConn, adOpenKeyset, adLockOptimistic
        With capRS
            .AddNew
            !pic = dest
            !x = txtx.Text
            !y = txty.Text
            !Width = txtw.Text
            !Height = txth.Text
            .Update
            .Close
        End With
    End If
    
    FileCopy tmp.Caption, App.Path & dest
    
    MsgBox "Sketch Successfully Saved!", vbInformation, "Cartographic"
    
End Sub

Private Sub cmdsizedown_Click()
    txth.Text = txth.Text - 1
    picpartout.Cls
    picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
End Sub

Private Sub cmdsizeleft_Click()
    txtw.Text = txtw.Text - 1
    picpartout.Cls
    picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
End Sub

Private Sub cmdsizeright_Click()
    txtw.Text = txtw.Text + 1
    picpartout.Cls
    picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
End Sub

Private Sub cmdsizeup_Click()
    txth.Text = txth.Text + 1
    picpartout.Cls
    picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    InitDlgs 'initalize save and open dialogs
    aligndisable
    sizedisable
    optdisable
    cmdSave.Enabled = False
End Sub

Private Sub optbeard_Click()
    If optbeard.Value = True Then
        cmdSave.Enabled = True
    End If
End Sub

Private Sub optbrow_Click()
    If optbrow.Value = True Then
        cmdSave.Enabled = True
    End If
End Sub

Private Sub optcap_Click()
    If optcap.Value = True Then
        cmdSave.Enabled = True
    End If
End Sub

Private Sub optears_Click()
    If optears.Value = True Then
        cmdSave.Enabled = True
    End If
End Sub

Private Sub opteyes_Click()
    If opteyes.Value = True Then
        cmdSave.Enabled = True
    End If
End Sub

Private Sub optglass_Click()
    If optglass.Value = True Then
        cmdSave.Enabled = True
    End If
End Sub

Private Sub opthair_Click()
    If opthair.Value = True Then
        cmdSave.Enabled = True
    End If
End Sub

Private Sub optjaw_Click()
    If optjaw.Value = True Then
        cmdSave.Enabled = True
    End If
End Sub

Private Sub optlips_Click()
    If optlips.Value = True Then
        cmdSave.Enabled = True
    End If
End Sub

Private Sub optnose_Click()
    If optnose.Value = True Then
        cmdSave.Enabled = True
    End If
End Sub

Private Sub txth_Change()
    If Not IsNumeric(txth.Text) = True Then
        txth.Text = ""
        txth.SetFocus
    End If
End Sub

Private Sub txth_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        If txth.Text = "" Then
            txth.Text = h
            picpartout.Cls
            picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
        Else
            picpartout.Cls
            picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
        End If
    End If
End Sub

Private Sub txtw_Change()
    If Not IsNumeric(txtw.Text) = True Then
        txtw.Text = ""
        txtw.SetFocus
    End If
End Sub

Private Sub txtw_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        If txtw.Text = "" Then
            txtw.Text = w
            picpartout.Cls
            picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
        Else
            picpartout.Cls
            picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
        End If
    End If
End Sub

Private Sub txtx_Change()
    If Not IsNumeric(txtx.Text) = True Then
        txtx.Text = ""
        txtx.SetFocus
    End If
End Sub

Private Sub txtx_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtx.Text = "" Then
            txtx.Text = x
            picpartout.Cls
            picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
        Else
            picpartout.Cls
            picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
        End If
    End If
End Sub

Private Sub txty_Change()
    If Not IsNumeric(txty.Text) = True Then
        txty.Text = ""
        txty.SetFocus
    End If
End Sub

Private Sub txty_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        If txty.Text = "" Then
            txty.Text = y
            picpartout.Cls
            picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
        Else
            picpartout.Cls
            picpartout.PaintPicture picpart.Picture, txtx.Text, txty.Text, txtw.Text, txth.Text, , , , , vbSrcAnd
        End If
    End If
End Sub
