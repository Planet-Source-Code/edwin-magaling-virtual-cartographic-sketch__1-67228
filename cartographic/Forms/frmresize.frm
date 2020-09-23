VERSION 5.00
Begin VB.Form frmresize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resize"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   Icon            =   "frmresize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5205
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   315
      TabIndex        =   0
      Top             =   225
      Width           =   4515
      Begin VB.TextBox txtwidth 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   1485
         TabIndex        =   7
         Top             =   1260
         Width           =   735
      End
      Begin VB.TextBox txtheight 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   315
         Width           =   735
      End
      Begin VB.CommandButton cmdupheight 
         Height          =   555
         Left            =   2970
         Picture         =   "frmresize.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   810
         Width           =   555
      End
      Begin VB.CommandButton cmddownheight 
         Height          =   555
         Left            =   2970
         Picture         =   "frmresize.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1530
         Width           =   555
      End
      Begin VB.CommandButton cmdupwidth 
         Height          =   555
         Left            =   3555
         Picture         =   "frmresize.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1170
         Width           =   555
      End
      Begin VB.CommandButton cmddownwidth 
         Height          =   555
         Left            =   2385
         Picture         =   "frmresize.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1170
         Width           =   555
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   405
         TabIndex        =   9
         Top             =   1305
         Width           =   855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1710
         TabIndex        =   8
         Top             =   315
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&Ok"
      Height          =   420
      Left            =   3600
      TabIndex        =   1
      Top             =   2970
      Width           =   1230
   End
End
Attribute VB_Name = "frmresize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub resize()
With frmcartographic
    .picout.Cls
        'If .picjaw.Picture <> 0 Then 'jaw
        '    .picjawout.Picture = .picjaw.Picture
        '    .picout.PaintPicture .picjawout.Picture, .lblx1, .lbly1, .lblw1, .lblh1, , , , , vbSrcAnd
        If .picjawout.Picture <> 0 Then
            .picout.PaintPicture .picjawout.Picture, .lblx1, .lbly1, .lblw1, .lblh1, , , , , vbSrcAnd
        End If
        
        'If .pichair.Picture <> 0 Then 'hair
        '    .pichairout.Picture = .pichair.Picture
        '    .picout.PaintPicture .pichairout.Picture, .lblx2, .lbly2, .lblw2, .lblh2, , , , , vbSrcAnd
        If .pichairout.Picture <> 0 Then
            .picout.PaintPicture .pichairout.Picture, .lblx2, .lbly2, .lblw2, .lblh2, , , , , vbSrcAnd
        End If
        
        'If .picears.Picture <> 0 Then 'ears
        '    .picearsout.Picture = .picears.Picture
        '    .picout.PaintPicture .picearsout.Picture, .lblx3, .lbly3, .lblw3, .lblh3, , , , , vbSrcAnd
        If .picearsout.Picture <> 0 Then
            .picout.PaintPicture .picearsout.Picture, .lblx3, .lbly3, .lblw3, .lblh3, , , , , vbSrcAnd
        End If
            
        'If .picbrow.Picture <> 0 Then 'eyebrow
        '    .picbrowout.Picture = .picbrow.Picture
        '    .picout.PaintPicture .picbrowout.Picture, .lblx4, .lbly4, .lblw4, .lblh4, , , , , vbSrcAnd
        If .picbrowout.Picture <> 0 Then
            .picout.PaintPicture .picbrowout.Picture, .lblx4, .lbly4, .lblw4, .lblh4, , , , , vbSrcAnd
        End If
        
        'If .piceyes.Picture <> 0 Then 'eyes
        '    .piceyesout.Picture = .piceyes.Picture
        '    .picout.PaintPicture .piceyesout.Picture, .lblx5, .lbly5, .lblw5, .lblh5, , , , , vbSrcAnd
        If .piceyesout.Picture <> 0 Then
            .picout.PaintPicture .piceyesout.Picture, .lblx5, .lbly5, .lblw5, .lblh5, , , , , vbSrcAnd
        End If
        
        'If .picnose.Picture <> 0 Then 'nose
        '    .picnoseout.Picture = .picnose.Picture
        '    .picout.PaintPicture .picnoseout.Picture, .lblx6, .lbly6, .lblw6, .lblh6, , , , , vbSrcAnd
        If .picnoseout.Picture <> 0 Then
            .picout.PaintPicture .picnoseout.Picture, .lblx6, .lbly6, .lblw6, .lblh6, , , , , vbSrcAnd
        End If
    
        'If .piclips.Picture <> 0 Then 'lips
        '    .piclipsout.Picture = .piclips.Picture
        '    .picout.PaintPicture .piclipsout.Picture, .lblx7, .lbly7, .lblw7, .lblh7, , , , , vbSrcAnd
        If .piclipsout.Picture <> 0 Then
            .picout.PaintPicture .piclipsout.Picture, .lblx7, .lbly7, .lblw7, .lblh7, , , , , vbSrcAnd
        End If
    
        'If .picglass.Picture <> 0 Then 'glass
        '    .picglassout.Picture = .picglass.Picture
        '    .picout.PaintPicture .picglassout.Picture, .lblx8, .lbly8, .lblw8, .lblh8, , , , , vbSrcAnd
        If .picglassout.Picture <> 0 Then
            .picout.PaintPicture .picglassout.Picture, .lblx8, .lbly8, .lblw8, .lblh8, , , , , vbSrcAnd
        End If
        
        'If .picbeard.Picture <> 0 Then 'beard
        '    .picbeardout.Picture = .picbeard.Picture
        '    .picout.PaintPicture .picbeardout.Picture, .lblx9, .lbly9, .lblw9, .lblh9, , , , , vbSrcAnd
        If .picbeardout.Picture <> 0 Then
            .picout.PaintPicture .picbeardout.Picture, .lblx9, .lbly9, .lblw9, .lblh9, , , , , vbSrcAnd
        End If
        
        'If .piccap.Picture <> 0 Then 'cap
        '    .piccapout.Picture = .piccap.Picture
        '    .picout.PaintPicture .piccapout.Picture, .lblx10, .lbly10, .lblw10, .lblh10, , , , , vbSrcAnd
        If .piccapout.Picture <> 0 Then
            .picout.PaintPicture .piccapout.Picture, .lblx10, .lbly10, .lblw10, .lblh10, , , , , vbSrcAnd
        End If

End With
End Sub
Private Sub cmddownwidth_Click()
With frmcartographic

    If Me.Caption = "Jaw Resize" Then 'jaw
        txtwidth.Text = txtwidth.Text - 1
        .lblw1 = txtwidth.Text
        '.lbly1 = txtyaxis.Text
    ElseIf Me.Caption = "Hair Resize" Then 'hair
        txtwidth.Text = txtwidth.Text - 1
        .lblw2 = txtwidth.Text
        '.lbly2 = txtyaxis.Text
    ElseIf Me.Caption = "Ears Resize" Then 'ears
        txtwidth.Text = txtwidth.Text - 1
        .lblw3 = txtwidth.Text
        '.lbly3 = txtyaxis.Text
    ElseIf Me.Caption = "Eyebrow Resize" Then 'eyebrow
        txtwidth.Text = txtwidth.Text - 1
        .lblw4 = txtwidth.Text
        '.lbly4 = txtyaxis.Text
    ElseIf Me.Caption = "Eyes Resize" Then 'eyes
        txtwidth.Text = txtwidth.Text - 1
        .lblw5 = txtwidth.Text
        '.lbly5 = txtyaxis.Text
    ElseIf Me.Caption = "Nose Resize" Then 'nose
        txtwidth.Text = txtwidth.Text - 1
        .lblw6 = txtwidth.Text
        '.lbly6 = txtyaxis.Text
    ElseIf Me.Caption = "Lips Resize" Then 'lips
        txtwidth.Text = txtwidth.Text - 1
        .lblw7 = txtwidth.Text
        '.lbly7 = txtyaxis.Text
    ElseIf Me.Caption = "Glass Resize" Then 'glass
        txtwidth.Text = txtwidth.Text - 1
        .lblw8 = txtwidth.Text
        '.lbly8 = txtyaxis.Text
    ElseIf Me.Caption = "Beard Resize" Then 'beard
        txtwidth.Text = txtwidth.Text - 1
        .lblw9 = txtwidth.Text
        '.lbly8 = txtyaxis.Text
    ElseIf Me.Caption = "Cap Resize" Then 'cap
        txtwidth.Text = txtwidth.Text - 1
        .lblw10 = txtwidth.Text
        '.lbly8 = txtyaxis.Text
    End If
     
End With

    resize
    
End Sub

Private Sub cmddownheight_Click()
With frmcartographic

    If Me.Caption = "Jaw Resize" Then 'jaw
        txtheight.Text = txtheight.Text - 1
        .lblh1 = txtheight.Text
        '.lbly1 = txtyaxis.Text
    ElseIf Me.Caption = "Hair Resize" Then 'hair
        txtheight.Text = txtheight.Text - 1
        .lblh2 = txtheight.Text
        '.lbly2 = txtyaxis.Text
    ElseIf Me.Caption = "Ears Resize" Then 'ears
        txtheight.Text = txtheight.Text - 1
        .lblh3 = txtheight.Text
        '.lbly3 = txtyaxis.Text
    ElseIf Me.Caption = "Eyebrow Resize" Then 'eyebrow
        txtheight.Text = txtheight.Text - 1
        .lblh4 = txtheight.Text
        '.lbly4 = txtyaxis.Text
    ElseIf Me.Caption = "Eyes Resize" Then 'eyes
        txtheight.Text = txtheight.Text - 1
        .lblh5 = txtheight.Text
        '.lbly5 = txtyaxis.Text
    ElseIf Me.Caption = "Nose Resize" Then 'nose
        txtheight.Text = txtheight.Text - 1
        .lblh6 = txtheight.Text
        '.lbly6 = txtyaxis.Text
    ElseIf Me.Caption = "Lips Resize" Then 'lips
        txtheight.Text = txtheight.Text - 1
        .lblh7 = txtheight.Text
        '.lbly7 = txtyaxis.Text
    ElseIf Me.Caption = "Glass Resize" Then 'glass
        txtheight.Text = txtheight.Text - 1
        .lblh8 = txtheight.Text
        '.lbly8 = txtyaxis.Text
    ElseIf Me.Caption = "Beard Resize" Then 'beard
        txtheight.Text = txtheight.Text - 1
        .lblh9 = txtheight.Text
        '.lbly8 = txtyaxis.Text
    ElseIf Me.Caption = "Cap Resize" Then 'cap
        txtheight.Text = txtheight.Text - 1
        .lblh10 = txtheight.Text
        '.lbly8 = txtyaxis.Text
    End If
     
End With

    resize
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdupwidth_Click()
With frmcartographic

    If Me.Caption = "Jaw Resize" Then 'jaw
        txtwidth.Text = txtwidth.Text + 1
        .lblw1 = txtwidth.Text
        '.lbly1 = txtyaxis.Text
    ElseIf Me.Caption = "Hair Resize" Then 'hair
        txtwidth.Text = txtwidth.Text + 1
        .lblw2 = txtwidth.Text
        '.lbly2 = txtyaxis.Text
    ElseIf Me.Caption = "Ears Resize" Then 'ears
        txtwidth.Text = txtwidth.Text + 1
        .lblw3 = txtwidth.Text
        '.lbly3 = txtyaxis.Text
    ElseIf Me.Caption = "Eyebrow Resize" Then 'eyebrow
        txtwidth.Text = txtwidth.Text + 1
        .lblw4 = txtwidth.Text
        '.lbly4 = txtyaxis.Text
    ElseIf Me.Caption = "Eyes Resize" Then 'eyes
        txtwidth.Text = txtwidth.Text + 1
        .lblw5 = txtwidth.Text
        '.lbly5 = txtyaxis.Text
    ElseIf Me.Caption = "Nose Resize" Then 'nose
        txtwidth.Text = txtwidth.Text + 1
        .lblw6 = txtwidth.Text
        '.lbly6 = txtyaxis.Text
    ElseIf Me.Caption = "Lips Resize" Then 'lips
        txtwidth.Text = txtwidth.Text + 1
        .lblw7 = txtwidth.Text
        '.lbly7 = txtyaxis.Text
    ElseIf Me.Caption = "Glass Resize" Then 'glass
        txtwidth.Text = txtwidth.Text + 1
        .lblw8 = txtwidth.Text
        '.lbly8 = txtyaxis.Text
    ElseIf Me.Caption = "Beard Resize" Then 'beard
        txtwidth.Text = txtwidth.Text + 1
        .lblw9 = txtwidth.Text
        '.lbly8 = txtyaxis.Text
    ElseIf Me.Caption = "Cap Resize" Then 'cap
        txtwidth.Text = txtwidth.Text + 1
        .lblw10 = txtwidth.Text
        '.lbly8 = txtyaxis.Text
    End If
     
End With

    resize
    
End Sub

Private Sub cmdupheight_Click()
With frmcartographic

    If Me.Caption = "Jaw Resize" Then 'jaw
        txtheight.Text = txtheight.Text + 1
        .lblh1 = txtheight.Text
        '.lbly1 = txtyaxis.Text
    ElseIf Me.Caption = "Hair Resize" Then 'hair
        txtheight.Text = txtheight.Text + 1
        .lblh2 = txtheight.Text
        '.lbly2 = txtyaxis.Text
    ElseIf Me.Caption = "Ears Resize" Then 'ears
        txtheight.Text = txtheight.Text + 1
        .lblh3 = txtheight.Text
        '.lbly3 = txtyaxis.Text
    ElseIf Me.Caption = "Eyebrow Resize" Then 'eyebrow
        txtheight.Text = txtheight.Text + 1
        .lblh4 = txtheight.Text
        '.lbly4 = txtyaxis.Text
    ElseIf Me.Caption = "Eyes Resize" Then 'eyes
        txtheight.Text = txtheight.Text + 1
        .lblh5 = txtheight.Text
        '.lbly5 = txtyaxis.Text
    ElseIf Me.Caption = "Nose Resize" Then 'nose
        txtheight.Text = txtheight.Text + 1
        .lblh6 = txtheight.Text
        '.lbly6 = txtyaxis.Text
    ElseIf Me.Caption = "Lips Resize" Then 'lips
        txtheight.Text = txtheight.Text + 1
        .lblh7 = txtheight.Text
        '.lbly7 = txtyaxis.Text
    ElseIf Me.Caption = "Glass Resize" Then 'glass
        txtheight.Text = txtheight.Text + 1
        .lblh8 = txtheight.Text
        '.lbly8 = txtyaxis.Text
    ElseIf Me.Caption = "Beard Resize" Then 'beard
        txtheight.Text = txtheight.Text + 1
        .lblh9 = txtheight.Text
        '.lbly8 = txtyaxis.Text
    ElseIf Me.Caption = "Cap Resize" Then 'cap
        txtheight.Text = txtheight.Text + 1
        .lblh10 = txtheight.Text
        '.lbly8 = txtyaxis.Text
    End If
     
End With

    resize
End Sub

Private Sub txtheight_Change()
    If Not IsNumeric(txtheight.Text) = True Then
        txtheight.Text = ""
        txtheight.SetFocus
    End If
End Sub

Private Sub txtheight_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtheight = "" Then
        With frmcartographic
            If Me.Caption = "Jaw Resize" Then 'jaw
                txtheight.Text = .lblh1
            ElseIf Me.Caption = "Hair Resize" Then 'hair
                 txtheight.Text = .lblh2
            ElseIf Me.Caption = "Ears Resize" Then 'ears
                 txtheight.Text = .lblh3
            ElseIf Me.Caption = "Eyebrow Resize" Then 'eyebrow
                 txtheight.Text = .lblh4
            ElseIf Me.Caption = "Eyes Resize" Then 'eyes
                 txtheight.Text = .lblh5
            ElseIf Me.Caption = "Nose Resize" Then 'nose
                 txtheight.Text = .lblh6
            ElseIf Me.Caption = "Lips Resize" Then 'lips
                 txtheight.Text = .lblh7
            ElseIf Me.Caption = "Glass Resize" Then 'glass
                 txtheight.Text = .lblh8
            ElseIf Me.Caption = "Beard Resize" Then 'beard
                 txtheight.Text = .lblh9
            ElseIf Me.Caption = "Cap Resize" Then 'cap
                 txtheight.Text = .lblh10
            End If
            End With
       Exit Sub
    End If
    With frmcartographic
    
        If Me.Caption = "Jaw Resize" Then 'jaw
            .lblh1 = txtheight.Text
        ElseIf Me.Caption = "Hair Resize" Then 'hair
            .lblh2 = txtheight.Text
        ElseIf Me.Caption = "Ears Resize" Then 'ears
            .lblh3 = txtheight.Text
        ElseIf Me.Caption = "Eyebrow Resize" Then 'eyebrow
            .lblh4 = txtheight.Text
        ElseIf Me.Caption = "Eyes Resize" Then 'eyes
            .lblh5 = txtheight.Text
        ElseIf Me.Caption = "Nose Resize" Then 'nose
            .lblh6 = txtheight.Text
        ElseIf Me.Caption = "Lips Resize" Then 'lips
            .lblh7 = txtheight.Text
        ElseIf Me.Caption = "Glass Resize" Then 'glass
            .lblh8 = txtheight.Text
        ElseIf Me.Caption = "Beard Resize" Then 'beard
            .lblh9 = txtheight.Text
        ElseIf Me.Caption = "Cap Resize" Then 'cap
            .lblh10 = txtheight.Text
        End If
     
    End With

    resize
    
End If
End Sub

Private Sub txtwidth_Change()
     If Not IsNumeric(txtwidth.Text) = True Then
        txtwidth.Text = ""
        txtwidth.SetFocus
    End If
End Sub

Private Sub txtwidth_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtwidth = "" Then
       With frmcartographic
            If Me.Caption = "Jaw Resize" Then 'jaw
                txtwidth.Text = .lblw1
            ElseIf Me.Caption = "Hair Resize" Then 'hair
                 txtwidth.Text = .lblw2
            ElseIf Me.Caption = "Ears Resize" Then 'ears
                 txtwidth.Text = .lblw3
            ElseIf Me.Caption = "Eyebrow Resize" Then 'eyebrow
                 txtwidth.Text = .lblw4
            ElseIf Me.Caption = "Eyes Resize" Then 'eyes
                 txtwidth.Text = .lblw5
            ElseIf Me.Caption = "Nose Resize" Then 'nose
                 txtwidth.Text = .lblw6
            ElseIf Me.Caption = "Lips Resize" Then 'lips
                 txtwidth.Text = .lblw7
            ElseIf Me.Caption = "Glass Resize" Then 'glass
                 txtwidth.Text = .lblw8
            ElseIf Me.Caption = "Beard Resize" Then 'beard
                 txtwidth.Text = .lblw9
            ElseIf Me.Caption = "Cap Resize" Then 'cap
                 txtwidth.Text = .lblw10
            End If
        End With
       Exit Sub
    End If
    With frmcartographic

        If Me.Caption = "Jaw Resize" Then 'jaw
            .lblw1 = txtwidth.Text
        ElseIf Me.Caption = "Hair Resize" Then 'hair
            .lblw2 = txtwidth.Text
        ElseIf Me.Caption = "Ears Resize" Then 'ears
            .lblw3 = txtwidth.Text
        ElseIf Me.Caption = "Eyebrow Resize" Then 'eyebrow
            .lblw4 = txtwidth.Text
        ElseIf Me.Caption = "Eyes Resize" Then 'eyes
            .lblw5 = txtwidth.Text
        ElseIf Me.Caption = "Nose Resize" Then 'nose
            .lblw6 = txtwidth.Text
        ElseIf Me.Caption = "Lips Resize" Then 'lips
            .lblw7 = txtwidth.Text
        ElseIf Me.Caption = "Glass Resize" Then 'glass
            .lblw8 = txtwidth.Text
        ElseIf Me.Caption = "Beard Resize" Then 'beard
            .lblw9 = txtwidth.Text
        ElseIf Me.Caption = "Cap Resize" Then 'cap
            .lblw10 = txtwidth.Text
        End If
     
    End With

    resize
End If
End Sub
