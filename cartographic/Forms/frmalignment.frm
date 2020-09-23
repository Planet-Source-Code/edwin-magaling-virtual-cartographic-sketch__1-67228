VERSION 5.00
Begin VB.Form frmalignment 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alignment"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   Icon            =   "frmalignment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5160
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      Caption         =   "&Ok"
      Height          =   420
      Left            =   3600
      TabIndex        =   1
      Top             =   2970
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   315
      TabIndex        =   0
      Top             =   225
      Width           =   4515
      Begin VB.CommandButton cmddownX 
         Height          =   555
         Left            =   2385
         Picture         =   "frmalignment.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1170
         Width           =   555
      End
      Begin VB.CommandButton cmdupX 
         Height          =   555
         Left            =   3555
         Picture         =   "frmalignment.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1170
         Width           =   555
      End
      Begin VB.CommandButton cmddownY 
         Height          =   555
         Left            =   2970
         Picture         =   "frmalignment.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1530
         Width           =   555
      End
      Begin VB.CommandButton cmdupY 
         Height          =   555
         Left            =   2970
         Picture         =   "frmalignment.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtyaxis 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   315
         Width           =   735
      End
      Begin VB.TextBox txtxaxis 
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
         TabIndex        =   4
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Y - Axis : "
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
         Left            =   1530
         TabIndex        =   3
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "X - Axis : "
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
         Left            =   270
         TabIndex        =   2
         Top             =   1260
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmalignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alignment()
With frmcartographic
    .picout.Cls
        'If .picjaw.Picture <> 0 Then 'jaw
            '.picjawout.Picture = .picjaw.Picture
            '.picout.PaintPicture .picjawout.Picture, .lblx1, .lbly1, .lblw1, .lblh1, , , , , vbSrcAnd
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

Private Sub cmddownX_Click()
With frmcartographic

    If Me.Caption = "Jaw Alignment" Then 'jaw
        txtxaxis.Text = txtxaxis.Text - 1
        .lblx1 = txtxaxis.Text
        '.lbly1 = txtyaxis.Text
    ElseIf Me.Caption = "Hair Alignment" Then 'hair
        txtxaxis.Text = txtxaxis.Text - 1
        .lblx2 = txtxaxis.Text
        '.lbly2 = txtyaxis.Text
    ElseIf Me.Caption = "Ears Alignment" Then 'ears
        txtxaxis.Text = txtxaxis.Text - 1
        .lblx3 = txtxaxis.Text
        '.lbly3 = txtyaxis.Text
    ElseIf Me.Caption = "Eyebrow Alignment" Then 'eyebrow
        txtxaxis.Text = txtxaxis.Text - 1
        .lblx4 = txtxaxis.Text
        '.lbly4 = txtyaxis.Text
    ElseIf Me.Caption = "Eyes Alignment" Then 'eyes
        txtxaxis.Text = txtxaxis.Text - 1
        .lblx5 = txtxaxis.Text
        '.lbly5 = txtyaxis.Text
    ElseIf Me.Caption = "Nose Alignment" Then 'nose
        txtxaxis.Text = txtxaxis.Text - 1
        .lblx6 = txtxaxis.Text
        '.lbly6 = txtyaxis.Text
    ElseIf Me.Caption = "Lips Alignment" Then 'lips
        txtxaxis.Text = txtxaxis.Text - 1
        .lblx7 = txtxaxis.Text
        '.lbly7 = txtyaxis.Text
    ElseIf Me.Caption = "Glass Alignment" Then 'glass
        txtxaxis.Text = txtxaxis.Text - 1
        .lblx8 = txtxaxis.Text
        '.lbly8 = txtyaxis.Text
    ElseIf Me.Caption = "Beard Alignment" Then 'beard
        txtxaxis.Text = txtxaxis.Text - 1
        .lblx9 = txtxaxis.Text
        '.lbly8 = txtyaxis.Text
    ElseIf Me.Caption = "Cap Alignment" Then 'cap
        txtxaxis.Text = txtxaxis.Text - 1
        .lblx10 = txtxaxis.Text
        '.lbly8 = txtyaxis.Text
    End If
        
End With

    alignment
    
End Sub

Private Sub cmddownY_Click()
With frmcartographic
    If Me.Caption = "Jaw Alignment" Then 'jaw
        txtyaxis.Text = txtyaxis.Text + 1
        '.lblx1 = txtxaxis.Text
        .lbly1 = txtyaxis.Text
    ElseIf Me.Caption = "Hair Alignment" Then 'hair
        txtyaxis.Text = txtyaxis.Text + 1
        '.lblx2 = txtxaxis.Text
        .lbly2 = txtyaxis.Text
    ElseIf Me.Caption = "Ears Alignment" Then 'ears
        txtyaxis.Text = txtyaxis.Text + 1
        '.lblx3 = txtxaxis.Text
        .lbly3 = txtyaxis.Text
    ElseIf Me.Caption = "Eyebrow Alignment" Then 'eyebrow
        txtyaxis.Text = txtyaxis.Text + 1
        '.lblx4 = txtxaxis.Text
        .lbly4 = txtyaxis.Text
    ElseIf Me.Caption = "Eyes Alignment" Then 'eyes
        txtyaxis.Text = txtyaxis.Text + 1
        '.lblx5 = txtxaxis.Text
        .lbly5 = txtyaxis.Text
    ElseIf Me.Caption = "Nose Alignment" Then 'nose
        txtyaxis.Text = txtyaxis.Text + 1
        '.lblx6 = txtxaxis.Text
        .lbly6 = txtyaxis.Text
    ElseIf Me.Caption = "Lips Alignment" Then 'lips
        txtyaxis.Text = txtyaxis.Text + 1
        '.lblx7 = txtxaxis.Text
        .lbly7 = txtyaxis.Text
    ElseIf Me.Caption = "Glass Alignment" Then 'glass
        txtyaxis.Text = txtyaxis.Text + 1
        '.lblx8 = txtxaxis.Text
        .lbly8 = txtyaxis.Text
    ElseIf Me.Caption = "Beard Alignment" Then 'beard
        txtyaxis.Text = txtyaxis.Text + 1
        '.lblx8 = txtxaxis.Text
        .lbly9 = txtyaxis.Text
    ElseIf Me.Caption = "Cap Alignment" Then 'cap
        txtyaxis.Text = txtyaxis.Text + 1
        '.lblx8 = txtxaxis.Text
        .lbly10 = txtyaxis.Text
    End If

End With

    alignment
    
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdupX_Click()
With frmcartographic

    If Me.Caption = "Jaw Alignment" Then 'jaw
        txtxaxis.Text = txtxaxis.Text + 1
        .lblx1 = txtxaxis.Text
        '.lbly1 = txtyaxis.Text
    ElseIf Me.Caption = "Hair Alignment" Then 'hair
        txtxaxis.Text = txtxaxis.Text + 1
        .lblx2 = txtxaxis.Text
        '.lbly2 = txtyaxis.Text
    ElseIf Me.Caption = "Ears Alignment" Then 'ears
        txtxaxis.Text = txtxaxis.Text + 1
        .lblx3 = txtxaxis.Text
        '.lbly3 = txtyaxis.Text
    ElseIf Me.Caption = "Eyebrow Alignment" Then 'eyebrow
        txtxaxis.Text = txtxaxis.Text + 1
        .lblx4 = txtxaxis.Text
        '.lbly4 = txtyaxis.Text
    ElseIf Me.Caption = "Eyes Alignment" Then 'eyes
        txtxaxis.Text = txtxaxis.Text + 1
        .lblx5 = txtxaxis.Text
        '.lbly5 = txtyaxis.Text
    ElseIf Me.Caption = "Nose Alignment" Then 'nose
        txtxaxis.Text = txtxaxis.Text + 1
        .lblx6 = txtxaxis.Text
        '.lbly6 = txtyaxis.Text
    ElseIf Me.Caption = "Lips Alignment" Then 'lips
        txtxaxis.Text = txtxaxis.Text + 1
        .lblx7 = txtxaxis.Text
        '.lbly7 = txtyaxis.Text
    ElseIf Me.Caption = "Glass Alignment" Then 'glass
        txtxaxis.Text = txtxaxis.Text + 1
        .lblx8 = txtxaxis.Text
        '.lbly8 = txtyaxis.Text
    ElseIf Me.Caption = "Beard Alignment" Then 'beard
        txtxaxis.Text = txtxaxis.Text + 1
        .lblx9 = txtxaxis.Text
        '.lbly8 = txtyaxis.Text
    ElseIf Me.Caption = "Cap Alignment" Then 'cap
        txtxaxis.Text = txtxaxis.Text + 1
        .lblx10 = txtxaxis.Text
        '.lbly8 = txtyaxis.Text
    End If
     
End With

    alignment

End Sub

Private Sub cmdupY_Click()
With frmcartographic

    If Me.Caption = "Jaw Alignment" Then 'jaw
        txtyaxis.Text = txtyaxis.Text - 1
        '.lblx1 = txtxaxis.Text
        .lbly1 = txtyaxis.Text
    ElseIf Me.Caption = "Hair Alignment" Then 'hair
        txtyaxis.Text = txtyaxis.Text - 1
        '.lblx2 = txtxaxis.Text
        .lbly2 = txtyaxis.Text
    ElseIf Me.Caption = "Ears Alignment" Then 'ears
        txtyaxis.Text = txtyaxis.Text - 1
        '.lblx3 = txtxaxis.Text
        .lbly3 = txtyaxis.Text
    ElseIf Me.Caption = "Eyebrow Alignment" Then 'eyebrow
        txtyaxis.Text = txtyaxis.Text - 1
        '.lblx4 = txtxaxis.Text
        .lbly4 = txtyaxis.Text
    ElseIf Me.Caption = "Eyes Alignment" Then 'eyes
        txtyaxis.Text = txtyaxis.Text - 1
        '.lblx5 = txtxaxis.Text
        .lbly5 = txtyaxis.Text
    ElseIf Me.Caption = "Nose Alignment" Then 'nose
        txtyaxis.Text = txtyaxis.Text - 1
        '.lblx6 = txtxaxis.Text
        .lbly6 = txtyaxis.Text
    ElseIf Me.Caption = "Lips Alignment" Then 'lips
        txtyaxis.Text = txtyaxis.Text - 1
        '.lblx7 = txtxaxis.Text
        .lbly7 = txtyaxis.Text
    ElseIf Me.Caption = "Glass Alignment" Then 'glass
        txtyaxis.Text = txtyaxis.Text - 1
        '.lblx8 = txtxaxis.Text
        .lbly8 = txtyaxis.Text
    ElseIf Me.Caption = "Beard Alignment" Then 'beard
        txtyaxis.Text = txtyaxis.Text - 1
        '.lblx8 = txtxaxis.Text
        .lbly9 = txtyaxis.Text
    ElseIf Me.Caption = "Cap Alignment" Then 'cap
        txtyaxis.Text = txtyaxis.Text - 1
        '.lblx8 = txtxaxis.Text
        .lbly10 = txtyaxis.Text
    End If
    
     
End With

    alignment
End Sub

Private Sub txtxaxis_Change()
    If Not IsNumeric(txtxaxis.Text) = True Then
        txtxaxis.Text = ""
        txtxaxis.SetFocus
    End If
End Sub

Private Sub txtxaxis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtxaxis = "" Then
            With frmcartographic
            If Me.Caption = "Jaw Alignment" Then 'jaw
                txtxaxis.Text = .lblx1
            ElseIf Me.Caption = "Hair Alignment" Then 'hair
                 txtxaxis.Text = .lblx2
            ElseIf Me.Caption = "Ears Alignment" Then 'ears
                 txtxaxis.Text = .lblx3
            ElseIf Me.Caption = "Eyebrow Alignment" Then 'eyebrow
                 txtxaxis.Text = .lblx4
            ElseIf Me.Caption = "Eyes Alignment" Then 'eyes
                 txtxaxis.Text = .lblx5
            ElseIf Me.Caption = "Nose Alignment" Then 'nose
                 txtxaxis.Text = .lblx6
            ElseIf Me.Caption = "Lips Alignment" Then 'lips
                 txtxaxis.Text = .lblx7
            ElseIf Me.Caption = "Glass Alignment" Then 'glass
                 txtxaxis.Text = .lblx8
            ElseIf Me.Caption = "Beard Alignment" Then 'beard
                 txtxaxis.Text = .lblx9
            ElseIf Me.Caption = "Cap Alignment" Then 'cap
                 txtxaxis.Text = .lblx10
            End If
            End With
            Exit Sub
        End If
        
        With frmcartographic

        If Me.Caption = "Jaw Alignment" Then 'jaw
            .lblx1 = txtxaxis.Text
        ElseIf Me.Caption = "Hair Alignment" Then 'hair
            .lblx2 = txtxaxis.Text
        ElseIf Me.Caption = "Ears Alignment" Then 'ears
            .lblx3 = txtxaxis.Text
        ElseIf Me.Caption = "Eyebrow Alignment" Then 'eyebrow
            .lblx4 = txtxaxis.Text
        ElseIf Me.Caption = "Eyes Alignment" Then 'eyes
            .lblx5 = txtxaxis.Text
        ElseIf Me.Caption = "Nose Alignment" Then 'nose
            .lblx6 = txtxaxis.Text
        ElseIf Me.Caption = "Lips Alignment" Then 'lips
            .lblx7 = txtxaxis.Text
        ElseIf Me.Caption = "Glass Alignment" Then 'glass
            .lblx8 = txtxaxis.Text
        ElseIf Me.Caption = "Beard Alignment" Then 'beard
            .lblx9 = txtxaxis.Text
        ElseIf Me.Caption = "Cap Alignment" Then 'cap
            .lblx10 = txtxaxis.Text
        End If
        End With
        
        alignment
        
    End If
End Sub

Private Sub txtyaxis_Change()
    If Not IsNumeric(txtyaxis.Text) = True Then
        txtyaxis.Text = ""
        txtyaxis.SetFocus
    End If
End Sub

Private Sub txtyaxis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtyaxis = "" Then
        With frmcartographic
            If Me.Caption = "Jaw Alignment" Then 'jaw
                txtyaxis.Text = .lbly1
            ElseIf Me.Caption = "Hair Alignment" Then 'hair
                 txtyaxis.Text = .lbly2
            ElseIf Me.Caption = "Ears Alignment" Then 'ears
                 txtyaxis.Text = .lbly3
            ElseIf Me.Caption = "Eyebrow Alignment" Then 'eyebrow
                 txtyaxis.Text = .lbly4
            ElseIf Me.Caption = "Eyes Alignment" Then 'eyes
                 txtyaxis.Text = .lbly5
            ElseIf Me.Caption = "Nose Alignment" Then 'nose
                 txtyaxis.Text = .lbly6
            ElseIf Me.Caption = "Lips Alignment" Then 'lips
                 txtyaxis.Text = .lbly7
            ElseIf Me.Caption = "Glass Alignment" Then 'glass
                 txtyaxis.Text = .lbly8
            ElseIf Me.Caption = "Beard Alignment" Then 'beard
                 txtyaxis.Text = .lbly9
            ElseIf Me.Caption = "Cap Alignment" Then 'cap
                 txtyaxis.Text = .lbly10
            End If
            End With
            Exit Sub
        End If
        
        With frmcartographic
    
        If Me.Caption = "Jaw Alignment" Then 'jaw
            .lbly1 = txtyaxis.Text
        ElseIf Me.Caption = "Hair Alignment" Then 'hair
            .lbly2 = txtyaxis.Text
        ElseIf Me.Caption = "Ears Alignment" Then 'ears
            .lbly3 = txtyaxis.Text
        ElseIf Me.Caption = "Eyebrow Alignment" Then 'eyebrow
            .lbly4 = txtyaxis.Text
        ElseIf Me.Caption = "Eyes Alignment" Then 'eyes
            .lbly5 = txtyaxis.Text
        ElseIf Me.Caption = "Nose Alignment" Then 'nose
            .lbly6 = txtyaxis.Text
        ElseIf Me.Caption = "Lips Alignment" Then 'lips
            .lbly7 = txtyaxis.Text
        ElseIf Me.Caption = "Glass Alignment" Then 'glass
            .lbly8 = txtyaxis.Text
        ElseIf Me.Caption = "Beard Alignment" Then 'beard
            .lbly9 = txtyaxis.Text
        ElseIf Me.Caption = "Cap Alignment" Then 'cap
            .lbly10 = txtyaxis.Text
        End If
        
        End With
        
        alignment
        
    End If
End Sub


