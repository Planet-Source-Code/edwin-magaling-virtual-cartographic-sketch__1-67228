VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcartographic 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Cartograpic Sketch - Sketching"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   13530
   Icon            =   "frmcartographic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   13530
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox piccapout 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6525
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   85
      Top             =   8100
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox piccap 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6300
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   84
      Top             =   8100
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picbeardout 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6525
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   79
      Top             =   7650
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picbeard 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6300
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   78
      Top             =   7650
      Visible         =   0   'False
      Width           =   180
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   600
      Left            =   5355
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "jaw"
            Object.ToolTipText     =   "Jaw"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "hair"
            Object.ToolTipText     =   "Hair"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ears"
            Object.ToolTipText     =   "Ears"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "brow"
            Object.ToolTipText     =   "Eyebrow"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "eyes"
            Object.ToolTipText     =   "Eyes"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nose"
            Object.ToolTipText     =   "Nose"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "lips"
            Object.ToolTipText     =   "Lips"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "glass"
            Object.ToolTipText     =   "Glass"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "beard"
            Object.ToolTipText     =   "Beard"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cap"
            Object.ToolTipText     =   "Cap"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox picglass 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6300
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   22
      Top             =   7200
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picglassout 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6525
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   21
      Top             =   7200
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox piclipsout 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6525
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   20
      Top             =   6750
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox piclips 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6300
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   19
      Top             =   6750
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picnoseout 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6525
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   18
      Top             =   6300
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picnose 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6300
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   17
      Top             =   6300
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox piceyesout 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6525
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   16
      Top             =   5850
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox piceyes 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6300
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   15
      Top             =   5850
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picbrow 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6300
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   14
      Top             =   5355
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picbrowout 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6525
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   13
      Top             =   5355
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picearsout 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6525
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   12
      Top             =   4860
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox pichairout 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6525
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   11
      Top             =   4365
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picjawout 
      AutoRedraw      =   -1  'True
      Height          =   420
      Left            =   6525
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   10
      Top             =   3825
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picears 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6300
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   9
      Top             =   4860
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox pichair 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6300
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   8
      Top             =   4365
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picjaw 
      AutoRedraw      =   -1  'True
      Height          =   420
      Left            =   6300
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   7
      Top             =   3825
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.CommandButton cmdselect 
      Caption         =   "&Select"
      Height          =   825
      Left            =   6300
      Picture         =   "frmcartographic.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2655
      Width           =   915
   End
   Begin VB.CommandButton cmdprevious 
      Caption         =   "<< &Previous"
      Height          =   375
      Left            =   225
      TabIndex        =   5
      Top             =   720
      Width           =   1230
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "&Next >>"
      Height          =   375
      Left            =   4950
      TabIndex        =   4
      Top             =   765
      Width           =   1230
   End
   Begin VB.PictureBox picout 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7530
      Left            =   7290
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   3
      Top             =   1305
      Width           =   6045
   End
   Begin VB.PictureBox picview 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7530
      Left            =   180
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   2
      Top             =   1350
      Width           =   6045
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7380
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   -2147483643
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":0884
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":14D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":212C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":2D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":39D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":4628
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":527C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":5ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":7522
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":A048
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7965
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":1EFBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":1FC13
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":20867
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":214BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":2210F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":22D63
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcartographic.frx":239B7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "New"
            Object.ToolTipText     =   "Creates new portrait."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Open"
            Object.ToolTipText     =   "Opens a saved portrait."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Save"
            Object.ToolTipText     =   "Saves the current work."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Print"
            Object.ToolTipText     =   "Prints the portrait"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblh10 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7155
      TabIndex        =   89
      Top             =   8235
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblw10 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7020
      TabIndex        =   88
      Top             =   8235
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbly10 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6885
      TabIndex        =   87
      Top             =   8235
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblx10 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6750
      TabIndex        =   86
      Top             =   8235
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblh9 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7155
      TabIndex        =   83
      Top             =   7785
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblw9 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7020
      TabIndex        =   82
      Top             =   7785
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbly9 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6885
      TabIndex        =   81
      Top             =   7785
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblx9 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6750
      TabIndex        =   80
      Top             =   7785
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblcap 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6930
      TabIndex        =   77
      Top             =   2295
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblcapout 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7110
      TabIndex        =   76
      Top             =   2295
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblbeard 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6930
      TabIndex        =   75
      Top             =   2025
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblbeardout 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7110
      TabIndex        =   74
      Top             =   2025
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblglassout 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6795
      TabIndex        =   73
      Top             =   2295
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblglass 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6615
      TabIndex        =   72
      Top             =   2295
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbllipsout 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6795
      TabIndex        =   71
      Top             =   2070
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbllips 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6615
      TabIndex        =   70
      Top             =   2070
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblnoseout 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6795
      TabIndex        =   69
      Top             =   1800
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblnose 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6615
      TabIndex        =   68
      Top             =   1800
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbleyesout 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6795
      TabIndex        =   67
      Top             =   1530
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbleyes 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6615
      TabIndex        =   66
      Top             =   1530
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblbrowout 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6795
      TabIndex        =   65
      Top             =   1305
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblbrow 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6615
      TabIndex        =   64
      Top             =   1305
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblearsout 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6795
      TabIndex        =   63
      Top             =   1035
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblears 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6615
      TabIndex        =   62
      Top             =   1035
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbljawout 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6795
      TabIndex        =   61
      Top             =   495
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbljaw 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6615
      TabIndex        =   60
      Top             =   495
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblhair 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6615
      TabIndex        =   59
      Top             =   765
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblhairout 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6795
      TabIndex        =   58
      Top             =   765
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblx5 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6750
      TabIndex        =   57
      Top             =   5850
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbly5 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6885
      TabIndex        =   56
      Top             =   5850
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblw5 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7020
      TabIndex        =   55
      Top             =   5850
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblh5 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7155
      TabIndex        =   54
      Top             =   5850
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblh6 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7155
      TabIndex        =   53
      Top             =   6345
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblw6 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7020
      TabIndex        =   52
      Top             =   6345
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbly6 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6885
      TabIndex        =   51
      Top             =   6345
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblx6 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6750
      TabIndex        =   50
      Top             =   6345
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblx7 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6750
      TabIndex        =   49
      Top             =   6840
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbly7 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6885
      TabIndex        =   48
      Top             =   6840
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblw7 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7020
      TabIndex        =   47
      Top             =   6840
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblh7 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7155
      TabIndex        =   46
      Top             =   6840
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblx8 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6750
      TabIndex        =   45
      Top             =   7335
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbly8 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6885
      TabIndex        =   44
      Top             =   7335
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblw8 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7020
      TabIndex        =   43
      Top             =   7335
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblh8 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7155
      TabIndex        =   42
      Top             =   7335
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblh4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7155
      TabIndex        =   41
      Top             =   5445
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblw4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7020
      TabIndex        =   40
      Top             =   5445
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbly4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6885
      TabIndex        =   39
      Top             =   5445
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblx4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6750
      TabIndex        =   38
      Top             =   5445
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblh3 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7155
      TabIndex        =   37
      Top             =   4950
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblw3 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7020
      TabIndex        =   36
      Top             =   4950
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbly3 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6885
      TabIndex        =   35
      Top             =   4950
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblx3 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6750
      TabIndex        =   34
      Top             =   4950
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblx2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6750
      TabIndex        =   33
      Top             =   4455
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbly2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6885
      TabIndex        =   32
      Top             =   4455
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblw2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7020
      TabIndex        =   31
      Top             =   4455
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblh2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7155
      TabIndex        =   30
      Top             =   4455
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblh1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7155
      TabIndex        =   29
      Top             =   3960
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblw1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7020
      TabIndex        =   28
      Top             =   3960
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbly1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6885
      TabIndex        =   27
      Top             =   3960
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblx1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6795
      TabIndex        =   26
      Top             =   3960
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OUTPUT SKETCH"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9270
      TabIndex        =   25
      Top             =   765
      Width           =   2010
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SKETCH SELECTION"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1935
      TabIndex        =   24
      Top             =   765
      Width           =   2370
   End
   Begin VB.Label tmp 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8955
      TabIndex        =   23
      Top             =   495
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufilenew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnufileopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnufilesep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufilesave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnufilesaveas 
         Caption         =   "Save as..."
      End
      Begin VB.Menu mnufilesep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileprint 
         Caption         =   "&Print.."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnufilesep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "&Close"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnueditremove 
         Caption         =   "&Remove"
         Begin VB.Menu mnueditremovejaw 
            Caption         =   "Jaw"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnueditremovehair 
            Caption         =   "Hair"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnueditremoveears 
            Caption         =   "Ears"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnueditremovebrow 
            Caption         =   "Eyebrow"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnueditremoveeyes 
            Caption         =   "Eyes"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnueditremovenose 
            Caption         =   "Nose"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnueditremovelips 
            Caption         =   "Lips"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnueditremoveglass 
            Caption         =   "Glass"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnueditremovebeard 
            Caption         =   "Beard"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnueditremovecap 
            Caption         =   "Cap"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuformat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuformatalign 
         Caption         =   "&Alignment"
         Begin VB.Menu mnuformatalignjaw 
            Caption         =   "Jaw"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatalignhair 
            Caption         =   "Hair"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatalignears 
            Caption         =   "Ears"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatalignbrow 
            Caption         =   "Eyebrow"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformataligneyes 
            Caption         =   "Eyes"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatalignnose 
            Caption         =   "Nose"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatalignlips 
            Caption         =   "Lips"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatalignglass 
            Caption         =   "Glass"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatalignbeard 
            Caption         =   "Beard"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformataligncap 
            Caption         =   "Cap"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuformatsize 
         Caption         =   "&Resize"
         Begin VB.Menu mnuformatsizejaw 
            Caption         =   "Jaw"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatsizehair 
            Caption         =   "Hair"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatsizeears 
            Caption         =   "Ears"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatsizebrow 
            Caption         =   "Eyebrow"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatsizeeyes 
            Caption         =   "Eyes"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatsizenose 
            Caption         =   "Nose"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatsizelips 
            Caption         =   "Lips"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatsizeglass 
            Caption         =   "Glass"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatsizebeard 
            Caption         =   "Beard"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuformatsizecap 
            Caption         =   "Cap"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnutoolsparts 
         Caption         =   "Head Parts Manager"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmcartographic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Saving As Boolean 'Is The App Saving A File?
Dim File As String ' holds the file name
Dim x1, y1, w1, h1 As Integer 'variables of jaw alignment & sizes
Dim x2, y2, w2, h2 As Integer 'variables of hair alignment & sizes
Dim x3, y3, w3, h3 As Integer 'variables of ears alignment & sizes
Dim x4, y4, w4, h4 As Integer 'variables of eyebrow alignment & sizes
Dim x5, y5, w5, h5 As Integer 'variables of eyes alignment & sizes
Dim x6, y6, w6, h6 As Integer 'variables of nose alignment & sizes
Dim x7, y7, w7, h7 As Integer 'variables of lips alignment & sizes
Dim x8, y8, w8, h8 As Integer 'variables of glass alignment & sizes
Dim x9, y9, w9, h9 As Integer 'variables of beard alignment & sizes
Dim x10, y10, w10, h10 As Integer 'variables of cap alignment & sizes
Dim xsketchby As String
Dim xusername, xuserlevel As String

Private Sub clearall()
    picview.Cls
    picout.Cls
    tmp.Caption = Empty

    picjaw.Picture = Nothing
    picjawout.Picture = Nothing
    lblx1.Caption = Empty
    lbly1.Caption = Empty
    lblw1.Caption = Empty
    lblh1.Caption = Empty
    lbljaw.Caption = Empty
    lbljawout.Caption = Empty
    
    pichair.Picture = Nothing
    pichairout.Picture = Nothing
    lblx2.Caption = Empty
    lbly2.Caption = Empty
    lblw2.Caption = Empty
    lblh2.Caption = Empty
    lblhair.Caption = Empty
    lblhairout.Caption = Empty

    picears.Picture = Nothing
    picearsout.Picture = Nothing
    lblx3.Caption = Empty
    lbly3.Caption = Empty
    lblw3.Caption = Empty
    lblh3.Caption = Empty
    lblears.Caption = Empty
    lblearsout.Caption = Empty
    
    picbrow.Picture = Nothing
    picbrowout.Picture = Nothing
    lblx4.Caption = Empty
    lbly4.Caption = Empty
    lblw4.Caption = Empty
    lblh4.Caption = Empty
    lblbrow.Caption = Empty
    lblbrowout.Caption = Empty
    
    piceyes.Picture = Nothing
    piceyesout.Picture = Nothing
    lblx5.Caption = Empty
    lbly5.Caption = Empty
    lblw5.Caption = Empty
    lblh5.Caption = Empty
    lbleyes.Caption = Empty
    lbleyesout.Caption = Empty
    
    picnose.Picture = Nothing
    picnoseout.Picture = Nothing
    lblx6.Caption = Empty
    lbly6.Caption = Empty
    lblw6.Caption = Empty
    lblh6.Caption = Empty
    lblnose.Caption = Empty
    lblnoseout.Caption = Empty

    piclips.Picture = Nothing
    piclipsout.Picture = Nothing
    lblx7.Caption = Empty
    lbly7.Caption = Empty
    lblw7.Caption = Empty
    lblh7.Caption = Empty
    lbllips.Caption = Empty
    lbllipsout.Caption = Empty

    picglass.Picture = Nothing
    picglassout.Picture = Nothing
    lblx8.Caption = Empty
    lbly8.Caption = Empty
    lblw8.Caption = Empty
    lblh8.Caption = Empty
    lblglass.Caption = Empty
    lblglassout.Caption = Empty
    
    picbeard.Picture = Nothing
    picbeardout.Picture = Nothing
    lblx9.Caption = Empty
    lbly9.Caption = Empty
    lblw9.Caption = Empty
    lblh9.Caption = Empty
    lblbeard.Caption = Empty
    lblbeardout.Caption = Empty
    
    piccap.Picture = Nothing
    piccapout.Picture = Nothing
    lblx10.Caption = Empty
    lbly10.Caption = Empty
    lblw10.Caption = Empty
    lblh10.Caption = Empty
    lblcap.Caption = Empty
    lblcapout.Caption = Empty
    
    With Toolbar1.Buttons
        .Item(3).Enabled = False
        .Item(4).Enabled = False
    End With
    
    cmdnext.Enabled = False
    cmdprevious.Enabled = False
    cmdselect.Enabled = False

End Sub

Private Sub remove()
    picout.Cls
    If picjawout.Picture <> 0 Then
        picout.PaintPicture picjawout.Picture, lblx1, lbly1, lblw1, lblh1, , , , , vbSrcAnd
    End If
    If pichairout.Picture <> 0 Then
        picout.PaintPicture pichairout.Picture, lblx2, lbly2, lblw2, lblh2, , , , , vbSrcAnd
    End If
    If picearsout.Picture <> 0 Then
        picout.PaintPicture picearsout.Picture, lblx3, lbly3, lblw3, lblh3, , , , , vbSrcAnd
    End If
    If picbrowout.Picture <> 0 Then
        picout.PaintPicture picbrowout.Picture, lblx4, lbly4, lblw4, lblh4, , , , , vbSrcAnd
    End If
    If piceyesout.Picture <> 0 Then
        picout.PaintPicture piceyesout.Picture, lblx5, lbly5, lblw5, lblh5, , , , , vbSrcAnd
    End If
    If picnoseout.Picture <> 0 Then
        picout.PaintPicture picnoseout.Picture, lblx6, lbly6, lblw6, lblh6, , , , , vbSrcAnd
    End If
    If piclipsout.Picture <> 0 Then
        picout.PaintPicture piclipsout.Picture, lblx7, lbly7, lblw7, lblh7, , , , , vbSrcAnd
    End If
    If picglassout.Picture <> 0 Then
        picout.PaintPicture picglassout.Picture, lblx8, lbly8, lblw8, lblh8, , , , , vbSrcAnd
    End If
    If picbeardout.Picture <> 0 Then
        picout.PaintPicture picbeardout.Picture, lblx9, lbly9, lblw9, lblh9, , , , , vbSrcAnd
    End If
    If piccapout.Picture <> 0 Then
        picout.PaintPicture piccapout.Picture, lblx10, lbly10, lblw10, lblh10, , , , , vbSrcAnd
    End If
    
End Sub

Private Sub newsketch()
    clearall
    
    With Toolbar2.Buttons
        .Item(1).Enabled = True
        .Item(2).Enabled = True
        .Item(3).Enabled = True
        .Item(4).Enabled = True
        .Item(5).Enabled = True
        .Item(6).Enabled = True
        .Item(7).Enabled = True
        .Item(8).Enabled = True
        .Item(9).Enabled = True
        .Item(10).Enabled = True
        .Item(1).Value = tbrUnpressed
        .Item(2).Value = tbrUnpressed
        .Item(3).Value = tbrUnpressed
        .Item(4).Value = tbrUnpressed
        .Item(5).Value = tbrUnpressed
        .Item(6).Value = tbrUnpressed
        .Item(7).Value = tbrUnpressed
        .Item(8).Value = tbrUnpressed
        .Item(9).Value = tbrUnpressed
        .Item(10).Value = tbrUnpressed
    End With
    
End Sub

Private Sub opensketch()
    
    File = Open_File(Me.hWnd) 'show the open file dlg
    If Trim(File) = "" Then Exit Sub ' make sure the file is correct
    'picout.Picture = LoadPicture(File) ' load the file
    clearall
    tmp.Caption = File
    
    Set mainRS = New ADODB.Recordset
    mainStr = "select * from cartographic where filename='" & tmp.Caption & "'"
    mainRS.Open mainStr, mainConn, adOpenKeyset, adLockReadOnly
    With mainRS
        picout.Cls
        If !jaw <> Empty Then
            lbljawout.Caption = !jaw
            lblx1.Caption = !x1
            lbly1.Caption = !y1
            lblw1.Caption = !w1
            lblh1.Caption = !h1
        End If
        If !hair <> Empty Then
            lblhairout.Caption = !hair
            lblx2.Caption = !x2
            lbly2.Caption = !y2
            lblw2.Caption = !w2
            lblh2.Caption = !h2
        End If
        If !ears <> Empty Then
            lblearsout.Caption = !ears
            lblx3.Caption = !x3
            lbly3.Caption = !y3
            lblw3.Caption = !w3
            lblh3.Caption = !h3
        End If
        If !eyebrow <> Empty Then
            lblbrowout.Caption = !eyebrow
            lblx4.Caption = !x4
            lbly4.Caption = !y4
            lblw4.Caption = !w4
            lblh4.Caption = !h4
        End If
        If !eyes <> Empty Then
            lbleyesout.Caption = !eyes
            lblx5.Caption = !x5
            lbly5.Caption = !y5
            lblw5.Caption = !w5
            lblh5.Caption = !h5
        End If
        If !nose <> Empty Then
            lblnoseout.Caption = !nose
            lblx6.Caption = !x6
            lbly6.Caption = !y6
            lblw6.Caption = !w6
            lblh6.Caption = !h6
        End If
        If !lips <> Empty Then
            lbllipsout.Caption = !lips
            lblx7.Caption = !x7
            lbly7.Caption = !y7
            lblw7.Caption = !w7
            lblh7.Caption = !h7
        End If
        If !glass <> Empty Then
            lblglassout.Caption = !glass
            lblx8.Caption = !x8
            lbly8.Caption = !y8
            lblw8.Caption = !w8
            lblh8.Caption = !h8
        End If
        If !beard <> Empty Then
            lblbeardout.Caption = !beard
            lblx9.Caption = !x9
            lbly9.Caption = !y9
            lblw9.Caption = !w9
            lblh9.Caption = !h9
        End If
        If !Cap <> Empty Then
            lblcapout.Caption = !Cap
            lblx10.Caption = !x10
            lbly10.Caption = !y10
            lblw10.Caption = !w10
            lblh10.Caption = !h10
        End If
        
        If lbljawout.Caption <> Empty Then 'jaw
            picjawout.Picture = LoadPicture(App.Path & lbljawout)
            picout.PaintPicture picjawout.Picture, lblx1, lbly1, lblw1, lblh1, , , , , vbSrcAnd
        End If
        If lblhairout.Caption <> Empty Then 'hair
            pichairout.Picture = LoadPicture(App.Path & lblhairout)
            picout.PaintPicture pichairout.Picture, lblx2, lbly2, lblw2, lblh2, , , , , vbSrcAnd
        End If
        If lblearsout.Caption <> Empty Then 'ears
            picearsout.Picture = LoadPicture(App.Path & lblearsout)
            picout.PaintPicture picearsout.Picture, lblx3, lbly3, lblw3, lblh3, , , , , vbSrcAnd
        End If
        If lblbrowout.Caption <> Empty Then 'eyebrow
            picbrowout.Picture = LoadPicture(App.Path & lblbrowout)
            picout.PaintPicture picbrowout.Picture, lblx4, lbly4, lblw4, lblh4, , , , , vbSrcAnd
        End If
        If lbleyesout.Caption <> Empty Then 'eyes
            piceyesout.Picture = LoadPicture(App.Path & lbleyesout)
            picout.PaintPicture piceyesout.Picture, lblx5, lbly5, lblw5, lblh5, , , , , vbSrcAnd
        End If
        If lblnoseout.Caption <> Empty Then 'nose
            picnoseout.Picture = LoadPicture(App.Path & lblnoseout)
            picout.PaintPicture picnoseout.Picture, lblx6, lbly6, lblw6, lblh6, , , , , vbSrcAnd
        End If
        If lbllipsout.Caption <> Empty Then 'lips
            piclipsout.Picture = LoadPicture(App.Path & lbllipsout)
            picout.PaintPicture piclipsout.Picture, lblx7, lbly7, lblw7, lblh7, , , , , vbSrcAnd
        End If
        If lblglassout.Caption <> Empty Then 'glass
            picglassout.Picture = LoadPicture(App.Path & lblglassout)
            picout.PaintPicture picglassout.Picture, lblx8, lbly8, lblw8, lblh8, , , , , vbSrcAnd
        End If
        If lblbeardout.Caption <> Empty Then 'beard
            picbeardout.Picture = LoadPicture(App.Path & lblbeardout)
            picout.PaintPicture picbeardout.Picture, lblx9, lbly9, lblw9, lblh9, , , , , vbSrcAnd
        End If
        If lblcapout.Caption <> Empty Then 'cap
            piccapout.Picture = LoadPicture(App.Path & lblcapout)
            picout.PaintPicture piccapout.Picture, lblx10, lbly10, lblw10, lblh10, , , , , vbSrcAnd
        End If
    End With
    
    
    
    If xusername = mainRS!sketchby Then
        Toolbar1.Buttons.Item(3).Enabled = True
        mnufilesave.Enabled = True
        mnufilesaveas.Enabled = True
     
        With Toolbar2.Buttons
            .Item(1).Enabled = True
            .Item(2).Enabled = True
            .Item(3).Enabled = True
            .Item(4).Enabled = True
            .Item(5).Enabled = True
            .Item(6).Enabled = True
            .Item(7).Enabled = True
            .Item(8).Enabled = True
            .Item(9).Enabled = True
            .Item(10).Enabled = True
        End With
        
    Else
        If xuserlevel = "Administrator" Then
            Toolbar1.Buttons.Item(3).Enabled = True
            mnufilesave.Enabled = True
            mnufilesaveas.Enabled = True
     
        With Toolbar2.Buttons
            .Item(1).Enabled = True
            .Item(2).Enabled = True
            .Item(3).Enabled = True
            .Item(4).Enabled = True
            .Item(5).Enabled = True
            .Item(6).Enabled = True
            .Item(7).Enabled = True
            .Item(8).Enabled = True
            .Item(9).Enabled = True
            .Item(10).Enabled = True
        End With
        
            Else
                MsgBox "You can't modify this skecth!", vbInformation, "Sketching"
                
                Toolbar1.Buttons.Item(3).Enabled = False
                mnufilesave.Enabled = False
                mnufilesaveas.Enabled = False
                
                mnueditremovejaw.Enabled = False
                mnueditremovehair.Enabled = False
                mnueditremovebrow.Enabled = False
                mnueditremoveeyes.Enabled = False
                mnueditremovenose.Enabled = False
                mnueditremoveears.Enabled = False
                mnueditremovebeard.Enabled = False
                mnueditremovecap.Enabled = False
                mnueditremoveglass.Enabled = False
                mnueditremovelips.Enabled = False
                
                mnuformatalignjaw.Enabled = False
                mnuformatalignhair.Enabled = False
                mnuformatalignbrow.Enabled = False
                mnuformataligneyes.Enabled = False
                mnuformatalignnose.Enabled = False
                mnuformatalignears.Enabled = False
                mnuformatalignbeard.Enabled = False
                mnuformataligncap.Enabled = False
                mnuformatalignglass.Enabled = False
                mnuformatalignlips.Enabled = False
                
                mnuformatsizejaw.Enabled = False
                mnuformatsizehair.Enabled = False
                mnuformatsizebrow.Enabled = False
                mnuformatsizeeyes.Enabled = False
                mnuformatsizenose.Enabled = False
                mnuformatsizeears.Enabled = False
                mnuformatsizebeard.Enabled = False
                mnuformatsizecap.Enabled = False
                mnuformatsizeglass.Enabled = False
                mnuformatsizelips.Enabled = False
                
                With Toolbar2.Buttons
                    .Item(1).Enabled = False
                    .Item(2).Enabled = False
                    .Item(3).Enabled = False
                    .Item(4).Enabled = False
                    .Item(5).Enabled = False
                    .Item(6).Enabled = False
                    .Item(7).Enabled = False
                    .Item(8).Enabled = False
                    .Item(9).Enabled = False
                    .Item(10).Enabled = False
                End With
            End If
        End If
    
    
End Sub

Private Sub saveassketch()
    File = Save_File(Me.hWnd) 'show save dlg
    If Trim(File) = "" Then Exit Sub ' error in name
    Saving = True ' start saving
    tmp.Caption = File '-- get rid of any unwanted chars (ie chr13, or 0)
    File = tmp.Caption '/
    If LCase(Right(File, 4) <> ".jpg") Then File = File & ".jpg"    ' add the jpg on the file
    tmp.Caption = File
    SavePicture picout.Image, File ' save the picture
    Saving = False ' no longer saving
    
    Set userRS = New ADODB.Recordset
    userStr = "Select username from users where status=" & 1
    userRS.Open userStr, userConn, adOpenKeyset, adLockReadOnly
    xsketchby = userRS!Username
    
    Set mainRS = New ADODB.Recordset
    mainStr = "select * from cartographic where filename='" & tmp.Caption & " '"
    mainRS.Open mainStr, mainConn, adOpenKeyset, adLockOptimistic
    If Not mainRS.EOF And Not mainRS.BOF Then
        With mainRS
        If lbljawout.Caption <> Empty Then
            !jaw = lbljawout.Caption
            !x1 = lblx1.Caption
            !y1 = lbly1.Caption
            !w1 = lblw1.Caption
            !h1 = lblh1.Caption
        Else
            !jaw = ""
            !x1 = 0
            !y1 = 0
            !w1 = 0
            !h1 = 0
        End If
        If lblhairout.Caption <> Empty Then
            !hair = lblhairout.Caption
            !x2 = lblx2.Caption
            !y2 = lbly2.Caption
            !w2 = lblw2.Caption
            !h2 = lblh2.Caption
        Else
            !hair = ""
            !x2 = 0
            !y2 = 0
            !w2 = 0
            !h2 = 0
        End If
        If lblearsout.Caption <> Empty Then
            !ears = lblearsout.Caption
            !x3 = lblx3.Caption
            !y3 = lbly3.Caption
            !w3 = lblw3.Caption
            !h3 = lblh3.Caption
        Else
            !ears = ""
            !x3 = 0
            !y3 = 0
            !w3 = 0
            !h3 = 0
        End If
        If lblbrowout.Caption <> Empty Then
            !eyebrow = lblbrowout.Caption
            !x4 = lblx4.Caption
            !y4 = lbly4.Caption
            !w4 = lblw4.Caption
            !h4 = lblh4.Caption
        Else
            !eyebrow = ""
            !x4 = 0
            !y4 = 0
            !w4 = 0
            !h4 = 0
        End If
        If lbleyesout.Caption <> Empty Then
            !eyes = lbleyesout.Caption
            !x5 = lblx5.Caption
            !y5 = lbly5.Caption
            !w5 = lblw5.Caption
            !h5 = lblh5.Caption
        Else
            !eyes = ""
            !x5 = 0
            !y5 = 0
            !w5 = 0
            !h5 = 0
        End If
        If lblnoseout.Caption <> Empty Then
            !nose = lblnoseout.Caption
            !x6 = lblx6.Caption
            !y6 = lbly6.Caption
            !w6 = lblw6.Caption
            !h6 = lblh6.Caption
        Else
            !nose = ""
            !x6 = 0
            !y6 = 0
            !w6 = 0
            !h6 = 0
        End If
        If lbllipsout.Caption <> Empty Then
            !lips = lbllipsout.Caption
            !x7 = lblx7.Caption
            !y7 = lbly7.Caption
            !w7 = lblw7.Caption
            !h7 = lblh7.Caption
        Else
            !lips = ""
            !x7 = 0
            !y7 = 0
            !w7 = 0
            !h7 = 0
        End If
        If lblglassout.Caption <> Empty Then
            !glass = lblglassout.Caption
            !x8 = lblx8.Caption
            !y8 = lbly8.Caption
            !w8 = lblw8.Caption
            !h8 = lblh8.Caption
        Else
            !glass = ""
            !x8 = 0
            !y8 = 0
            !w8 = 0
            !h8 = 0
        End If
        If lblbeardout.Caption <> Empty Then
            !beard = lblbeardout.Caption
            !x9 = lblx9.Caption
            !y9 = lbly9.Caption
            !w9 = lblw9.Caption
            !h9 = lblh9.Caption
        Else
            !beard = ""
            !x9 = 0
            !y9 = 0
            !w9 = 0
            !h9 = 0
        End If
        If lblcapout.Caption <> Empty Then
            !Cap = lblcapout.Caption
            !x10 = lblx10.Caption
            !y10 = lbly10.Caption
            !w10 = lblw10.Caption
            !h10 = lblh10.Caption
        Else
            !Cap = ""
            !x10 = 0
            !y10 = 0
            !w10 = 0
            !h10 = 0
        End If
        
            .Update
            .Close
        End With
Else
    Set mainRS = New ADODB.Recordset
    mainRS.Open "cartographic", mainConn, adOpenKeyset, adLockOptimistic
    With mainRS
        .AddNew
        !FileName = tmp.Caption
        If lbljawout.Caption <> Empty Then
            !jaw = lbljawout.Caption
            !x1 = lblx1.Caption
            !y1 = lbly1.Caption
            !w1 = lblw1.Caption
            !h1 = lblh1.Caption
        End If
        If lblhairout.Caption <> Empty Then
            !hair = lblhairout.Caption
            !x2 = lblx2.Caption
            !y2 = lbly2.Caption
            !w2 = lblw2.Caption
            !h2 = lblh2.Caption
        End If
        If lblearsout.Caption <> Empty Then
            !ears = lblearsout.Caption
            !x3 = lblx3.Caption
            !y3 = lbly3.Caption
            !w3 = lblw3.Caption
            !h3 = lblh3.Caption
        End If
        If lblbrowout.Caption <> Empty Then
            !eyebrow = lblbrowout.Caption
            !x4 = lblx4.Caption
            !y4 = lbly4.Caption
            !w4 = lblw4.Caption
            !h4 = lblh4.Caption
        End If
        If lbleyesout.Caption <> Empty Then
            !eyes = lbleyesout.Caption
            !x5 = lblx5.Caption
            !y5 = lbly5.Caption
            !w5 = lblw5.Caption
            !h5 = lblh5.Caption
        End If
        If lblnoseout.Caption <> Empty Then
            !nose = lblnoseout.Caption
            !x6 = lblx6.Caption
            !y6 = lbly6.Caption
            !w6 = lblw6.Caption
            !h6 = lblh6.Caption
        End If
        If lbllipsout.Caption <> Empty Then
            !lips = lbllipsout.Caption
            !x7 = lblx7.Caption
            !y7 = lbly7.Caption
            !w7 = lblw7.Caption
            !h7 = lblh7.Caption
        End If
        If lblglassout.Caption <> Empty Then
            !glass = lblglassout.Caption
            !x8 = lblx8.Caption
            !y8 = lbly8.Caption
            !w8 = lblw8.Caption
            !h8 = lblh8.Caption
        End If
        If lblbeardout.Caption <> Empty Then
            !beard = lblbeardout.Caption
            !x9 = lblx9.Caption
            !y9 = lbly9.Caption
            !w9 = lblw9.Caption
            !h9 = lblh9.Caption
        End If
        If lblcapout.Caption <> Empty Then
            !Cap = lblcapout.Caption
            !x10 = lblx10.Caption
            !y10 = lbly10.Caption
            !w10 = lblw10.Caption
            !h10 = lblh10.Caption
        End If
            !sketchby = xsketchby
        .Update
        .Close
    End With
End If
End Sub

Private Sub savesketch()
    
        Saving = True ' start saving
        SavePicture picout.Image, File ' save the picture
        Saving = False ' no longer saving
        
        Set mainRS = New ADODB.Recordset
        mainStr = "select * from cartographic where filename='" & tmp.Caption & " '"
        mainRS.Open mainStr, mainConn, adOpenKeyset, adLockOptimistic
        With mainRS
        If lbljawout.Caption <> Empty Then
            !jaw = lbljawout.Caption
            !x1 = lblx1.Caption
            !y1 = lbly1.Caption
            !w1 = lblw1.Caption
            !h1 = lblh1.Caption
        Else
            !jaw = ""
            !x1 = 0
            !y1 = 0
            !w1 = 0
            !h1 = 0
        End If
        If lblhairout.Caption <> Empty Then
            !hair = lblhairout.Caption
            !x2 = lblx2.Caption
            !y2 = lbly2.Caption
            !w2 = lblw2.Caption
            !h2 = lblh2.Caption
        Else
            !hair = ""
            !x2 = 0
            !y2 = 0
            !w2 = 0
            !h2 = 0
        End If
        If lblearsout.Caption <> Empty Then
            !ears = lblearsout.Caption
            !x3 = lblx3.Caption
            !y3 = lbly3.Caption
            !w3 = lblw3.Caption
            !h3 = lblh3.Caption
        Else
            !ears = ""
            !x3 = 0
            !y3 = 0
            !w3 = 0
            !h3 = 0
        End If
        If lblbrowout.Caption <> Empty Then
            !eyebrow = lblbrowout.Caption
            !x4 = lblx4.Caption
            !y4 = lbly4.Caption
            !w4 = lblw4.Caption
            !h4 = lblh4.Caption
        Else
            !eyebrow = ""
            !x4 = 0
            !y4 = 0
            !w4 = 0
            !h4 = 0
        End If
        If lbleyesout.Caption <> Empty Then
            !eyes = lbleyesout.Caption
            !x5 = lblx5.Caption
            !y5 = lbly5.Caption
            !w5 = lblw5.Caption
            !h5 = lblh5.Caption
        Else
            !eyes = ""
            !x5 = 0
            !y5 = 0
            !w5 = 0
            !h5 = 0
        End If
        If lblnoseout.Caption <> Empty Then
            !nose = lblnoseout.Caption
            !x6 = lblx6.Caption
            !y6 = lbly6.Caption
            !w6 = lblw6.Caption
            !h6 = lblh6.Caption
        Else
            !nose = ""
            !x6 = 0
            !y6 = 0
            !w6 = 0
            !h6 = 0
        End If
        If lbllipsout.Caption <> Empty Then
            !lips = lbllipsout.Caption
            !x7 = lblx7.Caption
            !y7 = lbly7.Caption
            !w7 = lblw7.Caption
            !h7 = lblh7.Caption
        Else
            !lips = ""
            !x7 = 0
            !y7 = 0
            !w7 = 0
            !h7 = 0
        End If
        If lblglassout.Caption <> Empty Then
            !glass = lblglassout.Caption
            !x8 = lblx8.Caption
            !y8 = lbly8.Caption
            !w8 = lblw8.Caption
            !h8 = lblh8.Caption
        Else
            !glass = ""
            !x8 = 0
            !y8 = 0
            !w8 = 0
            !h8 = 0
        End If
        If lblbeardout.Caption <> Empty Then
            !beard = lblbeardout.Caption
            !x9 = lblx9.Caption
            !y9 = lbly9.Caption
            !w9 = lblw9.Caption
            !h9 = lblh9.Caption
        Else
            !beard = ""
            !x9 = 0
            !y9 = 0
            !w9 = 0
            !h9 = 0
        End If
        If lblcapout.Caption <> Empty Then
            !Cap = lblcapout.Caption
            !x10 = lblx10.Caption
            !y10 = lbly10.Caption
            !w10 = lblw10.Caption
            !h10 = lblh10.Caption
        Else
            !Cap = ""
            !x10 = 0
            !y10 = 0
            !w10 = 0
            !h10 = 0
        End If
        
            .Update
            .Close
        End With
        
End Sub

Private Sub printsketch()
    Set mainRS = New ADODB.Recordset
    mainStr = "select * from cartographic where filename='" & tmp.Caption & "'"
    mainRS.Open mainStr, mainConn, adOpenKeyset, adLockOptimistic
    Set rptsketch.DataSource = mainRS
    Set rptsketch.Sections(3).Controls("Image1").Picture = LoadPicture(mainRS!FileName)
    rptsketch.Show vbModal
End Sub

Private Sub jawsize() 'jaw alignment & sizes
    With jawRS
        x1 = !x
        y1 = !y
        w1 = !Width
        h1 = !Height
    End With
End Sub

Private Sub hairsize() 'hair alignment & sizes
    With hairRS
        x2 = !x
        y2 = !y
        w2 = !Width
        h2 = !Height
    End With
End Sub

Private Sub earssize() 'ears alignment & sizes
    With earsRS
        x3 = !x
        y3 = !y
        w3 = !Width
        h3 = !Height
    End With
End Sub

Private Sub browsize() 'eyebrow alignment & sizes
    With browRS
        x4 = !x
        y4 = !y
        w4 = !Width
        h4 = !Height
    End With
End Sub

Private Sub eyessize() 'eyes alignment & sizes
    With eyesRS
        x5 = !x
        y5 = !y
        w5 = !Width
        h5 = !Height
    End With
End Sub

Private Sub nosesize() 'nose alignment & sizes
    With noseRS
        x6 = !x
        y6 = !y
        w6 = !Width
        h6 = !Height
    End With
End Sub

Private Sub lipssize() 'lips alignment & sizes
    With lipsRS
        x7 = !x
        y7 = !y
        w7 = !Width
        h7 = !Height
    End With
End Sub

Private Sub glasssize() 'glass alignment & sizes
    With glassRS
        x8 = !x
        y8 = !y
        w8 = !Width
        h8 = !Height
    End With
End Sub

Private Sub beardsize() 'beard alignment & sizes
    With beardRS
        x9 = !x
        y9 = !y
        w9 = !Width
        h9 = !Height
    End With
End Sub

Private Sub capsize() 'glass alignment & sizes
    With capRS
        x10 = !x
        y10 = !y
        w10 = !Width
        h10 = !Height
    End With
End Sub

Private Sub cmdenabledisable() 'command button disabled/enabled
    cmdnext.Enabled = True
    cmdprevious.Enabled = False
    cmdselect.Enabled = True
End Sub

Private Sub cmdnext_Click()
With Toolbar2.Buttons
    If .Item(1).Value = tbrPressed Then 'jaw
        If jawRS.EOF <> True Then
            cmdprevious.Enabled = True
            jawRS.MoveNext
            On Error GoTo cmderror1
            picview.Refresh
            lbljaw.Caption = jawRS!pic
            picjaw.Picture = LoadPicture(App.Path & jawRS!pic)
            jawsize
            picview.PaintPicture picjaw.Picture, x1, y1, w1, h1, , , , , vbSrcAnd
            picjaw.Refresh
        Else
cmderror1:
    picjaw.Picture = Nothing
    cmdnext.Enabled = False
    cmdprevious.Enabled = True
        End If
    End If
    
    If .Item(2).Value = tbrPressed Then 'hair
        If hairRS.EOF <> True Then
            cmdprevious.Enabled = True
            hairRS.MoveNext
            On Error GoTo cmderror2
            picview.Refresh
            lblhair.Caption = hairRS!pic
            pichair.Picture = LoadPicture(App.Path & hairRS!pic)
            hairsize
            picview.PaintPicture pichair.Picture, x2, y2, w2, h2, , , , , vbSrcAnd
            pichair.Refresh
        Else
cmderror2:
        pichair.Picture = Nothing
        cmdnext.Enabled = False
        cmdprevious.Enabled = True
        End If
    End If
    
    If .Item(3).Value = tbrPressed Then 'ears
        If earsRS.EOF <> True Then
            cmdprevious.Enabled = True
            earsRS.MoveNext
            On Error GoTo cmderror3
            picview.Refresh
            lblears.Caption = earsRS!pic
            picears.Picture = LoadPicture(App.Path & earsRS!pic)
            earssize
            picview.PaintPicture picears.Picture, x3, y3, w3, h3, , , , , vbSrcAnd
            picears.Refresh
        Else
cmderror3:
        picears.Picture = Nothing
        cmdnext.Enabled = False
        cmdprevious.Enabled = True
        End If
    End If
    
    If .Item(4).Value = tbrPressed Then 'eyebrow
        If browRS.EOF <> True Then
            cmdprevious.Enabled = True
            browRS.MoveNext
            On Error GoTo cmderror4
            picview.Refresh
            lblbrow.Caption = browRS!pic
            picbrow.Picture = LoadPicture(App.Path & browRS!pic)
            browsize
            picview.PaintPicture picbrow.Picture, x4, y4, w4, h4, , , , , vbSrcAnd
            picbrow.Refresh
        Else
cmderror4:
        picbrow.Picture = Nothing
        cmdnext.Enabled = False
        cmdprevious.Enabled = True
        End If
    End If
    
    If .Item(5).Value = tbrPressed Then 'eyes
        If eyesRS.EOF <> True Then
            cmdprevious.Enabled = True
            eyesRS.MoveNext
            On Error GoTo cmderror5
            picview.Refresh
            lbleyes.Caption = eyesRS!pic
            piceyes.Picture = LoadPicture(App.Path & eyesRS!pic)
            eyessize
            picview.PaintPicture piceyes.Picture, x5, y5, w5, h5, , , , , vbSrcAnd
            piceyes.Refresh
        Else
cmderror5:
        piceyes.Picture = Nothing
        cmdnext.Enabled = False
        cmdprevious.Enabled = True
        End If
    End If
    
    If .Item(6).Value = tbrPressed Then 'nose
        If noseRS.EOF <> True Then
            cmdprevious.Enabled = True
            noseRS.MoveNext
            On Error GoTo cmderror6
            picview.Refresh
            lblnose.Caption = noseRS!pic
            picnose.Picture = LoadPicture(App.Path & noseRS!pic)
            nosesize
            picview.PaintPicture picnose.Picture, x6, y6, w6, h6, , , , , vbSrcAnd
            picnose.Refresh
        Else
cmderror6:
        picnose.Picture = Nothing
        cmdnext.Enabled = False
        cmdprevious.Enabled = True
        End If
    End If
    
    If .Item(7).Value = tbrPressed Then 'lips
        If lipsRS.EOF <> True Then
            cmdprevious.Enabled = True
            lipsRS.MoveNext
            On Error GoTo cmderror7
            picview.Refresh
            lbllips.Caption = lipsRS!pic
            piclips.Picture = LoadPicture(App.Path & lipsRS!pic)
            lipssize
            picview.PaintPicture piclips.Picture, x7, y7, w7, h7, , , , , vbSrcAnd
            piclips.Refresh
        Else
cmderror7:
        piclips.Picture = Nothing
        cmdnext.Enabled = False
        cmdprevious.Enabled = True
        End If
    End If
    
       If .Item(8).Value = tbrPressed Then 'glass
        If glassRS.EOF <> True Then
            cmdprevious.Enabled = True
            glassRS.MoveNext
            On Error GoTo cmderror8
            picview.Refresh
            lblglass.Caption = glassRS!pic
            picglass.Picture = LoadPicture(App.Path & glassRS!pic)
            glasssize
            picview.PaintPicture picglass.Picture, x8, y8, w8, h8, , , , , vbSrcAnd
            picglass.Refresh
        Else
cmderror8:
        picglass.Picture = Nothing
        cmdnext.Enabled = False
        cmdprevious.Enabled = True
        End If
    End If
    
    If .Item(9).Value = tbrPressed Then 'beard
        If beardRS.EOF <> True Then
            cmdprevious.Enabled = True
            beardRS.MoveNext
            On Error GoTo cmderror9
            picview.Refresh
            lblbeard.Caption = beardRS!pic
            picbeard.Picture = LoadPicture(App.Path & beardRS!pic)
            beardsize
            picview.PaintPicture picbeard.Picture, x9, y9, w9, h9, , , , , vbSrcAnd
            picbeard.Refresh
        Else
cmderror9:
        picbeard.Picture = Nothing
        cmdnext.Enabled = False
        cmdprevious.Enabled = True
        End If
    End If
    
    If .Item(10).Value = tbrPressed Then 'cap
        If capRS.EOF <> True Then
            cmdprevious.Enabled = True
            capRS.MoveNext
            On Error GoTo cmderror10
            picview.Refresh
            lblcap.Caption = capRS!pic
            piccap.Picture = LoadPicture(App.Path & capRS!pic)
            capsize
            picview.PaintPicture piccap.Picture, x10, y10, w10, h10, , , , , vbSrcAnd
            piccap.Refresh
        Else
cmderror10:
        piccap.Picture = Nothing
        cmdnext.Enabled = False
        cmdprevious.Enabled = True
        End If
    End If

End With

    Exit Sub


End Sub

Private Sub cmdprevious_Click()
With Toolbar2.Buttons
    If .Item(1).Value = tbrPressed Then 'jaw
        If jawRS.BOF <> True Then
        cmdnext.Enabled = True
        jawRS.MovePrevious
        On Error GoTo cmderror1
        picview.Refresh
        lbljaw.Caption = jawRS!pic
        picjaw.Picture = LoadPicture(App.Path & jawRS!pic)
        jawsize
        picview.PaintPicture picjaw.Picture, x1, y1, w1, h1, , , , , vbSrcAnd
        picjaw.Refresh
        Else
cmderror1:
        picjaw.Picture = Nothing
        cmdprevious.Enabled = False
        cmdnext.Enabled = True
        End If
    End If
    
    If .Item(2).Value = tbrPressed Then 'hair
        If hairRS.BOF <> True Then
            cmdnext.Enabled = True
            hairRS.MovePrevious
            On Error GoTo cmderror2
            picview.Refresh
            lblhair.Caption = hairRS!pic
            pichair.Picture = LoadPicture(App.Path & hairRS!pic)
            hairsize
            picview.PaintPicture pichair.Picture, x2, y2, w2, h2, , , , , vbSrcAnd
            pichair.Refresh
        Else
cmderror2:
        pichair.Picture = Nothing
        cmdprevious.Enabled = False
        cmdnext.Enabled = True
        End If
    End If
    
    If .Item(3).Value = tbrPressed Then 'ears
        If earsRS.BOF <> True Then
            cmdnext.Enabled = True
            earsRS.MovePrevious
            On Error GoTo cmderror3
            picview.Refresh
            lblears.Caption = earsRS!pic
            picears.Picture = LoadPicture(App.Path & earsRS!pic)
            earssize
            picview.PaintPicture picears.Picture, x3, y3, w3, h3, , , , , vbSrcAnd
            picears.Refresh
        Else
cmderror3:
        picears.Picture = Nothing
        cmdprevious.Enabled = False
        cmdnext.Enabled = True
        End If
    End If
    
    If .Item(4).Value = tbrPressed Then 'eyebrow
        If browRS.BOF <> True Then
            cmdnext.Enabled = True
            browRS.MovePrevious
            On Error GoTo cmderror4
            picview.Refresh
            lblbrow.Caption = browRS!pic
            picbrow.Picture = LoadPicture(App.Path & browRS!pic)
            browsize
            picview.PaintPicture picbrow.Picture, x4, y4, w4, h4, , , , , vbSrcAnd
            picbrow.Refresh
        Else
cmderror4:
        picbrow.Picture = Nothing
        cmdprevious.Enabled = False
        cmdnext.Enabled = True
        End If
    End If
    
    If .Item(5).Value = tbrPressed Then 'eyes
        If eyesRS.BOF <> True Then
            cmdnext.Enabled = True
            eyesRS.MovePrevious
            On Error GoTo cmderror5
            picview.Refresh
            lbleyes.Caption = eyesRS!pic
            piceyes.Picture = LoadPicture(App.Path & eyesRS!pic)
            eyessize
            picview.PaintPicture piceyes.Picture, x5, y5, w5, h5, , , , , vbSrcAnd
            piceyes.Refresh
        Else
cmderror5:
        piceyes.Picture = Nothing
        cmdprevious.Enabled = False
        cmdnext.Enabled = True
        End If
    End If
    
    If .Item(6).Value = tbrPressed Then 'nose
        If noseRS.BOF <> True Then
            cmdnext.Enabled = True
            noseRS.MovePrevious
            On Error GoTo cmderror6
            picview.Refresh
            lblnose.Caption = noseRS!pic
            picnose.Picture = LoadPicture(App.Path & noseRS!pic)
            nosesize
            picview.PaintPicture picnose.Picture, x6, y6, w6, h6, , , , , vbSrcAnd
            picnose.Refresh
        Else
cmderror6:
        picnose.Picture = Nothing
        cmdprevious.Enabled = False
        cmdnext.Enabled = True
        End If
    End If
    
    If .Item(7).Value = tbrPressed Then 'lips
        If lipsRS.BOF <> True Then
            cmdnext.Enabled = True
            lipsRS.MovePrevious
            On Error GoTo cmderror7
            picview.Refresh
            lbllips.Caption = lipsRS!pic
            piclips.Picture = LoadPicture(App.Path & lipsRS!pic)
            lipssize
            picview.PaintPicture piclips.Picture, x7, y7, w7, h7, , , , , vbSrcAnd
            piclips.Refresh
        Else
cmderror7:
        piclips.Picture = Nothing
        cmdprevious.Enabled = False
        cmdnext.Enabled = True
        End If
    End If
    
    If .Item(8).Value = tbrPressed Then 'glass
        If glassRS.BOF <> True Then
            cmdnext.Enabled = True
            glassRS.MovePrevious
            On Error GoTo cmderror8
            picview.Refresh
            lblglass.Caption = glassRS!pic
            picglass.Picture = LoadPicture(App.Path & glassRS!pic)
            glasssize
            picview.PaintPicture picglass.Picture, x8, y8, w8, h8, , , , , vbSrcAnd
            picglass.Refresh
        Else
cmderror8:
        picglass.Picture = Nothing
        cmdprevious.Enabled = False
        cmdnext.Enabled = True
        End If
    End If
    
    If .Item(9).Value = tbrPressed Then 'beard
        If beardRS.BOF <> True Then
            cmdnext.Enabled = True
            beardRS.MovePrevious
            On Error GoTo cmderror9
            picview.Refresh
            lblbeard.Caption = beardRS!pic
            picbeard.Picture = LoadPicture(App.Path & beardRS!pic)
            beardsize
            picview.PaintPicture picbeard.Picture, x9, y9, w9, h9, , , , , vbSrcAnd
            picbeard.Refresh
        Else
cmderror9:
        picbeard.Picture = Nothing
        cmdprevious.Enabled = False
        cmdnext.Enabled = True
        End If
    End If
    
    If .Item(10).Value = tbrPressed Then 'cap
        If capRS.BOF <> True Then
            cmdnext.Enabled = True
            capRS.MovePrevious
            On Error GoTo cmderror10
            picview.Refresh
            lblcap.Caption = capRS!pic
            piccap.Picture = LoadPicture(App.Path & capRS!pic)
            capsize
            picview.PaintPicture piccap.Picture, x10, y10, w10, h10, , , , , vbSrcAnd
            piccap.Refresh
        Else
cmderror10:
        piccap.Picture = Nothing
        cmdprevious.Enabled = False
        cmdnext.Enabled = True
        End If
    End If
    
End With

End Sub

Private Sub cmdselect_Click()
          
    Toolbar1.Buttons.Item(3).Enabled = True
    mnufilesaveas.Enabled = True
    mnufilesave.Enabled = True
        
    picout.Cls
    If picjaw.Picture <> 0 Then 'jaw
        lblx1.Caption = x1
        lbly1.Caption = y1
        lblw1.Caption = w1
        lblh1.Caption = h1
        'mnuformatalignjaw.Enabled = True
        'mnuformatsizejaw.Enabled = True
        lbljawout.Caption = lbljaw.Caption
        picjawout.Picture = picjaw.Picture
        picout.PaintPicture picjawout.Picture, lblx1, lbly1, lblw1, lblh1, , , , , vbSrcAnd
    ElseIf picjawout.Picture <> 0 Then
        picout.PaintPicture picjawout.Picture, lblx1, lbly1, lblw1, lblh1, , , , , vbSrcAnd
    End If
    
    If pichair.Picture <> 0 Then 'hair
        lblx2.Caption = x2
        lbly2.Caption = y2
        lblw2.Caption = w2
        lblh2.Caption = h2
        'mnuformatalignhair.Enabled = True
        'mnuformatsizehair.Enabled = True
        lblhairout.Caption = lblhair.Caption
        pichairout.Picture = pichair.Picture
        picout.PaintPicture pichairout.Picture, lblx2, lbly2, lblw2, lblh2, , , , , vbSrcAnd
    ElseIf pichairout.Picture <> 0 Then
        picout.PaintPicture pichairout.Picture, lblx2, lbly2, lblw2, lblh2, , , , , vbSrcAnd
    End If
    
    If picears.Picture <> 0 Then 'ears
        lblx3.Caption = x3
        lbly3.Caption = y3
        lblw3.Caption = w3
        lblh3.Caption = h3
        'mnuformatalignears.Enabled = True
        'mnuformatsizeears.Enabled = True
        lblearsout.Caption = lblears.Caption
        picearsout.Picture = picears.Picture
        picout.PaintPicture picearsout.Picture, lblx3, lbly3, lblw3, lblh3, , , , , vbSrcAnd
    ElseIf picearsout.Picture <> 0 Then
        picout.PaintPicture picearsout.Picture, lblx3, lbly3, lblw3, lblh3, , , , , vbSrcAnd
    End If
    
    If picbrow.Picture <> 0 Then 'eyebrow
        lblx4.Caption = x4
        lbly4.Caption = y4
        lblw4.Caption = w4
        lblh4.Caption = h4
        'mnuformatalignbrow.Enabled = True
        'mnuformatsizebrow.Enabled = True
        lblbrowout.Caption = lblbrow.Caption
        picbrowout.Picture = picbrow.Picture
        picout.PaintPicture picbrowout.Picture, lblx4, lbly4, lblw4, lblh4, , , , , vbSrcAnd
    ElseIf picbrowout.Picture <> 0 Then
        picout.PaintPicture picbrowout.Picture, lblx4, lbly4, lblw4, lblh4, , , , , vbSrcAnd
    End If
    
    If piceyes.Picture <> 0 Then 'eyes
        lblx5.Caption = x5
        lbly5.Caption = y5
        lblw5.Caption = w5
        lblh5.Caption = h5
        'mnuformataligneyes.Enabled = True
        'mnuformatsizeeyes.Enabled = True
        lbleyesout.Caption = lbleyes.Caption
        piceyesout.Picture = piceyes.Picture
        picout.PaintPicture piceyesout.Picture, lblx5, lbly5, lblw5, lblh5, , , , , vbSrcAnd
    ElseIf piceyesout.Picture <> 0 Then
        picout.PaintPicture piceyesout.Picture, lblx5, lbly5, lblw5, lblh5, , , , , vbSrcAnd
    End If
    
    If picnose.Picture <> 0 Then 'nose
        lblx6.Caption = x6
        lbly6.Caption = y6
        lblw6.Caption = w6
        lblh6.Caption = h6
        'mnuformatalignnose.Enabled = True
        'mnuformatsizenose.Enabled = True
        lblnoseout.Caption = lblnose.Caption
        picnoseout.Picture = picnose.Picture
        picout.PaintPicture picnoseout.Picture, lblx6, lbly6, lblw6, lblh6, , , , , vbSrcAnd
    ElseIf picnoseout.Picture <> 0 Then
        picout.PaintPicture picnoseout.Picture, lblx6, lbly6, lblw6, lblh6, , , , , vbSrcAnd
    End If
    
    If piclips.Picture <> 0 Then 'lips
        lblx7.Caption = x7
        lbly7.Caption = y7
        lblw7.Caption = w7
        lblh7.Caption = h7
        'mnuformatalignlips.Enabled = True
        'mnuformatsizelips.Enabled = True
        lbllipsout.Caption = lbllips.Caption
        piclipsout.Picture = piclips.Picture
        picout.PaintPicture piclipsout.Picture, lblx7, lbly7, lblw7, lblh7, , , , , vbSrcAnd
    ElseIf piclipsout.Picture <> 0 Then
        picout.PaintPicture piclipsout.Picture, lblx7, lbly7, lblw7, lblh7, , , , , vbSrcAnd
    End If
    
    If picglass.Picture <> 0 Then 'glass
        lblx8.Caption = x8
        lbly8.Caption = y8
        lblw8.Caption = w8
        lblh8.Caption = h8
        'mnuformatalignglass.Enabled = True
        'mnuformatsizeglass.Enabled = True
        lblglassout.Caption = lblglass.Caption
        picglassout.Picture = picglass.Picture
        picout.PaintPicture picglassout.Picture, lblx8, lbly8, lblw8, lblh8, , , , , vbSrcAnd
    ElseIf picglassout.Picture <> 0 Then
        picout.PaintPicture picglassout.Picture, lblx8, lbly8, lblw8, lblh8, , , , , vbSrcAnd
    End If
    
    If picbeard.Picture <> 0 Then 'beard
        lblx9.Caption = x9
        lbly9.Caption = y9
        lblw9.Caption = w9
        lblh9.Caption = h9
        'mnuformatalignglass.Enabled = True
        'mnuformatsizeglass.Enabled = True
        lblbeardout.Caption = lblbeard.Caption
        picbeardout.Picture = picbeard.Picture
        picout.PaintPicture picbeardout.Picture, lblx9, lbly9, lblw9, lblh9, , , , , vbSrcAnd
    ElseIf picbeardout.Picture <> 0 Then
        picout.PaintPicture picbeardout.Picture, lblx9, lbly9, lblw9, lblh9, , , , , vbSrcAnd
    End If
    
    If piccap.Picture <> 0 Then 'cap
        lblx10.Caption = x10
        lbly10.Caption = y10
        lblw10.Caption = w10
        lblh10.Caption = h10
        'mnuformatalignglass.Enabled = True
        'mnuformatsizeglass.Enabled = True
        lblcapout.Caption = lblcap.Caption
        piccapout.Picture = piccap.Picture
        picout.PaintPicture piccapout.Picture, lblx10, lbly10, lblw10, lblh10, , , , , vbSrcAnd
    ElseIf piccapout.Picture <> 0 Then
        picout.PaintPicture piccapout.Picture, lblx10, lbly10, lblw10, lblh10, , , , , vbSrcAnd
    End If
    
End Sub

Private Sub Form_Load()

    partsdbConnect
    maindbConnect
    
    cmdnext.Enabled = False
    cmdprevious.Enabled = False
    cmdselect.Enabled = False
    
    With Toolbar2.Buttons
        .Item(1).Enabled = False
        .Item(2).Enabled = False
        .Item(3).Enabled = False
        .Item(4).Enabled = False
        .Item(5).Enabled = False
        .Item(6).Enabled = False
        .Item(7).Enabled = False
        .Item(8).Enabled = False
        .Item(9).Enabled = False
        .Item(10).Enabled = False
    End With
    
    With Toolbar1.Buttons
        .Item(3).Enabled = False
        .Item(4).Enabled = False
    End With
    
    mnufilesave.Enabled = False
    mnufilesaveas.Enabled = False
    mnufileprint.Enabled = False
    
    
    Set userRS = New ADODB.Recordset
    userStr = "Select username,userlevel from users where status=" & 1
    userRS.Open userStr, userConn, adOpenKeyset, adLockReadOnly
    xusername = userRS!Username
    xuserlevel = userRS!userlevel
    
    InitDlgs 'initalize save and open dialogs
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Close Sketching?", vbQuestion + vbYesNo, "Sketching") = vbYes Then
        Unload Me
    Else
        Cancel = True
    End If
End Sub

Private Sub mnueditremovebeard_Click()
    picbeardout.Picture = Nothing
    lblbeardout.Caption = Empty
    lblx9.Caption = Empty
    lbly9.Caption = Empty
    lblw9.Caption = Empty
    lblh9.Caption = Empty
    remove
End Sub

Private Sub mnueditremovebrow_Click()
    picbrowout.Picture = Nothing
    lblbrowout.Caption = Empty
    lblx4.Caption = Empty
    lbly4.Caption = Empty
    lblw4.Caption = Empty
    lblh4.Caption = Empty
    remove
End Sub

Private Sub mnueditremovecap_Click()
    piccapout.Picture = Nothing
    lblcapout.Caption = Empty
    lblx10.Caption = Empty
    lbly10.Caption = Empty
    lblw10.Caption = Empty
    lblh10.Caption = Empty
    remove
End Sub

Private Sub mnueditremoveears_Click()
    picearsout.Picture = Nothing
    lblearsout.Caption = Empty
    lblx3.Caption = Empty
    lbly3.Caption = Empty
    lblw3.Caption = Empty
    lblh3.Caption = Empty
    remove
End Sub

Private Sub mnueditremoveeyes_Click()
    piceyesout.Picture = Nothing
    lbleyesout.Caption = Empty
    lblx5.Caption = Empty
    lbly5.Caption = Empty
    lblw5.Caption = Empty
    lblh5.Caption = Empty
    remove
End Sub

Private Sub mnueditremoveglass_Click()
    picglassout.Picture = Nothing
    lblglassout.Caption = Empty
    lblx8.Caption = Empty
    lbly8.Caption = Empty
    lblw8.Caption = Empty
    lblh8.Caption = Empty
    remove
End Sub

Private Sub mnueditremovehair_Click()
    pichairout.Picture = Nothing
    lblhairout.Caption = Empty
    lblx2.Caption = Empty
    lbly2.Caption = Empty
    lblw2.Caption = Empty
    lblh2.Caption = Empty
    remove
End Sub

Private Sub mnueditremovejaw_Click()
    picjawout.Picture = Nothing
    lbljawout.Caption = Empty
    lblx1.Caption = Empty
    lbly1.Caption = Empty
    lblw1.Caption = Empty
    lblh1.Caption = Empty
    remove
End Sub

Private Sub mnueditremovelips_Click()
    piclipsout.Picture = Nothing
    lbllipsout.Caption = Empty
    lblx7.Caption = Empty
    lbly7.Caption = Empty
    lblw7.Caption = Empty
    lblh7.Caption = Empty
    remove
End Sub

Private Sub mnueditremovenose_Click()
    picnoseout.Picture = Nothing
    lblnoseout.Caption = Empty
    lblx6.Caption = Empty
    lbly6.Caption = Empty
    lblw6.Caption = Empty
    lblh6.Caption = Empty
    remove
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub

Private Sub mnuformataligbrow_Click()
With frmalignment
    .txtxaxis.Text = x4
    .txtyaxis.Text = y4
    .Caption = "Eyebrow Alignment"
    .Show vbModal
End With
End Sub

Private Sub mnufilenew_Click()
    newsketch
End Sub

Private Sub mnufileopen_Click()
    opensketch
End Sub

Private Sub mnufileprint_Click()
    printsketch
End Sub

Private Sub mnufilesave_Click()
    If tmp.Caption = Empty Then
        saveassketch
    Else
        File = tmp.Caption
        savesketch
    End If
End Sub

Private Sub mnufilesaveas_Click()
    saveassketch
End Sub

Private Sub mnuformatalignbeard_Click()
With frmalignment
    .txtxaxis.Text = lblx9
    .txtyaxis.Text = lbly9
    .Caption = "Beard Alignment"
    .Show vbModal
End With
End Sub

Private Sub mnuformatalignbrow_Click()
With frmalignment
 '   If tmp.Caption <> Empty Then
        .txtxaxis.Text = lblx4
        .txtyaxis.Text = lbly4
  '  Else
   '     .txtxaxis.Text = x4
   '     .txtyaxis.Text = y4
   ' End If
    
        .Caption = "Eyebrow Alignment"
        .Show vbModal
End With
End Sub

Private Sub mnuformataligncap_Click()
With frmalignment
    .txtxaxis.Text = lblx10
    .txtyaxis.Text = lbly10
    .Caption = "Cap Alignment"
    .Show vbModal
End With
End Sub

Private Sub mnuformatalignears_Click()
With frmalignment
    'If tmp.Caption <> Empty Then
        .txtxaxis.Text = lblx3
        .txtyaxis.Text = lbly3
   ' Else
    '    .txtxaxis.Text = x3
    '    .txtyaxis.Text = y3
   ' End If
    
        .Caption = "Ears Alignment"
        .Show vbModal
End With
End Sub

Private Sub mnuformataligneyes_Click()
With frmalignment
    'If tmp.Caption <> Empty Then
        .txtxaxis.Text = lblx5
        .txtyaxis.Text = lbly5
    'Else
     '   .txtxaxis.Text = x5
     '   .txtyaxis.Text = y5
    'End If
    
    .Caption = "Eyes Alignment"
    .Show vbModal
End With
End Sub

Private Sub mnuformatalignglass_Click()
With frmalignment
   ' If tmp.Caption <> Empty Then
        .txtxaxis.Text = lblx8
        .txtyaxis.Text = lbly8
    'Else
    '    .txtxaxis.Text = x8
    '    .txtyaxis.Text = y8
   ' End If
    
    .Caption = "Glass Alignment"
    .Show vbModal
End With
End Sub

Private Sub mnuformatalignhair_Click()
With frmalignment
    'If tmp.Caption <> Empty Then
        .txtxaxis.Text = lblx2
        .txtyaxis.Text = lbly2
        
    'Else
    '    .txtxaxis.Text = x2
    '    .txtyaxis.Text = y2
    'End If
    
    .Caption = "Hair Alignment"
    .Show vbModal
End With
End Sub

Private Sub mnuformatalignjaw_Click()
With frmalignment
    'If tmp.Caption <> Empty Then
        .txtxaxis.Text = lblx1
        .txtyaxis.Text = lbly1
    'Else
    '    .txtxaxis.Text = x1
    '    .txtyaxis.Text = y1
    'End If
    
    .Caption = "Jaw Alignment"
    .Show vbModal
End With
End Sub

Private Sub mnuformatalignlips_Click()
With frmalignment
    'If tmp.Caption <> Empty Then
        .txtxaxis.Text = lblx7
        .txtyaxis.Text = lbly7
    'Else
     '   .txtxaxis.Text = x7
     '   .txtyaxis.Text = y7
    'End If
    
    .Caption = "Lips Alignment"
    .Show vbModal
End With
End Sub

Private Sub mnuformatalignnose_Click()
With frmalignment
    'If tmp.Caption <> Empty Then
        .txtxaxis.Text = lblx6
        .txtyaxis.Text = lbly6
   ' Else
    '    .txtxaxis.Text = x6
    '    .txtyaxis.Text = y6
    'End If
    
    .Caption = "Nose Alignment"
    .Show vbModal
End With
End Sub

Private Sub mnuformatsizebeard_Click()
With frmresize
    'If tmp.Caption <> Empty Then
        .txtwidth.Text = lblw9
        .txtheight.Text = lblh9
    'Else
    '    .txtwidth.Text = w9
    '    .txtheight.Text = h9
    'End If
    
    .Caption = "Beard Resize"
    .Show vbModal
End With
End Sub

Private Sub mnuformatsizebrow_Click()
With frmresize
    'If tmp.Caption <> Empty Then
        .txtwidth.Text = lblw4
        .txtheight.Text = lblh4
    'Else
    '    .txtwidth.Text = w4
    '    .txtheight.Text = h4
    'End If
    
    .Caption = "Eyebrow Resize"
    .Show vbModal
End With
End Sub

Private Sub mnuformatsizecap_Click()
With frmresize
    'If tmp.Caption <> Empty Then
        .txtwidth.Text = lblw10
        .txtheight.Text = lblh10
   ' Else
   '     .txtwidth.Text = w10
   '     .txtheight.Text = h10
   ' End If
    
    .Caption = "Cap Resize"
    .Show vbModal
End With
End Sub

Private Sub mnuformatsizeears_Click()
With frmresize
    'If tmp.Caption <> Empty Then
        .txtwidth.Text = lblw3
        .txtheight.Text = lblh3
    'Else
      '  .txtwidth.Text = w3
     '   .txtheight.Text = h3
    'End If
    
    .Caption = "Ears Resize"
    .Show vbModal
End With
End Sub

Private Sub mnuformatsizeeyes_Click()
With frmresize
    'If tmp.Caption <> Empty Then
        .txtwidth.Text = lblw5
        .txtheight.Text = lblh5
    'Else
    '    .txtwidth.Text = w5
    '    .txtheight.Text = h5
    'End If
    
    .Caption = "Eyes Resize"
    .Show vbModal
End With
End Sub

Private Sub mnuformatsizeglass_Click()
With frmresize
    'If tmp.Caption <> Empty Then
        .txtwidth.Text = lblw8
        .txtheight.Text = lblh8
    'Else
    '    .txtwidth.Text = w8
    '    .txtheight.Text = h8
    'End If
    
    .Caption = "Glass Resize"
    .Show vbModal
End With
End Sub

Private Sub mnuformatsizehair_Click()
With frmresize
    'If tmp.Caption <> Empty Then
        .txtwidth.Text = lblw2
        .txtheight.Text = lblh2
    'Else
    '    .txtwidth.Text = w2
    '    .txtheight.Text = h2
    'End If
    
    .Caption = "Hair Resize"
    .Show vbModal
End With
End Sub

Private Sub mnuformatsizejaw_Click()
With frmresize
    'If tmp.Caption <> Empty Then
        .txtwidth.Text = lblw1
        .txtheight.Text = lblh1
    'Else
    '    .txtwidth.Text = w1
    '    .txtheight.Text = h1
    'End If
    
    .Caption = "Jaw Resize"
    .Show vbModal
End With
End Sub

Private Sub mnuformatsizelips_Click()
With frmresize
    'If tmp.Caption <> Empty Then
        .txtwidth.Text = lblw7
        .txtheight.Text = lblh7
    'Else
    '    .txtwidth.Text = w7
    '    .txtheight.Text = h7
    'End If
    
    .Caption = "Lips Resize"
    .Show vbModal
End With
End Sub

Private Sub mnuformatsizenose_Click()
With frmresize
    'If tmp.Caption <> Empty Then
        .txtwidth.Text = lblw6
        .txtheight.Text = lblh6
    'Else
    '    .txtwidth.Text = w6
    '    .txtheight.Text = h6
    'End If
    
    .Caption = "Nose Resize"
    .Show vbModal
End With
End Sub

Private Sub mnuhelp_Click()

End Sub

Private Sub mnutoolsparts_Click()
    frmpartmanager.Show vbModal
End Sub

Private Sub picbrowout_Change()
    If picbrowout.Picture = Empty Then
        mnuformatalignbrow.Enabled = False
        mnuformatsizebrow.Enabled = False
        mnueditremovebrow.Enabled = False
    Else
        mnuformatalignbrow.Enabled = True
        mnuformatsizebrow.Enabled = True
        mnueditremovebrow.Enabled = True
    End If
End Sub

Private Sub picearsout_Change()
    If picearsout.Picture = Empty Then
        mnuformatalignears.Enabled = False
        mnuformatsizeears.Enabled = False
        mnueditremoveears.Enabled = False
    Else
        mnuformatalignears.Enabled = True
        mnuformatsizeears.Enabled = True
        mnueditremoveears.Enabled = True
    End If
End Sub


Private Sub piceyesout_Change()
    If piceyesout.Picture = Empty Then
        mnuformataligneyes.Enabled = False
        mnuformatsizeeyes.Enabled = False
        mnueditremoveeyes.Enabled = False
    Else
        mnuformataligneyes.Enabled = True
        mnuformatsizeeyes.Enabled = True
        mnueditremoveeyes.Enabled = True
    End If
End Sub

Private Sub picglassout_Change()
    If picglassout.Picture = Empty Then
        mnuformatalignglass.Enabled = False
        mnuformatsizeglass.Enabled = False
        mnueditremoveglass.Enabled = False
    Else
        mnuformatalignglass.Enabled = True
        mnuformatsizeglass.Enabled = True
        mnueditremoveglass.Enabled = True
    End If
End Sub

Private Sub pichairout_Change()
      If pichairout.Picture = Empty Then
        mnuformatalignhair.Enabled = False
        mnuformatsizehair.Enabled = False
        mnueditremovehair.Enabled = False
    Else
        mnuformatalignhair.Enabled = True
        mnuformatsizehair.Enabled = True
        mnueditremovehair.Enabled = True
    End If
End Sub

Private Sub picjawout_Change()
    If picjawout.Picture = Empty Then
        mnuformatalignjaw.Enabled = False
        mnuformatsizejaw.Enabled = False
        mnueditremovejaw.Enabled = False
    Else
        mnuformatalignjaw.Enabled = True
        mnuformatsizejaw.Enabled = True
        mnueditremovejaw.Enabled = True
    End If
End Sub

Private Sub piclipsout_Change()
    If piclipsout.Picture = Empty Then
        mnuformatalignlips.Enabled = False
        mnuformatsizelips.Enabled = False
        mnueditremovelips.Enabled = False
    Else
        mnuformatalignlips.Enabled = True
        mnuformatsizelips.Enabled = True
        mnueditremovelips.Enabled = True
    End If
End Sub

Private Sub picnoseout_Change()
    If picnoseout.Picture = Empty Then
        mnuformatalignnose.Enabled = False
        mnuformatsizenose.Enabled = False
        mnueditremovenose.Enabled = False
    Else
        mnuformatalignnose.Enabled = True
        mnuformatsizenose.Enabled = True
        mnueditremovenose.Enabled = True
    End If
End Sub

Private Sub picbeardout_Change()
    If picbeardout.Picture = Empty Then
        mnuformatalignbeard.Enabled = False
        mnuformatsizebeard.Enabled = False
        mnueditremovebeard.Enabled = False
    Else
        mnuformatalignbeard.Enabled = True
        mnuformatsizebeard.Enabled = True
        mnueditremovebeard.Enabled = True
    End If
End Sub

Private Sub piccapout_Change()
    If piccapout.Picture = Empty Then
        mnuformataligncap.Enabled = False
        mnuformatsizecap.Enabled = False
        mnueditremovecap.Enabled = False
    Else
        mnuformataligncap.Enabled = True
        mnuformatsizecap.Enabled = True
        mnueditremovecap.Enabled = True
    End If
End Sub



Private Sub tmp_Change()
    If tmp.Caption <> Empty Then
        Toolbar1.Buttons.Item(4).Enabled = True
        mnufileprint.Enabled = True
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1 'creates new cartographic sketch
        
    newsketch
    
Case 2 'opens cartographic sketch
    
    opensketch
    
          
Case 3 'saving current work
    
    If tmp.Caption = Empty Then
        saveassketch
    Else
        File = tmp.Caption
        savesketch
    End If
    
Case 4  'printing sketch

    printsketch
    
End Select

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1 'jaw
    pichair.Picture = Nothing
    picears.Picture = Nothing
    picbrow.Picture = Nothing
    piceyes.Picture = Nothing
    picnose.Picture = Nothing
    piclips.Picture = Nothing
    picglass.Picture = Nothing
    picbeard.Picture = Nothing
    piccap.Picture = Nothing
    pichair.Refresh
    picears.Refresh
    picbrow.Refresh
    piceyes.Refresh
    picnose.Refresh
    piclips.Refresh
    picglass.Refresh
    picbeard.Refresh
    piccap.Refresh
    
    Set jawRS = New ADODB.Recordset
    jawRS.Open "jaw", partsConn, adOpenKeyset, adLockReadOnly
    jawRS.MoveFirst
    jawsize
    picview.Cls
    lbljaw.Caption = jawRS!pic
    picjaw.Picture = LoadPicture(App.Path & jawRS!pic)
    picview.PaintPicture picjaw.Picture, x1, y1, w1, h1, , , , , vbSrcAnd
    picjaw.Refresh
    With Toolbar2.Buttons
        .Item(1).Value = tbrPressed
        .Item(2).Value = tbrUnpressed
        .Item(3).Value = tbrUnpressed
        .Item(4).Value = tbrUnpressed
        .Item(5).Value = tbrUnpressed
        .Item(6).Value = tbrUnpressed
        .Item(7).Value = tbrUnpressed
        .Item(8).Value = tbrUnpressed
        .Item(9).Value = tbrUnpressed
        .Item(10).Value = tbrUnpressed
    End With
    
    cmdenabledisable
    
Case 2 'hair
    picjaw.Picture = Nothing
    picears.Picture = Nothing
    picbrow.Picture = Nothing
    piceyes.Picture = Nothing
    picnose.Picture = Nothing
    piclips.Picture = Nothing
    picglass.Picture = Nothing
    picbeard.Picture = Nothing
    piccap.Picture = Nothing
    picjaw.Refresh
    picears.Refresh
    picbrow.Refresh
    piceyes.Refresh
    picnose.Refresh
    piclips.Refresh
    picglass.Refresh
    picbeard.Refresh
    piccap.Refresh
    
    Set hairRS = New ADODB.Recordset
    hairRS.Open "hair", partsConn, adOpenKeyset, adLockReadOnly
    hairRS.MoveFirst
    hairsize
    picview.Cls
    lblhair.Caption = hairRS!pic
    pichair.Picture = LoadPicture(App.Path & hairRS!pic)
    picview.PaintPicture pichair.Picture, x2, y2, w2, h2, , , , , vbSrcAnd
    pichair.Refresh
    With Toolbar2.Buttons
        .Item(1).Value = tbrUnpressed
        .Item(2).Value = tbrPressed
        .Item(3).Value = tbrUnpressed
        .Item(4).Value = tbrUnpressed
        .Item(5).Value = tbrUnpressed
        .Item(6).Value = tbrUnpressed
        .Item(7).Value = tbrUnpressed
        .Item(8).Value = tbrUnpressed
        .Item(9).Value = tbrUnpressed
        .Item(10).Value = tbrUnpressed
    End With
    
    cmdenabledisable
    
Case 3 'ears
    picjaw.Picture = Nothing
    pichair.Picture = Nothing
    picbrow.Picture = Nothing
    piceyes.Picture = Nothing
    picnose.Picture = Nothing
    piclips.Picture = Nothing
    picglass.Picture = Nothing
    picbeard.Picture = Nothing
    piccap.Picture = Nothing
    picjaw.Refresh
    pichair.Refresh
    picbrow.Refresh
    piceyes.Refresh
    picnose.Refresh
    piclips.Refresh
    picglass.Refresh
    picbeard.Refresh
    piccap.Refresh
    
    Set earsRS = New ADODB.Recordset
    earsRS.Open "ears", partsConn, adOpenKeyset, adLockReadOnly
    earsRS.MoveFirst
    earssize
    picview.Cls
    lblears.Caption = earsRS!pic
    picears.Picture = LoadPicture(App.Path & earsRS!pic)
    picview.PaintPicture picears.Picture, x3, y3, w3, h3, , , , , vbSrcAnd
    picears.Refresh
    With Toolbar2.Buttons
        .Item(1).Value = tbrUnpressed
        .Item(2).Value = tbrUnpressed
        .Item(3).Value = tbrPressed
        .Item(4).Value = tbrUnpressed
        .Item(5).Value = tbrUnpressed
        .Item(6).Value = tbrUnpressed
        .Item(7).Value = tbrUnpressed
        .Item(8).Value = tbrUnpressed
        .Item(9).Value = tbrUnpressed
        .Item(10).Value = tbrUnpressed
    End With

    cmdenabledisable
    
Case 4 'eyebrow
    picjaw.Picture = Nothing
    pichair.Picture = Nothing
    picears.Picture = Nothing
    piceyes.Picture = Nothing
    picnose.Picture = Nothing
    piclips.Picture = Nothing
    picglass.Picture = Nothing
    picbeard.Picture = Nothing
    piccap.Picture = Nothing
    picjaw.Refresh
    pichair.Refresh
    picears.Refresh
    piceyes.Refresh
    picnose.Refresh
    piclips.Refresh
    picglass.Refresh
    picbeard.Refresh
    piccap.Refresh
    
    Set browRS = New ADODB.Recordset
    browRS.Open "brow", partsConn, adOpenKeyset, adLockReadOnly
    browRS.MoveFirst
    browsize
    picview.Cls
    lblbrow.Caption = browRS!pic
    picbrow.Picture = LoadPicture(App.Path & browRS!pic)
    picview.PaintPicture picbrow.Picture, x4, y4, w4, h4, , , , , vbSrcAnd
    picbrow.Refresh
    With Toolbar2.Buttons
        .Item(1).Value = tbrUnpressed
        .Item(2).Value = tbrUnpressed
        .Item(3).Value = tbrUnpressed
        .Item(4).Value = tbrPressed
        .Item(5).Value = tbrUnpressed
        .Item(6).Value = tbrUnpressed
        .Item(7).Value = tbrUnpressed
        .Item(8).Value = tbrUnpressed
        .Item(9).Value = tbrUnpressed
        .Item(10).Value = tbrUnpressed
    End With
    
    cmdenabledisable
    
Case 5 'eyes
    picjaw.Picture = Nothing
    pichair.Picture = Nothing
    picears.Picture = Nothing
    picbrow.Picture = Nothing
    picnose.Picture = Nothing
    piclips.Picture = Nothing
    picglass.Picture = Nothing
    picbeard.Picture = Nothing
    piccap.Picture = Nothing
    picjaw.Refresh
    pichair.Refresh
    picears.Refresh
    picbrow.Refresh
    picnose.Refresh
    piclips.Refresh
    picglass.Refresh
    picbeard.Refresh
    piccap.Refresh
    
    Set eyesRS = New ADODB.Recordset
    eyesRS.Open "eyes", partsConn, adOpenKeyset, adLockReadOnly
    eyesRS.MoveFirst
    eyessize
    picview.Cls
    lbleyes.Caption = eyesRS!pic
    piceyes.Picture = LoadPicture(App.Path & eyesRS!pic)
    picview.PaintPicture piceyes.Picture, x5, y5, w5, h5, , , , , vbSrcAnd
    piceyes.Refresh
    With Toolbar2.Buttons
        .Item(1).Value = tbrUnpressed
        .Item(2).Value = tbrUnpressed
        .Item(3).Value = tbrUnpressed
        .Item(4).Value = tbrUnpressed
        .Item(5).Value = tbrPressed
        .Item(6).Value = tbrUnpressed
        .Item(7).Value = tbrUnpressed
        .Item(8).Value = tbrUnpressed
        .Item(9).Value = tbrUnpressed
        .Item(10).Value = tbrUnpressed
    End With
    
    cmdenabledisable
    
Case 6 'nose
    picjaw.Picture = Nothing
    pichair.Picture = Nothing
    picears.Picture = Nothing
    picbrow.Picture = Nothing
    piceyes.Picture = Nothing
    piclips.Picture = Nothing
    picglass.Picture = Nothing
    picbeard.Picture = Nothing
    piccap.Picture = Nothing
    picjaw.Refresh
    pichair.Refresh
    picears.Refresh
    picbrow.Refresh
    piceyes.Refresh
    piclips.Refresh
    picglass.Refresh
    picbeard.Refresh
    piccap.Refresh

    Set noseRS = New ADODB.Recordset
    noseRS.Open "nose", partsConn, adOpenKeyset, adLockReadOnly
    noseRS.MoveFirst
    nosesize
    picview.Cls
    lblnose.Caption = noseRS!pic
    picnose.Picture = LoadPicture(App.Path & noseRS!pic)
    picview.PaintPicture picnose.Picture, x6, y6, w6, h6, , , , , vbSrcAnd
    picnose.Refresh
    With Toolbar2.Buttons
        .Item(1).Value = tbrUnpressed
        .Item(2).Value = tbrUnpressed
        .Item(3).Value = tbrUnpressed
        .Item(4).Value = tbrUnpressed
        .Item(5).Value = tbrUnpressed
        .Item(6).Value = tbrPressed
        .Item(7).Value = tbrUnpressed
        .Item(8).Value = tbrUnpressed
        .Item(9).Value = tbrUnpressed
        .Item(10).Value = tbrUnpressed
    End With
    
    cmdenabledisable
    
Case 7 'lips
    picjaw.Picture = Nothing
    pichair.Picture = Nothing
    picears.Picture = Nothing
    picbrow.Picture = Nothing
    piceyes.Picture = Nothing
    picnose.Picture = Nothing
    picglass.Picture = Nothing
    picbeard.Picture = Nothing
    piccap.Picture = Nothing
    picjaw.Refresh
    pichair.Refresh
    picears.Refresh
    picbrow.Refresh
    piceyes.Refresh
    picnose.Refresh
    picglass.Refresh
    picbeard.Refresh
    piccap.Refresh
    
    Set lipsRS = New ADODB.Recordset
    lipsRS.Open "lips", partsConn, adOpenKeyset, adLockReadOnly
    lipsRS.MoveFirst
    lipssize
    picview.Cls
    lbllips.Caption = lipsRS!pic
    piclips.Picture = LoadPicture(App.Path & lipsRS!pic)
    picview.PaintPicture piclips.Picture, x7, y7, w7, h7, , , , , vbSrcAnd
    piclips.Refresh
    With Toolbar2.Buttons
        .Item(1).Value = tbrUnpressed
        .Item(2).Value = tbrUnpressed
        .Item(3).Value = tbrUnpressed
        .Item(4).Value = tbrUnpressed
        .Item(5).Value = tbrUnpressed
        .Item(6).Value = tbrUnpressed
        .Item(7).Value = tbrPressed
        .Item(8).Value = tbrUnpressed
        .Item(9).Value = tbrUnpressed
        .Item(10).Value = tbrUnpressed
    End With
    
    cmdenabledisable
    
Case 8 'glass
    picjaw.Picture = Nothing
    pichair.Picture = Nothing
    picears.Picture = Nothing
    picbrow.Picture = Nothing
    piceyes.Picture = Nothing
    picnose.Picture = Nothing
    piclips.Picture = Nothing
    picbeard.Picture = Nothing
    piccap.Picture = Nothing
    picjaw.Refresh
    pichair.Refresh
    picears.Refresh
    picbrow.Refresh
    piceyes.Refresh
    picnose.Refresh
    piclips.Refresh
    picbeard.Refresh
    piccap.Refresh
    
    Set glassRS = New ADODB.Recordset
    glassRS.Open "glass", partsConn, adOpenKeyset, adLockReadOnly
    glassRS.MoveFirst
    glasssize
    picview.Cls
    lblglass.Caption = glassRS!pic
    picglass.Picture = LoadPicture(App.Path & glassRS!pic)
    picview.PaintPicture picglass.Picture, x8, y8, w8, h8, , , , , vbSrcAnd
    picglass.Refresh
    With Toolbar2.Buttons
        .Item(1).Value = tbrUnpressed
        .Item(2).Value = tbrUnpressed
        .Item(3).Value = tbrUnpressed
        .Item(4).Value = tbrUnpressed
        .Item(5).Value = tbrUnpressed
        .Item(6).Value = tbrUnpressed
        .Item(7).Value = tbrUnpressed
        .Item(8).Value = tbrPressed
        .Item(9).Value = tbrUnpressed
        .Item(10).Value = tbrUnpressed
    End With

    cmdenabledisable

Case 9 'beard
    picjaw.Picture = Nothing
    pichair.Picture = Nothing
    picears.Picture = Nothing
    picbrow.Picture = Nothing
    piceyes.Picture = Nothing
    picnose.Picture = Nothing
    piclips.Picture = Nothing
    picglass.Picture = Nothing
    piccap.Picture = Nothing
    picjaw.Refresh
    pichair.Refresh
    picears.Refresh
    picbrow.Refresh
    piceyes.Refresh
    picnose.Refresh
    piclips.Refresh
    picglass.Refresh
    piccap.Refresh
    
    Set beardRS = New ADODB.Recordset
    beardRS.Open "beard", partsConn, adOpenKeyset, adLockReadOnly
    beardRS.MoveFirst
    beardsize
    picview.Cls
    lblbeard.Caption = beardRS!pic
    picbeard.Picture = LoadPicture(App.Path & beardRS!pic)
    picview.PaintPicture picbeard.Picture, x9, y9, w9, h9, , , , , vbSrcAnd
    picbeard.Refresh
    With Toolbar2.Buttons
        .Item(1).Value = tbrUnpressed
        .Item(2).Value = tbrUnpressed
        .Item(3).Value = tbrUnpressed
        .Item(4).Value = tbrUnpressed
        .Item(5).Value = tbrUnpressed
        .Item(6).Value = tbrUnpressed
        .Item(7).Value = tbrUnpressed
        .Item(8).Value = tbrUnpressed
        .Item(9).Value = tbrPressed
        .Item(10).Value = tbrUnpressed
    End With

    cmdenabledisable

Case 10 'cap
    picjaw.Picture = Nothing
    pichair.Picture = Nothing
    picears.Picture = Nothing
    picbrow.Picture = Nothing
    piceyes.Picture = Nothing
    picnose.Picture = Nothing
    piclips.Picture = Nothing
    picglass.Picture = Nothing
    picbeard.Picture = Nothing
    picjaw.Refresh
    pichair.Refresh
    picears.Refresh
    picbrow.Refresh
    piceyes.Refresh
    picnose.Refresh
    piclips.Refresh
    picglass.Refresh
    picbeard.Refresh
    
    Set capRS = New ADODB.Recordset
    capRS.Open "cap", partsConn, adOpenKeyset, adLockReadOnly
    capRS.MoveFirst
    capsize
    picview.Cls
    lblcap.Caption = capRS!pic
    piccap.Picture = LoadPicture(App.Path & capRS!pic)
    picview.PaintPicture piccap.Picture, x10, y10, w10, h10, , , , , vbSrcAnd
    piccap.Refresh
    With Toolbar2.Buttons
        .Item(1).Value = tbrUnpressed
        .Item(2).Value = tbrUnpressed
        .Item(3).Value = tbrUnpressed
        .Item(4).Value = tbrUnpressed
        .Item(5).Value = tbrUnpressed
        .Item(6).Value = tbrUnpressed
        .Item(7).Value = tbrUnpressed
        .Item(8).Value = tbrUnpressed
        .Item(9).Value = tbrUnpressed
        .Item(10).Value = tbrPressed
    End With

    cmdenabledisable


End Select

End Sub
