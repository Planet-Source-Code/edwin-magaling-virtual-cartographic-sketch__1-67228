VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmprofile 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Cartographic  Sketch - Profile"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmprofile.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   605
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   717
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdsearch 
      Height          =   405
      Left            =   5580
      Picture         =   "frmprofile.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   765
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print"
      Height          =   450
      Left            =   7785
      TabIndex        =   28
      Top             =   8415
      Width           =   1275
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&Delete"
      Height          =   450
      Left            =   5670
      TabIndex        =   27
      Top             =   8415
      Width           =   1275
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "&Find"
      Height          =   450
      Left            =   1755
      TabIndex        =   24
      Top             =   8415
      Width           =   1275
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit"
      Height          =   450
      Left            =   4365
      TabIndex        =   26
      Top             =   8415
      Width           =   1275
   End
   Begin VB.CommandButton cmdloadsketch 
      Caption         =   "Add Cartographic Sketch"
      Height          =   375
      Left            =   6660
      TabIndex        =   22
      Top             =   4140
      Width           =   3390
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Cl&ose"
      Height          =   450
      Left            =   9090
      TabIndex        =   29
      Top             =   8415
      Width           =   1275
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "&New"
      Height          =   450
      Left            =   450
      TabIndex        =   23
      Top             =   8415
      Width           =   1275
   End
   Begin VB.CommandButton cmdsaveupdate 
      Caption         =   "&Save"
      Height          =   450
      Left            =   3060
      TabIndex        =   25
      Top             =   8415
      Width           =   1275
   End
   Begin MSComCtl2.DTPicker dtcommit 
      Height          =   375
      Left            =   2100
      TabIndex        =   4
      Top             =   1725
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   19660801
      CurrentDate     =   38165
   End
   Begin VB.TextBox txtplace 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2115
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2160
      Width           =   3345
   End
   Begin VB.TextBox txtcase 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2100
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1320
      Width           =   3390
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Additional Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   5970
      TabIndex        =   45
      Top             =   4680
      Width           =   4575
      Begin VB.TextBox txtcarrying 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   20
         Top             =   1485
         Width           =   2535
      End
      Begin VB.TextBox txtlanguage 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   19
         Top             =   1005
         Width           =   2535
      End
      Begin VB.TextBox txtappearance 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   18
         Top             =   555
         Width           =   2535
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   2055
         Width           =   2535
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Appearance:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   49
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Language:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   48
         Top             =   1005
         Width           =   1020
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carrying / Armed With:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   47
         Top             =   1365
         Width           =   1575
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Others:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   46
         Top             =   2055
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Basic Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5460
      Left            =   330
      TabIndex        =   33
      Top             =   2790
      Width           =   5415
      Begin VB.ComboBox cbosex 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   1710
         Width           =   1455
      End
      Begin VB.TextBox txtage 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   10
         Top             =   2085
         Width           =   1395
      End
      Begin VB.TextBox txtlocation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   17
         Top             =   4920
         Width           =   3300
      End
      Begin VB.TextBox txtalias 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   16
         Top             =   4515
         Width           =   3300
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   15
         Top             =   4065
         Width           =   3300
      End
      Begin VB.TextBox txtcomplexion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   14
         Top             =   3660
         Width           =   3300
      End
      Begin VB.TextBox txtbuild 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   13
         Top             =   3255
         Width           =   3300
      End
      Begin VB.TextBox txtweight 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   12
         Top             =   2850
         Width           =   1395
      End
      Begin VB.TextBox txtheight 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   11
         Top             =   2445
         Width           =   1395
      End
      Begin VB.TextBox txtdesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   8
         Top             =   1320
         Width           =   3300
      End
      Begin MSComCtl2.DTPicker dtdate 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   495
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   19660801
         CurrentDate     =   38165
      End
      Begin MSComCtl2.DTPicker dttime 
         Height          =   330
         Left            =   1800
         TabIndex        =   7
         Top             =   945
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   19660802
         CurrentDate     =   38407
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbs."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3285
         TabIndex        =   53
         Top             =   2925
         Width           =   330
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   51
         Top             =   4920
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alias / A.K.A.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   44
         Top             =   4545
         Width           =   1380
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   43
         Top             =   4065
         Width           =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complexion:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   42
         Top             =   3660
         Width           =   1155
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   41
         Top             =   3300
         Width           =   525
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weight:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   40
         Top             =   2850
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   39
         Top             =   2445
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   38
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   37
         Top             =   1725
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   36
         Top             =   1320
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   35
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   34
         Top             =   600
         Width           =   540
      End
   End
   Begin VB.TextBox txtcaseno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2100
      MaxLength       =   20
      TabIndex        =   1
      Top             =   765
      Width           =   3390
   End
   Begin VB.Label tmp 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6345
      TabIndex        =   52
      Top             =   180
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgsketch 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3840
      Left            =   6705
      Stretch         =   -1  'True
      Top             =   225
      Width           =   3300
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   225
      TabIndex        =   50
      Top             =   90
      Width           =   1170
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   12
      X2              =   396
      Y1              =   41
      Y2              =   41
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Place of Incident:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   32
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Commited:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   345
      TabIndex        =   31
      Top             =   1755
      Width           =   1545
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Case:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1335
      TabIndex        =   30
      Top             =   1365
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Case No.:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   525
      TabIndex        =   0
      Top             =   765
      Width           =   1365
   End
End
Attribute VB_Name = "frmprofile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xsketchby As String
Dim xusername, xuserlevel As String

Private Sub disable()
    txtcaseno.Enabled = False
    txtcase.Enabled = False
    dtcommit.Enabled = False
    txtplace.Enabled = False
    dtdate.Enabled = False
    dttime.Enabled = False
    txtdesc.Enabled = False
    cbosex.Enabled = False
    txtage.Enabled = False
    txtheight.Enabled = False
    txtweight.Enabled = False
    txtbuild.Enabled = False
    txtcomplexion.Enabled = False
    txtname.Enabled = False
    txtalias.Enabled = False
    txtlocation.Enabled = False
    txtappearance.Enabled = False
    txtlanguage.Enabled = False
    txtcarrying.Enabled = False
    txtothers.Enabled = False
    cmdloadsketch.Enabled = False
End Sub

Private Sub enable()
    txtcaseno.Enabled = True
    txtcase.Enabled = True
    dtcommit.Enabled = True
    txtplace.Enabled = True
    dtdate.Enabled = True
    dttime.Enabled = True
    txtdesc.Enabled = True
    cbosex.Enabled = True
    txtage.Enabled = True
    txtheight.Enabled = True
    txtweight.Enabled = True
    txtbuild.Enabled = True
    txtcomplexion.Enabled = True
    txtname.Enabled = True
    txtalias.Enabled = True
    txtlocation.Enabled = True
    txtappearance.Enabled = True
    txtlanguage.Enabled = True
    txtcarrying.Enabled = True
    txtothers.Enabled = True
    cmdloadsketch.Enabled = True
End Sub

Private Sub buttonenable()
    cmdnew.Enabled = True
    cmdfind.Enabled = True
    cmdsearch.Visible = False
End Sub

Private Sub buttondisable()
    cmdsaveupdate.Enabled = False
    cmdedit.Enabled = False
    
    cmdDelete.Enabled = False
    cmdprint.Enabled = False
End Sub

Private Sub clearall()
    tmp.Caption = ""
    imgsketch.Picture = Nothing
    txtcaseno.Text = ""
    txtcase.Text = ""
    dtcommit.Value = Date
    txtplace.Text = ""
    dtdate.Value = Date
    dttime.Value = Time
    txtdesc.Text = ""
    cbosex.Text = ""
    txtage.Text = ""
    txtheight.Text = ""
    txtweight.Text = ""
    txtbuild.Text = ""
    txtcomplexion.Text = ""
    txtname.Text = ""
    txtalias.Text = ""
    txtlocation.Text = ""
    txtappearance.Text = ""
    txtlanguage.Text = ""
    txtcarrying.Text = ""
    txtothers.Text = ""
End Sub

Private Sub cmdclose_Click()
    If cmdclose.Caption = "Cl&ose" Then
        Unload Me
    Else
        clearall
        buttonenable
        buttondisable
        disable
        cmdsaveupdate.Caption = "&Save"
        cmdclose.Caption = "Cl&ose"
    End If
    
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Are you to delete this record?", vbQuestion + vbYesNo, "Profile") = vbNo Then
    Exit Sub
End If
    Set mainCmd = New ADODB.Command
    mainStr = "Delete * from profile where caseno='" & txtcaseno.Text & "'"
    With mainCmd
        .ActiveConnection = mainConn
        .CommandType = adCmdText
        .CommandText = mainStr
        .Execute
    End With
    clearall
    MsgBox "Record Successfully Deleted!", vbInformation, "Profile"
    disable
    buttondisable
    buttonenable
    cmdclose.Caption = "Cl&ose"
End Sub

Private Sub cmdedit_Click()
    cmdsaveupdate.Enabled = True
    cmdsaveupdate.Caption = "&Update"
    cmdedit.Enabled = False
    cmdDelete.Enabled = False
    cmdsearch.Visible = False
    enable
    txtcaseno.Enabled = False
End Sub

Private Sub cmdfind_Click()
    clearall
    txtcaseno.Enabled = True
    txtcaseno.SetFocus
    cmdnew.Enabled = False
    cmdfind.Enabled = False
    cmdsearch.Visible = True
    cmdclose.Caption = "&Cancel"
End Sub

Private Sub cmdloadsketch_Click()
    File = Open_File(Me.hWnd) 'show the open file dlg
    If Trim(File) = "" Then Exit Sub ' make sure the file is correct
    imgsketch.Picture = LoadPicture(File) ' load the file
    tmp.Caption = File
End Sub

Private Sub cmdnew_Click()
    clearall
    enable
    txtcaseno.SetFocus
    cmdloadsketch.Enabled = True
    cmdsaveupdate.Enabled = True
    cmdclose.Caption = "&Cancel"
    cmdnew.Enabled = False
    cmdfind.Enabled = False
    cmdprint.Enabled = False
End Sub

Private Sub cmdprint_Click()
    Set mainRS = New ADODB.Recordset
    mainStr = "Select * from profile where caseno='" & txtcaseno.Text & "'"
    mainRS.Open mainStr, mainConn, adOpenKeyset, adLockReadOnly
    Set rptprofile.DataSource = mainRS
    Set rptprofile.Sections(3).Controls("Image1").Picture = LoadPicture(mainRS!FileName)
    rptprofile.Show vbModal
End Sub

Private Sub cmdsaveupdate_Click()
If txtcaseno = "" Then
    MsgBox "Complete Neccessary In formation!", vbInformation, "Profile"
    txtcaseno.SetFocus
    Exit Sub
End If
If txtcase = "" Then
    MsgBox "Complete Neccessary In formation!", vbInformation, "Profile"
    txtcase.SetFocus
    Exit Sub
End If
If txtplace = "" Then
    MsgBox "Complete Neccessary In formation!", vbInformation, "Profile"
    txtplace.SetFocus
    Exit Sub
End If
If txtdesc = "" Then
    MsgBox "Complete Neccessary In formation!", vbInformation, "Profile"
    txtdesc.SetFocus
    Exit Sub
End If
If cbosex = "" Then
    MsgBox "Complete Neccessary In formation!", vbInformation, "Profile"
    cbosex.SetFocus
    Exit Sub
End If
If txtage = "" Then
    MsgBox "Complete Neccessary In formation!", vbInformation, "Profile"
    txtage.SetFocus
    Exit Sub
End If
If txtheight = "" Then
    MsgBox "Complete Neccessary In formation!", vbInformation, "Profile"
    txtheight.SetFocus
    Exit Sub
End If
If txtweight = "" Then
    MsgBox "Complete Neccessary In formation!", vbInformation, "Profile"
    txtweight.SetFocus
    Exit Sub
End If
If txtbuild = "" Then
    MsgBox "Complete Neccessary In formation!", vbInformation, "Profile"
    txtbuild.SetFocus
    Exit Sub
End If
If txtappearance = "" Then
    MsgBox "Complete Neccessary In formation!", vbInformation, "Profile"
    txtappearance.SetFocus
    Exit Sub
End If

If cmdsaveupdate.Caption = "&Save" Then
    Set mainRS = New ADODB.Recordset
    mainStr = "select * from profile where caseno='" & txtcaseno.Text & "'"
    mainRS.Open mainStr, mainConn, adOpenKeyset, adLockReadOnly
    If Not mainRS.EOF And Not mainRS.BOF Then
        MsgBox "Case Number Already Exist!", vbExclamation, "Profile"
        Exit Sub
    End If
    
    If MsgBox("Save this record?", vbQuestion + vbYesNo, "Profile") = vbNo Then
        Exit Sub
    End If
    
    Set userRS = New ADODB.Recordset
    userStr = "Select username from users where status=" & 1
    userRS.Open userStr, userConn, adOpenKeyset, adLockReadOnly
    xsketchby = userRS!Username
    
    Set mainRS = New ADODB.Recordset
    mainRS.Open "profile", mainConn, adOpenKeyset, adLockOptimistic
        With mainRS
            .AddNew
            !FileName = tmp.Caption
            !caseno = txtcaseno.Text
            !Case = txtcase.Text
            !datecommit = dtcommit.Value
            !placeincident = txtplace.Text
            !recdate = dtdate.Value
            !rectime = dttime.Value
            !Description = txtdesc.Text
            !sex = cbosex.Text
            !age = txtage.Text
            !Height = txtheight.Text
            !Weight = txtweight.Text
            !build = txtbuild.Text
            !complexion = txtcomplexion.Text
            !Name = txtname.Text
            !Alias = txtalias.Text
            !Location = txtlocation.Text
            !Appearance = txtappearance.Text
            !language = txtlanguage.Text
            !carrying = txtcarrying.Text
            !others = txtothers.Text
            !sketchby = xsketchby
            .Update
            .Close
        End With
        MsgBox "Record Successfully Saved!", vbInformation, "Profile"
Else
    If MsgBox("Update this record?", vbQuestion + vbYesNo, "Profile") = vbNo Then
        Exit Sub
    End If
        Set mainRS = New ADODB.Recordset
        mainStr = "Select * from profile where caseno='" & txtcaseno.Text & "'"
        mainRS.Open mainStr, mainConn, adOpenKeyset, adLockOptimistic
        With mainRS
            !FileName = tmp.Caption
            !Case = txtcase.Text
            !datecommit = dtcommit.Value
            !placeincident = txtplace.Text
            !recdate = dtdate.Value
            !rectime = dttime.Value
            !Description = txtdesc.Text
            !sex = cbosex.Text
            !age = txtage.Text
            !Height = txtheight.Text
            !Weight = txtweight.Text
            !build = txtbuild.Text
            !complexion = txtcomplexion.Text
            !Name = txtname.Text
            !Alias = txtalias.Text
            !Location = txtlocation.Text
            !Appearance = txtappearance.Text
            !language = txtlanguage.Text
            !carrying = txtcarrying.Text
            !others = txtothers.Text
            .Update
            .Close
        End With
    MsgBox "Record Successfully Updated!", vbInformation, "Profile"
End If

    disable
    buttonenable
    buttondisable
    cmdprint.Enabled = True
    cmdclose.Caption = "Cl&ose"
    
End Sub

Private Sub cmdsearch_Click()
If txtcaseno.Text = "" Then
    MsgBox "Record not Found!", vbExclamation, "Profile"
    txtcaseno.SetFocus
    Exit Sub
End If
    Set mainRS = New ADODB.Recordset
    mainStr = "Select * from profile where caseno='" & txtcaseno.Text & "'"
    mainRS.Open mainStr, mainConn, adOpenKeyset, adLockReadOnly
    If mainRS.EOF And mainRS.BOF Then
        MsgBox "Record not Found!", vbExclamation, "Profile"
        txtcaseno.Text = ""
        txtcaseno.SetFocus
        Exit Sub
    End If
    With mainRS
        If !FileName <> Empty Then
            tmp.Caption = !FileName
            imgsketch.Picture = LoadPicture(tmp.Caption)
        End If
        txtcase.Text = !Case
        dtcommit.Value = !datecommit
        txtplace.Text = !placeincident
        dtdate.Value = !recdate
        dttime.Value = !rectime
        txtdesc.Text = !Description
        cbosex.Text = !sex
        txtage.Text = !age
        txtheight.Text = !Height
        txtweight.Text = !Weight
        txtbuild.Text = !build
        txtcomplexion.Text = !complexion
        txtname.Text = !Name
        txtalias.Text = !Alias
        txtlocation.Text = !Location
        txtappearance.Text = !Appearance
        txtlanguage.Text = !language
        txtcarrying.Text = !carrying
        txtothers.Text = !others
    End With
      
        If xusername = mainRS!sketchby Then
                cmdedit.Enabled = True
                cmdDelete.Enabled = True
        Else
            If xuserlevel = "Administrator" Then
                cmdedit.Enabled = True
                cmdDelete.Enabled = True
            Else
                MsgBox "You can't modify this profile!", vbInformation, "Profile"
                cmdedit.Enabled = False
                cmdDelete.Enabled = False
            End If
        End If
    
    cmdprint.Enabled = True
    
End Sub

Private Sub Form_Load()
    maindbConnect
    
    clearall
    disable
    buttondisable
    cbosex.AddItem "Male"
    cbosex.AddItem "Female"
    
    Set userRS = New ADODB.Recordset
    userStr = "Select username,userlevel from users where status=" & 1
    userRS.Open userStr, userConn, adOpenKeyset, adLockReadOnly
    xusername = userRS!Username
    xuserlevel = userRS!userlevel
    
    InitDlgs 'initalize open dialog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Close Profile?", vbQuestion + vbYesNo, "Profile") = vbYes Then
        Unload Me
    Else
        Cancel = True
    End If
End Sub

Private Sub txtage_Change()
    If Not IsNumeric(txtage.Text) = True Then
        txtage.Text = ""
        'txtage.SetFocus
    End If
End Sub

Private Sub txtcaseno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdsearch.Value = True
    End If
End Sub

Private Sub txtweight_Change()
     If Not IsNumeric(txtweight.Text) = True Then
        txtweight.Text = ""
        'txtweight.SetFocus
    End If
End Sub
