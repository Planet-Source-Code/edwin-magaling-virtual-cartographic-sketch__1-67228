VERSION 5.00
Begin VB.Form frmhelptips2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13905
   Icon            =   "Tips2.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   13905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "Cl&ose"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      MaskColor       =   &H00FFC0C0&
      Picture         =   "Tips2.frx":5D52
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "<< &Back"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      MaskColor       =   &H00FFC0C0&
      Picture         =   "Tips2.frx":9B75
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   8505
      Left            =   0
      Picture         =   "Tips2.frx":D998
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13905
   End
End
Attribute VB_Name = "frmhelptips2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
frmhelptips1.Show vbModal
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

