VERSION 5.00
Begin VB.Form frmhelptips1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13920
   Icon            =   "Tips1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   13920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "&Next >>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      MaskColor       =   &H00FFC0C0&
      Picture         =   "Tips1.frx":5D52
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   8505
      Left            =   0
      Picture         =   "Tips1.frx":9B75
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13905
   End
End
Attribute VB_Name = "frmhelptips1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
frmhelptips2.Show vbModal
End Sub

