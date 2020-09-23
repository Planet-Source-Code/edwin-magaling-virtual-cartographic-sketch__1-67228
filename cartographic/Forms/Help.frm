VERSION 5.00
Begin VB.Form frmhelpabout1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13710
   ForeColor       =   &H8000000E&
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Help.frx":212A
   ScaleHeight     =   8400
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "&Next >>"
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
      Left            =   12480
      MaskColor       =   &H00FFC0C0&
      Picture         =   "Help.frx":288C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   8535
      Left            =   -240
      Picture         =   "Help.frx":66AF
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   13935
   End
End
Attribute VB_Name = "frmhelpabout1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
    frmhelpabout2.Show vbModal

End Sub

