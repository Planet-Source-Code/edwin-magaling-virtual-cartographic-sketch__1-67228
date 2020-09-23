VERSION 5.00
Begin VB.Form frmhelpabout2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13920
   Icon            =   "Help2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Help2.frx":212A
   ScaleHeight     =   8490
   ScaleWidth      =   13920
   StartUpPosition =   1  'CenterOwner
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
      Left            =   12555
      MaskColor       =   &H00FFC0C0&
      Picture         =   "Help2.frx":B2E0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7785
      UseMaskColor    =   -1  'True
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
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
      Left            =   11475
      MaskColor       =   &H00FFC0C0&
      Picture         =   "Help2.frx":F103
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7785
      UseMaskColor    =   -1  'True
      Width           =   1005
   End
End
Attribute VB_Name = "frmhelpabout2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
frmhelpabout3.Show vbModal

End Sub

Private Sub Command2_Click()
Unload Me
frmhelpabout1.Show vbModal
End Sub

