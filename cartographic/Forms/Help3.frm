VERSION 5.00
Begin VB.Form frmhelpabout3 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13920
   Icon            =   "Help3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Help3.frx":212A
   ScaleHeight     =   8490
   ScaleWidth      =   13920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
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
      Left            =   1800
      MaskColor       =   &H00FFC0C0&
      Picture         =   "Help3.frx":A3AE
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
      Left            =   720
      MaskColor       =   &H00FFC0C0&
      Picture         =   "Help3.frx":E1D1
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7785
      UseMaskColor    =   -1  'True
      Width           =   1005
   End
End
Attribute VB_Name = "frmhelpabout3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
    frmhelpabout2.Show vbModal
End Sub

