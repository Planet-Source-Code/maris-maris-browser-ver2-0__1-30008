VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bugz"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form4.frx":0442
   ScaleHeight     =   3555
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   2400
      Left            =   3600
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "I'll Email You!"
      Height          =   375
      Left            =   1560
      Picture         =   "Form4.frx":3F88
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Email Me right away to Maris_kan@yahoo.com"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Have Bugs To Report to Me?"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Command1.Visible = True
End Sub
