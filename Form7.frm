VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Media File:"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form7.frx":0442
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "The Media"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "File:"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
End Sub
