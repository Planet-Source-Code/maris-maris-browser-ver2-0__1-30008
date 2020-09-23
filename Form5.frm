VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quit Confirmation."
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Caption         =   "No"
      Height          =   375
      Left            =   4680
      Picture         =   "Form5.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Yes"
      Height          =   375
      Left            =   4680
      Picture         =   "Form5.frx":3F88
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Are You Sure You Want to Quit?"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub OKButton_Click()
Unload FrmSysTray
End
End Sub
