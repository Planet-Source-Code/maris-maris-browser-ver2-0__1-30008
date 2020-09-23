VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Maris URL Chooser:"
   ClientHeight    =   2985
   ClientLeft      =   3930
   ClientTop       =   3945
   ClientWidth     =   4680
   Icon            =   "url.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "url.frx":0442
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go There!"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.geocities.com/Maris_kan"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter URL:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cboAddress_Click
Unload Me
frmBrowser.cboAddress.Text = Text1.Text
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    frmBrowser.brwWebBrowser.Navigate Text1.Text
End Sub

Private Sub Text1_Change()
Form2.Caption = "URL Selected:" & Text1.Text
On Error GoTo Inv_File
Exit Sub
Inv_File:
MsgBox "Invalid File!", vbOKOnly, "Maris Browser"
End Sub

