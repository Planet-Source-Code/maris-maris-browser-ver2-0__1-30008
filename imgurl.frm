VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H8000000E&
   Caption         =   "Maris Image URL Chooser:"
   ClientHeight    =   3195
   ClientLeft      =   2985
   ClientTop       =   4005
   ClientWidth     =   4680
   Icon            =   "imgurl.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go There!"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
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
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cboAddress_Click
Unload Me
Form3.Text1.Text = Text1.Text
End Sub


Private Sub cboAddress_Click()
    frmBrowser.picAddress.Picture = LoadPicture(Text1.Text)
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Text1_Change()
Form7.Caption = "Image URL Selected:" & Text1.Text
On Error GoTo FONTNAME_ERROR
Exit Sub
FONTNAME_ERROR:
MsgBox "Invalid File!"
End Sub
