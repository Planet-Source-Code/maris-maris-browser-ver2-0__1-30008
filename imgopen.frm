VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Image Selector:"
   ClientHeight    =   3705
   ClientLeft      =   3615
   ClientTop       =   2625
   ClientWidth     =   4260
   Icon            =   "imgopen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "imgopen.frx":0442
   ScaleHeight     =   3705
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   1920
      ReadOnly        =   0   'False
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Maris File Selector: only gif, jpg, tif and bmp extensions accepted."
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3975
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cboAddress_Click
On Error GoTo Inv_File
Unload Me
If Dir1 = "C:\" Then
Form3.Text1.Text = Dir1.Path & File1.Filename
Else
Form3.Text1.Text = Dir1.Path & "\" & File1.Filename
Exit Sub
Inv_File:
MsgBox "Invalid File!", vbOKOnly, "Maris Browser"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Dir1_Click()
File1.Path = Dir1
Refresh
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1
Refresh
End Sub

Private Sub cboAddress_Click()
On Error GoTo Inv_File
If Dir1 = "C:\" Then
frmBrowser.picAddress.Picture = LoadPicture(Dir1.Path & File1)
Else
frmBrowser.picAddress.Picture = LoadPicture(Dir1.Path & "\" & File1)
Inv_File:
MsgBox "Invalid File!", vbOKOnly, "Maris Browser"
End If
End Sub


Private Sub File1_Click()
If Dir1 = "C:\" Then
Form6.Caption = "Image Selected:" & Dir1.Path & File1.Filename
Else
Form6.Caption = "Image Selected:" & Dir1.Path & "\" & File1.Filename
End If
End Sub

Private Sub File1_DblClick()
If Dir1 = "C:\" Then
frmBrowser.brwWebBrowser.Navigate Dir1.Path & File1.Filename
Else
frmBrowser.brwWebBrowser.Navigate Dir1.Path & "\" & File1.Filename
On Error GoTo Inv_File
Exit Sub
Inv_File:
MsgBox "Invalid File!", vbOKOnly, "Maris Browser"
Unload Me
End If
End Sub
