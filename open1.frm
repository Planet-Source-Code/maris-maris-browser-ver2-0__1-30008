VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File Selector:"
   ClientHeight    =   3705
   ClientLeft      =   3930
   ClientTop       =   1620
   ClientWidth     =   4260
   Icon            =   "open1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "open1.frx":0442
   ScaleHeight     =   3705
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      Picture         =   "open1.frx":3F88
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   240
      Picture         =   "open1.frx":7ACE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   1920
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
      BackStyle       =   0  'Transparent
      Caption         =   "Maris File Selector:"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Dir1 = "C:\" Then
frmBrowser.MediaPlayer1.Filename = Dir1.Path & File1.Filename
Else
frmBrowser.MediaPlayer1.Filename = Dir1.Path & "\" & File1.Filename
End If
Unload Form8
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
    If mbDontNavigateNow Then Exit Sub
    If Dir1 = "C:\" Then
    frmBrowser.brwWebBrowser.Navigate Dir1.Path & File1.Filename
    Else
    frmBrowser.brwWebBrowser.Navigate Dir1.Path & "\" & File1.Filename
    End If
End Sub


Private Sub File1_Click()
If Dir1 = "C:\" Then
Form1.Caption = "File Selected:" & Dir1.Path & File1.Filename
Else
Form1.Caption = "File Selected:" & Dir1.Path & "\" & File1.Filename
End If
End Sub

Private Sub File1_DblClick()
If Dir1 = "C:\" Then
frmBrowser.MediaPlayer1.Filename = Dir1.Path & File1.Filename
Else
frmBrowser.MediaPlayer1.Filename = Dir1.Path & "\" & File1.Filename
End If
Unload Me
End Sub
