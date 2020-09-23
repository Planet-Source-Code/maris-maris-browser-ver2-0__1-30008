VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "DHTMLED.OCX"
Begin VB.Form Form9 
   BackColor       =   &H80000009&
   Caption         =   "Maris Browser > DHTML Editor"
   ClientHeight    =   7380
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6690
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form9.frx":0442
   ScaleHeight     =   7380
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Save!"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "C:\noname.htm"
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin DHTMLEDLibCtl.DHTMLEdit DHTMLEdit1 
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6495
      ActivateApplets =   -1  'True
      ActivateActiveXControls=   -1  'True
      ActivateDTCs    =   -1  'True
      ShowDetails     =   -1  'True
      ShowBorders     =   -1  'True
      Appearance      =   1
      Scrollbars      =   -1  'True
      ScrollbarAppearance=   1
      SourceCodePreservation=   -1  'True
      AbsoluteDropMode=   -1  'True
      SnapToGrid      =   -1  'True
      SnapToGridX     =   50
      SnapToGridY     =   50
      BrowseMode      =   0   'False
      UseDivOnCarriageReturn=   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Save File in  :"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "   Maris DHTML Editor"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   6015
   End
   Begin VB.Menu open 
      Caption         =   "&Open"
   End
   Begin VB.Menu new 
      Caption         =   "&New File"
   End
   Begin VB.Menu save 
      Caption         =   "&Save Document"
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit Mode"
   End
   Begin VB.Menu preview 
      Caption         =   "&Preview"
   End
   Begin VB.Menu exit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private m_cAC As New cAutoComplete
 
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "You have to type in a filename and add .htm to it", vbOKOnly, "Maris Browser"
Else
Form9.DHTMLEdit1.SaveDocument (Text1.Text)
Label2.Visible = False
Text1.Visible = False
Command1.Visible = False
MsgBox "File: " & Text1.Text & " Saved!", vbOKOnly, "Maris Browser"
End If
End Sub

Private Sub edit_Click()
edit.Enabled = False
preview.Enabled = True
DHTMLEdit1.BrowseMode = False
MsgBox "Running in Edit Mode!", vbOKOnly, "Maris Browser"
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_load()
Text1 = ""
Call m_cAC.Attach(Text1)


edit.Enabled = False
preview.Enabled = True
End Sub

Private Sub new_Click()
DHTMLEdit1.NewDocument
MsgBox "New Document Opened!", vbOKOnly, "Maris Browser"
End Sub

Private Sub open_Click()
Form10.Show
End Sub

Private Sub preview_Click()
edit.Enabled = True
preview.Enabled = False
DHTMLEdit1.BrowseMode = True
MsgBox "Running in Preview Mode!", vbOKOnly, "Maris Browser"
End Sub

Private Sub save_Click()
Label2.Visible = True
Text1.Visible = True
Command1.Visible = True
End Sub
