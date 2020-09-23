VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H8000000A&
   Caption         =   "Maris Browser Options"
   ClientHeight    =   5430
   ClientLeft      =   2655
   ClientTop       =   2340
   ClientWidth     =   6225
   Icon            =   "optios.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      Picture         =   "optios.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      Picture         =   "optios.frx":3F88
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   375
      Left            =   960
      Picture         =   "optios.frx":7ACE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   8493
      _Version        =   393216
      TabOrientation  =   2
      TabHeight       =   520
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Email Me!"
      TabPicture(0)   =   "optios.frx":B614
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Version"
      TabPicture(1)   =   "optios.frx":B630
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Settings"
      TabPicture(2)   =   "optios.frx":B64C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label8"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label9"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Check1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Text1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Command4"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Check2"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "delete"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Check3"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Text2"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Command5"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Command6"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      Begin VB.CommandButton Command6 
         Caption         =   "Use Default"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71640
         Picture         =   "optios.frx":B6A0
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Use Blank "
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73200
         Picture         =   "optios.frx":CBE0
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000008&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   -73440
         TabIndex        =   20
         Text            =   "http://maris.ipfox.com"
         Top             =   4080
         Width           =   3855
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Run Maris Browser on Startup"
         Height          =   495
         Left            =   -74640
         TabIndex        =   19
         Top             =   2400
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CommandButton delete 
         Caption         =   "Delete History"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74520
         Picture         =   "optios.frx":E120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show Browser Speed Bar"
         Height          =   255
         Left            =   -74640
         TabIndex        =   17
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Browse"
         Height          =   375
         Left            =   -70680
         Picture         =   "optios.frx":F660
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -71280
         TabIndex        =   14
         Text            =   "Click on Browse!"
         Top             =   960
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "About Myself"
         Height          =   4815
         Left            =   360
         TabIndex        =   1
         Top             =   0
         Width           =   5295
         Begin VB.PictureBox Picture1 
            Height          =   3255
            Left            =   2160
            Picture         =   "optios.frx":131A6
            ScaleHeight     =   3195
            ScaleWidth      =   2955
            TabIndex        =   2
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label1 
            Caption         =   "Full Name : Maris Kannan"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   4935
         End
         Begin VB.Label Label2 
            Caption         =   "Country of Birth : India"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label3 
            Caption         =   "Age : 16  (16-06-84)"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Hobbies : Cricket, Tennis, Badminton, Volleyball.. and others"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   3600
            Width           =   4695
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Allow Pop-up Windows"
         Height          =   255
         Left            =   -74640
         Picture         =   "optios.frx":14880
         TabIndex        =   16
         Top             =   2160
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Starting URL:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   21
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Add Your Own ToolBar Image!"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   3135
         Left            =   -72120
         TabIndex        =   13
         Top             =   240
         Width           =   2535
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Settings:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   615
         Left            =   -74280
         TabIndex        =   10
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label6 
         Caption         =   "Version 1.0 Released in 2001."
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1095
         Left            =   -73800
         TabIndex        =   8
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Maris Browser"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74040
         TabIndex        =   7
         Top             =   240
         Width           =   3615
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Path As String
Public Function run()
Value& = MsgBox("Would you like MarisBrowser to run at startup?", vbYesNo, "Maris Browser")
If Value& = vbYes Then
'SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "MarisBrowser", App.Path & "\" & App.EXEName & ".exe"
Else
'SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "MarisBrowser", "<NonRun>"
End If
End Function




Private Sub Command1_Click()
If Check2.Value = 1 Then
frmBrowser.Label1.Visible = True
frmBrowser.Picture1(0).Visible = True
Else
If Check2.Value = 0 Then
frmBrowser.Label1.Visible = False
frmBrowser.Picture1(0).Visible = False

RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "Speed Bar", Check2.Value
RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "Allow Pop", Check1.Value
RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "On Startup", Check3.Value
RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "info", Text1.Text
RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "url", Text2.Text

If Check3.Value = 1 Then
run
Else
End If
End If
End If



End Sub

Private Sub Command2_Click()
If Check1.Value = 1 Then
MsgBox "Maris Browser - Allowing Pop-Ups to be displayed" ', vbOKOnly, "Maris Browser"
frmBrowser.AllowPopup = True
Else
If Check1.Value = 0 Then
MsgBox "Maris Browser - Disallowing Pop-Ups to be displayed" ', vbOKOnly, "Maris Browser"
frmBrowser.AllowPopup = False
End If
End If
RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "Speed Bar", Check2.Value
RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "Allow Pop", Check1.Value
RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "On Startup", Check3.Value
RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "info", Text1.Text
RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "url", Text2.Text

If Check3.Value = 1 Then
run
Else
End If
Unload Me
frmBrowser.Show
End Sub

Private Sub Command3_Click()
Unload Me
frmBrowser.Show
End Sub


Private Sub Command4_Click()
Form6.Show
End Sub



Private Sub Command5_Click()
Text2.Text = "about:Blank"
End Sub

Private Sub Command6_Click()
Text2.Text = "www.geocities.com/Mariskan"
End Sub

Private Sub delete_Click()
frmBrowser.cboAddress.Clear
End Sub

Private Sub Form_Load()
RGCreateKey HKEY_LOCAL_MACHINE, SettingsPath
options
End Sub
