VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "Maris Browser"
   ClientHeight    =   6120
   ClientLeft      =   2655
   ClientTop       =   2340
   ClientWidth     =   7335
   ForeColor       =   &H0080FF80&
   Icon            =   "Maris Browser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H80000008&
      Height          =   5895
      Left            =   720
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
      Begin VB.CommandButton Command6 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   24
         Top             =   0
         Width           =   255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   0
         TabIndex        =   22
         Text            =   "Select a choice."
         Top             =   240
         Width           =   1455
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   5025
         Left            =   0
         TabIndex        =   21
         Top             =   600
         Width           =   1455
         ExtentX         =   2566
         ExtentY         =   8864
         ViewMode        =   1
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   -1  'True
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000008&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H0080FF80&
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   5640
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Search"
      ForeColor       =   &H0080FF80&
      Height          =   4095
      Left            =   2160
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
      Begin VB.ListBox List1 
         Height          =   3555
         IntegralHeight  =   0   'False
         ItemData        =   "Maris Browser.frx":0442
         Left            =   0
         List            =   "Maris Browser.frx":0444
         TabIndex        =   16
         Top             =   240
         Width           =   4935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Close!"
         Height          =   195
         Left            =   4080
         TabIndex        =   18
         Top             =   3840
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000008&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   1
         Left            =   0
         ScaleHeight     =   195
         ScaleWidth      =   4905
         TabIndex        =   17
         Top             =   3840
         Width           =   4965
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   4095
      Left            =   7080
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   255
         Left            =   3240
         Picture         =   "Maris Browser.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Select File"
         Height          =   255
         Left            =   240
         Picture         =   "Maris Browser.frx":1986
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
         Height          =   3765
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   4290
         AudioStream     =   -1
         AutoSize        =   -1  'True
         AutoStart       =   -1  'True
         AnimationAtStart=   -1  'True
         AllowScan       =   -1  'True
         AllowChangeDisplaySize=   -1  'True
         AutoRewind      =   -1  'True
         Balance         =   0
         BaseURL         =   ""
         BufferingTime   =   5
         CaptioningID    =   ""
         ClickToPlay     =   -1  'True
         CursorType      =   0
         CurrentPosition =   -1
         CurrentMarker   =   0
         DefaultFrame    =   ""
         DisplayBackColor=   0
         DisplayForeColor=   16777215
         DisplayMode     =   0
         DisplaySize     =   4
         Enabled         =   -1  'True
         EnableContextMenu=   -1  'True
         EnablePositionControls=   -1  'True
         EnableFullScreenControls=   0   'False
         EnableTracker   =   -1  'True
         Filename        =   "c:\baby.avi"
         InvokeURLs      =   -1  'True
         Language        =   -1
         Mute            =   0   'False
         PlayCount       =   1
         PreviewMode     =   0   'False
         Rate            =   1
         SAMILang        =   ""
         SAMIStyle       =   ""
         SAMIFileName    =   ""
         SelectionStart  =   -1
         SelectionEnd    =   -1
         SendOpenStateChangeEvents=   -1  'True
         SendWarningEvents=   -1  'True
         SendErrorEvents =   -1  'True
         SendKeyboardEvents=   0   'False
         SendMouseClickEvents=   0   'False
         SendMouseMoveEvents=   0   'False
         SendPlayStateChangeEvents=   -1  'True
         ShowCaptioning  =   0   'False
         ShowControls    =   -1  'True
         ShowAudioControls=   -1  'True
         ShowDisplay     =   0   'False
         ShowGotoBar     =   0   'False
         ShowPositionControls=   -1  'True
         ShowStatusBar   =   -1  'True
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   0   'False
         Volume          =   0
         WindowlessVideo =   0   'False
      End
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   7095
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   11025
      ExtentX         =   19447
      ExtentY         =   12515
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   3  'Align Left
      Height          =   5445
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   3
      Top             =   675
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   9604
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            Key             =   "Home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6180
      Top             =   1500
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      Picture         =   "Maris Browser.frx":2EC6
      ScaleHeight     =   675
      ScaleWidth      =   7335
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton Command4 
         Caption         =   "Show!"
         Height          =   195
         Left            =   4080
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   4560
         Top             =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Report Bugs"
         Height          =   495
         Left            =   11040
         Picture         =   "Maris Browser.frx":6A0C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox Picture1 
         FillColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   5880
         ScaleHeight     =   195
         ScaleWidth      =   1755
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cboAddress 
         Height          =   315
         ItemData        =   "Maris Browser.frx":7F4C
         Left            =   240
         List            =   "Maris Browser.frx":7F5F
         TabIndex        =   2
         Text            =   "Enter URL Here:"
         ToolTipText     =   "URL"
         Top             =   360
         Width           =   3805
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   8400
         TabIndex        =   14
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "     Browser's Progress:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   13
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "|> Media Player"
         Height          =   255
         Left            =   9120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Address:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Tag             =   "&Address:"
         Top             =   120
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2670
      Top             =   2325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maris Browser.frx":7FBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maris Browser.frx":829E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maris Browser.frx":8580
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maris Browser.frx":93D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maris Browser.frx":BB88
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maris Browser.frx":E33C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   720
      TabIndex        =   12
      Top             =   7820
      Width           =   10965
   End
   Begin VB.Menu open 
      Caption         =   "&Open"
      Index           =   1
   End
   Begin VB.Menu options 
      Caption         =   "&Options"
      Index           =   1
   End
   Begin VB.Menu source 
      Caption         =   "&View Source"
   End
   Begin VB.Menu mnuFindFiles 
      Caption         =   "&Search"
   End
   Begin VB.Menu editor 
      Caption         =   "&DHTML Editor"
   End
   Begin VB.Menu mnuShow 
      Caption         =   "&Show I-Options"
   End
   Begin VB.Menu print 
      Caption         =   "&Print"
   End
   Begin VB.Menu exit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dayAdded As Boolean
Dim NewLocation As String
Dim Today As String
Dim TodayInHistory As Integer
Dim ThisDayName As String
Dim SlashNumber
Dim Position
Dim OldLocation As String
Dim KeyNumber
Dim DayNumber As Integer
Dim nodCN As Node
Dim nodUrl As Node
Dim length
Dim tex

Dim cg
Dim y

Dim version
Dim xadd1
Dim xadd2
Dim xadd3
Dim xadd4
Dim xadd5
Dim xadd6
Dim xadd7
Dim xadd8
Dim xadd9
Dim xadd10
Dim maxh
Dim xal
Dim x
Dim pos1
Dim pos2
Dim pos3
Dim pos4
Dim size

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public AllowPopup As Boolean 'This is for Pop-up windows
Dim WithEvents FormSys As FrmSysTray
Attribute FormSys.VB_VarHelpID = -1
Public StartingAddress As String
Dim mbDontNavigateNow As Boolean
Dim CurrentPercent As Integer

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Dim PicHeight%, hLB&, FileSpec$, UseFileSpec%
Dim TotalDirs%, TotalFiles%, Running%

Const vbBackslash = "\"
Const vbAllFiles = "*.*"
Const vbKeyDot = 46

Private Sub SaveHistory()
Close
Open App.Path & "\history.txt" For Output As #1
Dim currentNode As Node
For Each currentNode In treeHistory.Nodes
Select Case currentNode.Text
    Case "Sunday"
        Print #1, "Sunday"
    Case "Monday"
        Print #1, "Monday"
    Case "Tuesday"
        Print #1, "Tuesday"
    Case "Wednesday"
        Print #1, "Wednesday"
    Case "Thursday"
        Print #1, "Thursday"
    Case "Friday"
        Print #1, "Friday"
    Case "Saturday"
        Print #1, "Saturday"
    Case Else
        ' If currentNode.text is not a day, it might
        ' either a computer name or a complete URL
        If currentNode.Children > 0 Then
        ' currentNode.Children > 0 means currentNode.text
        ' is a computer name, then print one tab
            Print #1, vbTab; currentNode.Text
        Else
        ' currentNode.text is a complete URL, then print
        ' two tabs
            Print #1, vbTab; vbTab; currentNode.Text
        End If
End Select
Next currentNode
Close #1
End Sub
Private Sub LoadHistory()
On Error Resume Next
' This code will search a text file for tab to create a
' treeview depending on the tab number.
' I found this code on PSC, but I modified it to suit my
' needs
Dim tree_nodes() As Node
fnum = FreeFile
' Initialize KeyNumber and DayNumber
KeyNumber = 1
DayNumber = 1
' Open the history file
file_name = App.Path & "\history.txt"
    Open file_name For Input As fnum
    
    treeHistory.Nodes.Clear
    Do While Not EOF(fnum)
        ' Get a line.
        Line Input #fnum, text_line

        ' Find the level of indentation.
        Level = 1
        Do While Left$(text_line, 1) = vbTab
            Level = Level + 1
            text_line = Mid$(text_line, 2)
        Loop

        ' Make room for the new node.
        If Level > num_nodes Then
            num_nodes = Level
            ReDim Preserve tree_nodes(1 To num_nodes)
        End If

        ' Add the new node.
        Select Case Level
        Case 1
        ' If Level = 1, that means we have a day name
            Set tree_nodes(Level) = treeHistory.Nodes.Add(, , "day" & KeyNumber, text_line, 1)
                ' keyNumber will be used later in this
                ' sub and in the DeleteHistory sub
                KeyNumber = KeyNumber + 1
                If text_line = Today Then
                    ' Expand the day node
                    tree_nodes(Level).Expanded = True
                    ' TodayInHistory will be used later
                    ' in this sub and in the
                    ' DeleteHistory sub
                    TodayInHistory = 1
                    ' dayAdded will be used in the
                    ' AddToday sub
                    dayAdded = True
                    ' Today will be used in the AddToday
                    ' sub
                    Today = "day" & (KeyNumber - 1)
                    ' DayNumber will be used in the
                    ' DeleteHistory sub
                    DayNumber = KeyNumber
                End If
        Case 2
        ' If Level = 2, that means we have a computer
        ' name
            ' If TodayInHistory = 0, that means that
            ' today's name was not added to the history
            ' tree from the saved file yet. For example,
            ' if today is Wednesday and the loaded node
            ' is Tuesday (or Monday), this node will not
            ' be used to add a URL to history while
            ' using the browser, it will be used only
            ' to view the saved history, that's why there
            ' is no need to create a key for it
            If TodayInHistory = 0 Then
                Set tree_nodes(Level) = treeHistory.Nodes.Add(tree_nodes(Level - 1), tvwChild, , text_line, 2, 3)
            Else
            ' If TodayInHistory = 1, that means that today's
            ' name was added from the saved file, so a key is
            ' necessary to prevent adding a URL that is already
            ' in the history tree
                
            End If
        Case Else
            ' Same explanation as above
            If TodayInHistory = 0 Then
                Set tree_nodes(Level) = treeHistory.Nodes.Add(tree_nodes(Level - 1), tvwChild, , text_line, 4)
            Else
            End If
       End Select
    Loop

    Close fnum
End Sub
Public Sub TestToday()
' This sub is used to find out what day is today
Select Case Weekday(Now())
    Case 1
        Today = "Sunday"
    Case 2
        Today = "Monday"
    Case 3
        Today = "Tuesday"
    Case 4
        Today = "Wednesday"
    Case 5
        Today = "Thursday"
    Case 6
        Today = "Friday"
    Case 7
        Today = "Saturday"
End Select
' ThisDayName will be used in the AddToday sub
ThisDayName = Today
End Sub
Public Sub urlTest()
SlashNumber = 0
NewLocation = ""

length = Len(OldLocation)
' Count the slash number
For Position = 1 To length
    If Mid(OldLocation, Position, 1) = "/" Then
        SlashNumber = SlashNumber + 1
    End If
Next Position

Select Case SlashNumber
    Case 0
    ' Example : www.yahoo.com
        ' If there are not any slashes in the URL then
        ' there is no need to change it
        NewLocation = OldLocation
        ' If a slash is not added at the end of
        ' OldLocation, this will generate en error as
        ' NewLocation and OldLocation are used as keys
        ' in the TreeHistory
        OldLocation = OldLocation & "/"
        AddComputerNameToHistory
    Case 1
    ' Example : www.yahoo.com/r
        ' Call the OneSlashURL sub
        OneSlashURL
    Case 2
        If Left(OldLocation, 7) <> "http://" Then
        ' Example : www.yahoo.com/r/m1
            ' Call the OneSlashURL sub
            OneSlashURL
        Else
        ' Example : http://www.yahoo.com
            ' Call the TwoSlashURL sub
            TwoSlashURL
        End If
    Case Else
        If Left(OldLocation, 7) <> "http://" Then
        ' Example : www.yahoo.com/homer/?http://greetings.yahoo.com
            ' Call the OneSlashURL sub
            OneSlashURL
        Else
        ' Example : http://www.yahoo.com/r/m1
            ' Call the ThreeSlashURL sub
            ThreeSlashURL
        End If
End Select
End Sub
Public Sub OneSlashURL()
' This sub is used to retrieve the computer name from a
' URL if it looks like this : "www.yahoo.com/r/m1"
SlashNumber = 0
Position = 1
NewLocation = "" ' Null string

' The computer name in a URL is located before the first
' slash if there is no "http://" in it
While SlashNumber = 0
    If Mid(OldLocation, Position, 1) = "/" Then
        SlashNumber = SlashNumber + 1
    End If
    ' When the slash number is 1, the computer name is
    ' found
    If (SlashNumber = 0) And (Mid(OldLocation, Position, 1) <> "/") Then
        NewLocation = NewLocation & Mid(OldLocation, Position, 1)
    End If
    Position = Position + 1
Wend
' Call the AddComputerNameToHistory sub
AddComputerNameToHistory
End Sub

Public Sub TwoSlashURL()
' This sub is used to retrieve the computer name from a
' URL if it looks like this : "http://www.yahoo.com"
    ' If the slash number is 2, add "/" at the end of
    ' the URL so it can be used in the
    ' ThreeSlashURL sub because if the slash number
    ' is smaller than 3, we will have an infinite loop
    OldLocation = OldLocation & "/"
    ' Call the ThreeSlashURL sub
    ThreeSlashURL

End Sub

Public Sub ThreeSlashURL()
' This sub is used to retrieve the computer name from a
' URL if it looks like this : "http://www.yahoo.com/r/m1"
SlashNumber = 0
Position = 1
NewLocation = "" ' Null string

' The computer name in a URL is located between the
' "http://" and the next slash, which makes the slash
' number equals to 3
While SlashNumber < 3
    If Mid(OldLocation, Position, 1) = "/" Then
        SlashNumber = SlashNumber + 1
    End If
    ' When the slash number is 2, the computer name
    ' begins
    If (SlashNumber = 2) And (Mid(OldLocation, Position, 1) <> "/") Then
        NewLocation = NewLocation & Mid(OldLocation, Position, 1)
    End If
    Position = Position + 1
Wend
' Call the AddComputerNameToHistory sub
AddComputerNameToHistory
End Sub

Public Sub AddComputerNameToHistory()
' Error number 35602 is generated when the key is not
' unique. Since the NewLocation (Computer Name) is used
' as a key, the ErrHandler will work like the following:
' if the error number is not 35602, add the NewLocation
' to the HistoryTree. This is easier than assigning a
' different key to each node
On Error GoTo ErrHandler
' If you remove the WebBrowser1.GoBack from the Form_Load
' the NewLocation will be a null string and the
' OldLocation will be"http:///", that's why you will have
' to add "And OldLocation <> "http:///" in the following
' If statement
ErrHandler:
If Err.Number <> 35602 Then
Set nodUrl = treeHistory.Nodes.Add(Today, tvwChild, NewLocation, NewLocation, 2, 3)
' Sort the nodes
nodUrl.Sorted = True
End If
' Call the AddUrlToHistory sub
AddUrlToHistory
End Sub


Public Sub AddUrlToHistory()
' Same explanation as AddComputerNameToHistory
On Error Resume Next
ErrHandler2:
If Err.Number <> 35602 Then
treeHistory.Nodes.Add NewLocation, tvwChild, OldLocation, OldLocation, 4
End If
End Sub

Public Sub AddToday()
If dayAdded = False Then
    Set nodCN = treeHistory.Nodes.Add(, , Today, ThisDayName, 1)
    nodCN.Sorted = True
    nodCN.Expanded = True
    ' Change the value of dayAdded to True to prevent
    ' from adding the today's name to the TreeHistory
    ' again
    dayAdded = True
End If
End Sub

Public Sub DeleteHistory()
' In the LoadHistory sub the KeyNumber is increased by 1
' each time a name of a day is found. If today's name is
' found, the value in KeyNumber will be assigned to
' DayNumber, and the value 1 is assigned to
' TodayInHistory. If there are more days (after today) in
' the history file, the KeyNumber will increase and
' becomes greater than DayNumber.
' Here is an example of how the DeleteHistory sub works:
' if today is Wednesday, and Thursday was found in the
' history file, that means that this is the last week's
' history and it has to be cleared.
' But if today's name was not found in the history file,
' the value of KeyNumber will not be assigned to
' DayNumber (in the LoadHistory sub) which means that the
' value of KeyNumber will be greater than DayNumber and
' the history file will be cleared. To prevent that from
' happening, TodayInHistory is also used in the
' DeleteHistory sub like the following:

If (DayNumber < KeyNumber) And (TodayInHistory = 1) Then
    Open App.Path & "\history.txt" For Output As #4
    Close #4
    ' The TreeHistory must be cleared or else the old
    ' history will still be visible in it
    treeHistory.Nodes.Clear
    ' dayAdded will be used in the AddToday sub
    dayAdded = False
End If
End Sub

Private Function HyperJump(ByVal URL As String) As Long
    HyperJump = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function


Private Sub brwWebBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Label4.Caption = "Document Loaded Succesfully"
End Sub

Private Sub brwWebBrowser_DownloadBegin()

Label4.Caption = "Starting to Download...."
End Sub


Private Sub brwWebBrowser_NewWindow2(ppDisp As Object, Cancel As Boolean)
If AllowPopup = True Then
    Cancel = False
    DoEvents
ElseIf AllowPopup = False Then
    Cancel = True
End If
End Sub

Private Sub brwWebBrowser_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
Label4.Caption = "Reading " & Progress & "  of  " & ProgressMax

CurrentPercent = CurrentPercent + 1
If CurrentPercent < 101 Then
  UpdateProgress Picture1(0), CurrentPercent
Else
  timTimer.Enabled = False
  CurrentPercent = 0
  End If
End Sub

Private Sub brwWebBrowser_SetSecureLockIcon(ByVal SecureLockIcon As Long)
Label4.Caption = "Secured"
End Sub

Private Sub brwWebBrowser_StatusTextChange(ByVal Text As String)
Label4.Caption = Text
End Sub


Private Sub Combo1_Click()
Dim nFolder As SpecialShellFolderIDs
  Dim pidl As Long
  Dim cbpidl As Integer
  Dim abpidl() As Byte
  Dim avpidl As Variant
  Dim sPath As Long
  
  Label5 = ""
  DoEvents
  MousePointer = vbHourglass
  
  nFolder = Combo1.ItemData(Combo1.ListIndex)
  
  ' Get the pointer to the folder's item ID list from
  ' it's specified folder ID, returns 0 on success
  If SHGetSpecialFolderLocation(hwnd, nFolder, pidl) = NOERROR Then
    If pidl Then
      
      ' Get the folder's item ID list size
      cbpidl = GetPIDLSize(pidl)
      If cbpidl Then
        
        ' Reallocate the byte array and copy the folder's item ID list to the array.
        ReDim abpidl(cbpidl - 1)   ' array is zero based
        MoveMemory abpidl(0), ByVal pidl, cbpidl
        
        ' Load the pidl's byte aray into the variant, tada, a SAFEARRAY...
        avpidl = abpidl
        
        ' Navigate the browser to the folder's pidl...!!!
        WebBrowser1.Navigate2 avpidl
        WebBrowser1.Visible = True
        
      End If   ' cbpidl
      
      ' Free the memory the shell allocated for the pidl
      Call CoTaskMemFree(pidl)
      
    End If   ' pidl
  End If   ' SHGetSpecialFolderLocation

  ' Show what's happening with the folder...
  If (pidl = 0) Then
    Label5 = "The folder does not exist on this system."
  
  Else
    Label5 = GetSpecialFolderPath(hwnd, nFolder)
  End If
  
  MousePointer = vbDefault

End Sub

Private Sub Command1_Click()
Form4.Show
End Sub


Private Sub Command2_Click()
Form8.Show
End Sub

Private Sub Command3_Click()
Frame1.Visible = False
End Sub

Private Sub Command4_Click()
Frame2.Visible = True
Command4.Visible = False
End Sub

Private Sub Command5_Click()
Frame2.Visible = False
Command4.Visible = True
End Sub


Private Sub Command6_Click()
Frame3.Visible = False
mnuShow.Enabled = True
End Sub

Private Sub editor_Click()
Form9.Show
End Sub

Private Sub email_Click()
frmBrowser1.Show
End Sub

Private Sub exit_Click()
Form5.Show
End Sub




Private Sub Form_Load()

 Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen

Splash.Show
      
Dim rtn As Long
rtn = SetWindowWord(Splash.hwnd, GWW_HWNDPARENT, frmBrowser.hwnd) 'let both forms load together

Call Wait(1)
Splash.Progress.Value = 10
Call Wait(1)
Splash.Progress.Value = 20
Splash.Label3.Caption = "Loading Option Settings..."
Call Wait(1)
Splash.Progress.Value = 30
Call Wait(1)
Splash.Progress.Value = 40
Call Wait(1)
Splash.Progress.Value = 50
Splash.Label3.Caption = "Initializing..."
Call Wait(1)
Splash.Progress.Value = 60
Splash.Label3.Caption = "Loading..."
Call Wait(1)
Splash.Progress.Value = 70
Splash.Label3.Caption = "Loading...."
Call Wait(1)
Splash.Progress.Value = 80
Splash.Label3.Caption = "Loading....."
Call Wait(1)
Splash.Progress.Value = 90
Splash.Label3.Caption = "Loading......"
Call Wait(1)
Splash.Progress.Value = 100
Splash.Label3.Caption = "Presenting Maris Browser..."
frmBrowser.Visible = True
Call Wait(1)

Unload Splash

frmBrowser.SetFocus


 
 
 
ScaleMode = vbPixels

With Combo1
    .AddItem "Favorites"
    '.ItemData(.NewIndex) = CSIDL_FAVORITES
    .AddItem "Cookies"
    .ItemData(.NewIndex) = CSIDL_COOKIES
    .AddItem "History"
    .ItemData(.NewIndex) = CSIDL_HISTORY
  
    .ListIndex = 0   ' invokes a Combo1_Click
  End With
  
  
  
    PicHeight% = Picture1(1).Height
    hLB& = List1.hwnd
    ' This speeds things a bit but will consume close to 6MB of memory...!!!
    SendMessage hLB&, LB_INITSTORAGE, 30000&, ByVal 30000& * 200
'    Move (Screen.Width - Width) * 0.5, (Screen.Height - Height) * 0.5
    
Label2.Caption = "Time is " + Format(Now, "hh:mm:ssAM/PM")
Label2.ToolTipText = "Today is " + Format(Date, "dddd") & ", " & Format$(Date, "dd-mm-yyyy")
Label2.Refresh

picAddress.Refresh
   Set FormSys = New FrmSysTray
   Load FormSys
   Set FormSys.FSys = Me
    On Error Resume Next
    Me.Show
    'Form_Resize
    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15

    'If Len(StartingAddress) > 0 Then
        'Form3.Text2.Text = cboAddress.Text
        'cboAddress.Text = StartingAddress
     '   timTimer.Enabled = True
        brwWebBrowser.Navigate Form3.Text2.Text 'StartingAddress
      '  End If
    
End Sub

Private Sub brwWebBrowser_DownloadComplete()
    Label4.Caption = "Download Done!"
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName & " - Maris Browser"
End Sub

Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    cboAddress.Text = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
    
End Sub

Private Sub cboAddress_Click()
    
     If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
    cboAddress.AddItem brwWebBrowser.LocationURL
    
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    timTimer.Enabled = True
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub


Private Sub Label3_Click()
Frame1.Visible = True
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.ForeColor = &H80000012
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Label3.ForeColor = &HFF&
End Sub

Private Sub List1_DblClick()
frmBrowser.brwWebBrowser.Navigate List1
End Sub

Private Sub MediaPlayer1_Error()
On Error Resume Next
'MsgBox "Sorry. But there seems to be a problem with the file Format. Check and Try Again.", vbOKOnly, "Maris Browser"
End Sub

Private Sub SearchDirs(curpath$)
    Dim dirs%, dirbuf$(), i%
    
    
    Picture1(1).Cls
    Picture1(1).Print "Searching " & curpath$
    
    
    DoEvents
    If Not Running% Then Exit Sub
    
    
    hItem& = FindFirstFile(curpath$ & vbAllFiles, WFD)
    If hItem& <> INVALID_HANDLE_VALUE Then
        
        Do
           
            If (WFD.dwFileAttributes And vbDirectory) Then
                
                
                If Asc(WFD.cFileName) <> vbKeyDot Then
                    
                    TotalDirs% = TotalDirs% + 1
                    
                    If (dirs% Mod 10) = 0 Then ReDim Preserve dirbuf$(dirs% + 10)
                    dirs% = dirs% + 1
                    dirbuf$(dirs%) = Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
                End If
            
            
            ElseIf Not UseFileSpec% Then
                TotalFiles% = TotalFiles% + 1
            End If
        
        
        Loop While FindNextFile(hItem&, WFD)
        
        
        Call FindClose(hItem&)
    
    End If

    
    If UseFileSpec% Then
        
        SendMessage hLB&, WM_SETREDRAW, 0, 0
        Call SearchFileSpec(curpath$)
        
        SendMessage hLB&, WM_VSCROLL, SB_BOTTOM, 0
        SendMessage hLB&, WM_SETREDRAW, 1, 0
    End If
    
    
    For i% = 1 To dirs%: SearchDirs curpath$ & dirbuf$(i%) & vbBackslash: Next i%
  
End Sub

Private Sub SearchFileSpec(curpath$)

    
    hFile& = FindFirstFile(curpath$ & FileSpec$, WFD)
    If hFile& <> INVALID_HANDLE_VALUE Then
        
        Do
           
            DoEvents
            If Not Running% Then Exit Sub
            
            
            SendMessage hLB&, LB_ADDSTRING, 0, _
                ByVal curpath$ & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
        
        
        Loop While FindNextFile(hFile&, WFD)
        
        
        Call FindClose(hFile&)
    
    End If

End Sub

Private Sub mnuFindFiles_Click()
Frame2.Visible = True
If Running% Then: Running% = False: Exit Sub
    
    Dim drvbitmask&, maxpwr%, pwr%
    On Error Resume Next
    
    FileSpec$ = InputBox("Enter Filename to be Searched:" & vbCrLf & vbCrLf & _
                                    "Searching will begin at drive A and continue " & _
                                    "until no more drives are found.  " & _
                                    "Double-Click the File to run it and " & _
                                    "Click Stop! at any time." & vbCrLf & _
                                    "The * and ? wildcards can be used.", _
                                    "Find File(s)", "*.*")
    
   If Len(FileSpec$) = 0 Then Exit Sub
    
    MousePointer = 11
    Running% = True
    UseFileSpec% = True
    mnuFindFiles.Caption = "&Stop!"
    
    List1.Clear
    
    drvbitmask& = GetLogicalDrives()
    
    If drvbitmask& Then
        
        
        maxpwr% = Int(Log(drvbitmask&) / Log(2))   ' a little math...
        For pwr% = 0 To maxpwr%
            If Running% And (2 ^ pwr% And drvbitmask&) Then _
                Call SearchDirs(Chr$(vbKeyA + pwr%) & ":\")
        Next
    End If
    
    Running% = False
    UseFileSpec% = False
    mnuFindFiles.Caption = "&Search"
    
    MousePointer = 0

    Picture1(1).Cls
    Picture1(1).Print "File(s) Found: " & List1.ListCount & " items found matching " & """" & FileSpec$ & """"
    Beep
End Sub

Private Sub mnuShow_Click()
Frame3.Visible = True
mnuShow.Enabled = False
End Sub

Private Sub open1_Click()
Form1.Show
End Sub

Private Sub options1_Click()
Form3.Show
End Sub


Private Sub open_Click(Index As Integer)
Form1.Show
End Sub

Private Sub options_Click(Index As Integer)
Form3.Show
End Sub

Private Sub picAddress_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.ForeColor = &H80000012
End Sub


Private Sub print_Click()
Dim eQuery As OLECMDF

    On Error Resume Next
    eQuery = brwWebBrowser.QueryStatusWB(OLECMDID_PRINT)
    If Err.Number = 0 Then
            If eQuery And OLECMDF_ENABLED Then
                brwWebBrowser.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, "", ""
          Else
                MsgBox "The Print command is currently disabled."
            End If
    End If
    If Err.Number <> 0 Then MsgBox "Print command Error: " & Err.Description
End Sub

Private Sub source_Click()
Dim dfg As String
dfg = brwWebBrowser.Document.documentElement.innerHTML
frmDocument.rtftext.Text = dfg
frmDocument.Show
End Sub

Private Sub Timer1_Timer()
Label2.Caption = "Time is " + Format(Now, "hh:mm:ssAM/PM")
Label2.Refresh
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        CurrentPercent = brwWebBrowser.Busy
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Working..."
        
        End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
     
    timTimer.Enabled = True
     
    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
            Label4.Caption = "Going Back..."
            
        Case "Forward"
            brwWebBrowser.GoForward
            Label4.Caption = "Going Forward.."
        Case "Refresh"
            brwWebBrowser.Refresh
            Label4.Caption = "Refreshing....."
        Case "Home"
            brwWebBrowser.GoHome
            Label4.Caption = "Going Home..."
        Case "Search"
            brwWebBrowser.GoSearch
            Label4.Caption = "Going to Search Page..."
        Case "Stop"
            timTimer.Enabled = False
            brwWebBrowser.Stop
            Label4.Caption = "Stopping..."
            Me.Caption = brwWebBrowser.LocationName & " - Maris Browser"
    End Select

End Sub
Private Sub CmdExit_Click()
   Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   FormSys.MeQueryUnload Me, Cancel, UnloadMode
End Sub
Private Sub FormSys_TIcon(F As Form)
   Me.Icon = F.Icon
End Sub
Public Function UpdateProgress(pb As Control, ByVal Percent)
Dim Num$
If Not pb.AutoRedraw Then
pb.AutoRedraw = -1
End If
Refresh
pb.Cls
pb.ScaleWidth = 100
pb.DrawMode = 10
Num$ = Format$(Percent, "###") + "% Completed"
pb.CurrentX = 50 - pb.TextWidth(Num$) / 2
pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(Num$)) / 2
pb.Print Num$
pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
pb.Refresh

End Function


