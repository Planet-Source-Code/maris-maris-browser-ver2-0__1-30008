VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Splash 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maris Browser loading...."
   ClientHeight    =   1935
   ClientLeft      =   2670
   ClientTop       =   2280
   ClientWidth     =   3855
   ControlBox      =   0   'False
   Icon            =   "Display Screen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1935
   ScaleWidth      =   3855
   Begin ComctlLib.ProgressBar Progress 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Line Line 
      X1              =   240
      X2              =   3600
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Checking System settings..."
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait..."
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen
End Sub



