VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Protector of Moon Base IX!!!"
   ClientHeight    =   5595
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Main.frx":0000
   ScaleHeight     =   5595
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAce 
      Caption         =   "Ace"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCaptain 
      Caption         =   "Captain"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdNovice 
      Caption         =   "Novice"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "High Score:"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label lblHighScore 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1800
      TabIndex        =   10
      Top             =   720
      Width           =   210
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Choose from 3 Levels of Difficulty ==>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   3090
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H00FF0000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   255
      Left            =   240
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Good luck in protecting Earth's nineth moon base from the asteroid onslaught!!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   705
      Left            =   600
      TabIndex        =   8
      Top             =   4080
      Width           =   2745
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H00FF0000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   255
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "The more difficult the level you play, the more points you will score.  In return you'll have a weaker base hull."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   945
      Left            =   600
      TabIndex        =   7
      Top             =   3000
      Width           =   2745
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H00FF0000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   255
      Left            =   240
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H00FF0000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   255
      Left            =   240
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Use the space bar to fire special weapon.  Be wise when doing this because of its destructive force."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   705
      Left            =   600
      TabIndex        =   6
      Top             =   2040
      Width           =   2745
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Use the mouse buttons to fire primary weapons."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   600
      TabIndex        =   5
      Top             =   1320
      Width           =   2745
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Created By:  Randall E. Delk, Jr."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2985
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Shot From Protector"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   4200
      TabIndex        =   3
      Top             =   3960
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   3645
      Left            =   3480
      Picture         =   "Main.frx":C2EC2
      Top             =   120
      Width           =   3810
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuEnd 
         Caption         =   "End"
      End
   End
   Begin VB.Menu mnuScore 
      Caption         =   "High &Score"
      Begin VB.Menu mnuReset 
         Caption         =   "Reset High Score"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuStory 
         Caption         =   "Story Line"
      End
      Begin VB.Menu mnuBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAce_Click()
    HULL = 100
    SCORE = 100
    SPEED = 1
    
    Load frmProtector
    Unload Me
End Sub

Private Sub cmdCaptain_Click()
    HULL = 200
    SCORE = 50
    SPEED = 1
    
    Load frmProtector
    Unload Me
End Sub

Private Sub cmdNovice_Click()
    HULL = 400
    SCORE = 25
    SPEED = 1
    
    Load frmProtector
    Unload Me
End Sub

Private Sub Form_Load()
    Open "C:\Programming\Semag\Protector\ProtectorHS.txt" For Input As #1
    Input #1, modInfo.HIGHSCORE
    Close
    
    lblHighScore.Caption = HIGHSCORE
    
    Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Open "C:\Programming\Semag\Protector\ProtectorHS.txt" For Output As #1
    Print #1, lblHighScore.Caption
    Close
End Sub

Private Sub mnuEnd_Click()
    End
End Sub

Private Sub mnuReset_Click()
    lblHighScore.Caption = "0"

    Open "C:\Programming\Semag\Protector\ProtectorHS.txt" For Output As #1
    Print #1, lblHighScore.Caption
    Close
End Sub
