VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmProtector 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Protect the base!!!"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   Moveable        =   0   'False
   Picture         =   "Protector.frx":0000
   ScaleHeight     =   6930
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl mmcLaser 
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   "Wave"
      FileName        =   "\\MSDWT_WC\users\DELKRAN000\Semag\Protector\LASER.WAV"
   End
   Begin VB.Timer timerImages 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3240
      Top             =   3240
   End
   Begin VB.Timer timerAsteroidsA 
      Interval        =   1
      Left            =   720
      Top             =   2880
   End
   Begin VB.Timer timerLaser 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   3960
   End
   Begin VB.Timer timerAsteroidsB 
      Interval        =   1
      Left            =   1200
      Top             =   2880
   End
   Begin VB.Timer timerAsteroidsC 
      Interval        =   1
      Left            =   1680
      Top             =   2880
   End
   Begin VB.Timer timerHit 
      Interval        =   10
      Left            =   1200
      Top             =   3960
   End
   Begin VB.Timer timerAsteroidsD 
      Interval        =   1
      Left            =   2160
      Top             =   2880
   End
   Begin VB.Timer timerBoom 
      Interval        =   250
      Left            =   1680
      Top             =   3960
   End
   Begin VB.Timer timerSpecial 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   2280
      Top             =   4440
   End
   Begin VB.Timer timerS_Lasers 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   2280
      Top             =   4920
   End
   Begin VB.Image Turret_Y 
      Height          =   450
      Left            =   1680
      Picture         =   "Protector.frx":C2EC2
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   5880
      Picture         =   "Protector.frx":C32A8
      Top             =   120
      Width           =   930
   End
   Begin VB.Image Image4 
      Height          =   225
      Left            =   120
      Picture         =   "Protector.frx":C3696
      Top             =   120
      Width           =   735
   End
   Begin VB.Image imgSpecial 
      Height          =   270
      Left            =   1920
      Picture         =   "Protector.frx":C3A5E
      Top             =   1920
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Label lblScore 
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
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   210
   End
   Begin VB.Label lblHull 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   6960
      TabIndex        =   0
      Top             =   0
      Width           =   510
   End
   Begin VB.Image Asteroid_X 
      Height          =   315
      Left            =   720
      Picture         =   "Protector.frx":C3FBC
      Top             =   3480
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Asteroid_Y 
      Height          =   315
      Left            =   1200
      Picture         =   "Protector.frx":C433B
      Top             =   3480
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   4320
      X2              =   5040
      Y1              =   5520
      Y2              =   6480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   3360
      X2              =   2640
      Y1              =   5520
      Y2              =   6480
   End
   Begin VB.Image Turret 
      Height          =   450
      Index           =   0
      Left            =   2520
      Picture         =   "Protector.frx":C46D0
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image Turret 
      Height          =   450
      Index           =   1
      Left            =   4920
      Picture         =   "Protector.frx":C4AB6
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   1
      Left            =   6840
      Picture         =   "Protector.frx":C4E9C
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   2
      Left            =   7320
      Picture         =   "Protector.frx":C524E
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   5
      Left            =   5880
      Picture         =   "Protector.frx":C5600
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   7
      Left            =   5400
      Picture         =   "Protector.frx":C59B2
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   4
      Left            =   4440
      Picture         =   "Protector.frx":C5D64
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   6
      Left            =   3960
      Picture         =   "Protector.frx":C6116
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   0
      Left            =   120
      Picture         =   "Protector.frx":C64C8
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   8
      Left            =   3480
      Picture         =   "Protector.frx":C687A
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   9
      Left            =   3000
      Picture         =   "Protector.frx":C6C2C
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   3
      Left            =   6360
      Picture         =   "Protector.frx":C6FDE
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   10
      Left            =   2040
      Picture         =   "Protector.frx":C7390
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   11
      Left            =   1560
      Picture         =   "Protector.frx":C7742
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   12
      Left            =   1080
      Picture         =   "Protector.frx":C7AF4
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   13
      Left            =   600
      Picture         =   "Protector.frx":C7EA6
      Top             =   6480
      Width           =   255
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   0
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   1
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   2
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   3
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   4
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   5
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   6
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   7
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   8
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   9
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   10
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   11
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   12
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   13
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   14
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Line S_Laser 
      BorderColor     =   &H00FF0000&
      Index           =   15
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   0
      Left            =   0
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   1
      Left            =   480
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   2
      Left            =   960
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   3
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   4
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   5
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   6
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   7
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   8
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   9
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   10
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   11
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   12
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   13
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   14
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Energy 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   15
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   63
      Left            =   7200
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":C8258
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   62
      Left            =   6720
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":C8882
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   61
      Left            =   6240
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":C8EAC
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   60
      Left            =   5760
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":C94D6
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   59
      Left            =   5280
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":C9B00
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   58
      Left            =   4800
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CA12A
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   57
      Left            =   4320
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CA754
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   56
      Left            =   3840
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CAD7E
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   55
      Left            =   3360
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CB3A8
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   54
      Left            =   2880
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CB9D2
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   53
      Left            =   2400
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CBFFC
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   52
      Left            =   1920
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CC626
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   51
      Left            =   1440
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CCC50
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   50
      Left            =   960
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CD27A
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   49
      Left            =   480
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CD8A4
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   48
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CDECE
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   47
      Left            =   7200
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CE4F8
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   46
      Left            =   6720
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CEB22
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   45
      Left            =   6240
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CF14C
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   44
      Left            =   5760
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CF776
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   43
      Left            =   5280
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":CFDA0
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   42
      Left            =   4800
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D03CA
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   41
      Left            =   4320
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D09F4
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   40
      Left            =   3840
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D101E
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   39
      Left            =   3360
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D1648
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   38
      Left            =   2880
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D1C72
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   37
      Left            =   2400
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D229C
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   36
      Left            =   1920
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D28C6
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   35
      Left            =   1440
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D2EF0
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   34
      Left            =   960
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D351A
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   33
      Left            =   480
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D3B44
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   32
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D416E
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   31
      Left            =   7200
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D4798
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   30
      Left            =   6720
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D4DC2
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   29
      Left            =   6240
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D53EC
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   28
      Left            =   5760
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D5A16
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   27
      Left            =   5280
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D6040
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   26
      Left            =   4800
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D666A
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   25
      Left            =   4320
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D6C94
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   24
      Left            =   3840
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D72BE
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   23
      Left            =   3360
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D78E8
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   22
      Left            =   2880
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D7F12
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   21
      Left            =   2400
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D853C
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   20
      Left            =   1920
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D8B66
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   19
      Left            =   1440
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D9190
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   18
      Left            =   960
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D97BA
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   17
      Left            =   480
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":D9DE4
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   16
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DA40E
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   0
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DAA38
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   1
      Left            =   480
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DB062
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   2
      Left            =   960
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DB68C
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   3
      Left            =   1440
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DBCB6
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   4
      Left            =   1920
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DC2E0
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   5
      Left            =   2400
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DC90A
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   6
      Left            =   2880
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DCF34
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   7
      Left            =   3360
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DD55E
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   8
      Left            =   3840
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DDB88
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   9
      Left            =   4320
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DE1B2
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   10
      Left            =   4800
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DE7DC
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   11
      Left            =   5280
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DEE06
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   12
      Left            =   5760
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DF430
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   13
      Left            =   6240
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":DFA5A
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   14
      Left            =   6720
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":E0084
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   15
      Left            =   7200
      MousePointer    =   2  'Cross
      Picture         =   "Protector.frx":E06AE
      Top             =   1080
      Width           =   345
   End
End
Attribute VB_Name = "frmProtector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Index As Integer
Dim Shots As Integer
Dim Disap As Integer
Dim Size As Integer
Dim Pos As Integer
Dim T_Hull As Double
Dim T_Score As Double
Dim Bonus As Integer

Private Sub Asteroid_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If timerBoom.Enabled = False Then
    Else: Asteroid(Index).Picture = Asteroid_X.Picture
        Line1.X1 = Asteroid(Index).Left + Asteroid(Index).Width / 2
        Line2.X1 = Asteroid(Index).Left + Asteroid(Index).Width / 2
        Line1.Y1 = Asteroid(Index).Top + Asteroid(Index).Height / 2
        Line2.Y1 = Asteroid(Index).Top + Asteroid(Index).Height / 2
        Line1.Visible = True
        Line2.Visible = True
        timerLaser.Enabled = True
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mmcLaser.Command = "open"
    mmcLaser.Command = "play"
    
    If timerHit.Enabled = False Then
    Else:
            Line1.X1 = X
            Line2.X1 = X
            Line1.Y1 = Y
            Line2.Y1 = Y
            Line1.Visible = True
            Line2.Visible = True
            timerLaser.Enabled = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If timerBoom.Enabled = False Then
    ElseIf lblHull.Caption <= 25 Then
        timerImages.Enabled = True
        imgSpecial.Visible = True
    ElseIf KeyCode = 32 Then
            timerAsteroidsA.Enabled = False
            timerAsteroidsB.Enabled = False
            timerAsteroidsC.Enabled = False
            timerAsteroidsD.Enabled = False
            timerBoom.Enabled = False
            timerHit.Enabled = False
            timerSpecial.Enabled = True
    End If
End Sub

Private Sub mmcLaser_Done(NotifyCode As Integer)
    mmcLaser.Command = "close"
End Sub

Private Sub timerHit_Timer()
    For z = 0 To 63
        If Asteroid(z).Top >= 6600 Then
            If z >= 0 And z <= 15 Then
                Asteroid(z).Top = -1000
                lblHull.Caption = lblHull.Caption - 1
            ElseIf z >= 16 And z <= 31 Then
                Asteroid(z).Top = -3000
                lblHull.Caption = lblHull.Caption - 1
            ElseIf z >= 32 And z <= 47 Then
                Asteroid(z).Top = -5000
                lblHull.Caption = lblHull.Caption - 1
            ElseIf z >= 48 And z <= 63 Then
                Asteroid(z).Top = -7000
                lblHull.Caption = lblHull.Caption - 1
            End If
        End If
        
        If lblHull.Caption <= 0 Then
            If lblScore.Caption > HIGHSCORE Then
                HIGHSCORE = lblScore.Caption
                frmMain.lblHighScore.Caption = HIGHSCORE
            End If
            
            MsgBox ("The Base Has Been Lost!!!")
            
            Load frmMain
            frmMain.Show
            Unload Me
        End If
    Next z
End Sub

Private Sub timerBoom_Timer()
    For z = 0 To 63
    If Asteroid(z).Picture = Asteroid_X.Picture Then
        
        lblScore.Caption = lblScore.Caption + SCORE
        
        If lblScore.Caption / T_Hull = 1 Then
            lblHull.Caption = lblHull.Caption + 50
            T_Hull = T_Hull + 25000
        ElseIf lblScore.Caption / T_Score = 1 Then
            lblScore.Caption = lblScore.Caption + Bonus
            T_Score = T_Score + 15000
        End If
        
        If z >= 0 And z <= 15 Then
                Asteroid(z).Top = -1000
            ElseIf z >= 16 And z <= 31 Then
                Asteroid(z).Top = -3000
            ElseIf z >= 32 And z <= 47 Then
                Asteroid(z).Top = -5000
            ElseIf z >= 48 And z <= 63 Then
                Asteroid(z).Top = -7000
        End If
        
        Asteroid(z).Picture = Asteroid_Y.Picture
    End If
    Next z
End Sub

Private Sub timerImages_Timer()
    If imgSpecial.Visible = True Then
        imgSpecial.Visible = False
        timerImages.Enabled = False
    End If
End Sub

Private Sub timerSpecial_Timer()
    For t = 0 To 13
        S_Turret(t).Top = S_Turret(t).Top - 120
    Next t
        
    If S_Turret(Index).Top = 6480 Then
        timerSpecial.Enabled = False
        timerS_Lasers.Enabled = True
    End If
End Sub

Private Sub timerS_Lasers_Timer()
    For l = 0 To 15
        If S_Laser(l).Visible = True Then
            Pos = Pos + 480
            
            For z = 0 To 63
                If Asteroid(z).Left = Pos And (Asteroid(z).Top >= -300 And Asteroid(z).Top <= 6300) Then
                    Asteroid(z).Picture = Asteroid_X.Picture
                End If
            Next z
            
            S_Laser(l).Visible = False
        End If
    Next l
    
    If Shots < 15 Then
        If Size <= 3 Then
            Energy(Shots + 1).Visible = True
            Energy(Shots + 1).Height = Energy(Shots + 1).Height + 50
            Energy(Shots + 1).Top = Energy(Shots + 1).Top - 50
            Size = Size + 1
        ElseIf Size = 4 Then
            S_Laser(Shots + 1).Visible = True
            Energy(Shots + 1).Visible = False
            Shots = Shots + 1
            Size = 1
        End If
    ElseIf Shots = 15 And Disap < 15 Then
        S_Laser(Disap + 1).Visible = False
        Disap = Disap + 1
    ElseIf Disap = 15 Then
        If S_Turret(Index).Top = 6960 Then
            For e = 0 To 15
                Energy(e).Height = 100
                Energy(e).Top = 6360
            Next e
            
            Shots = -1
            Disap = -1
            Pos = -480
            lblHull.Caption = lblHull.Caption - 25
        
            timerAsteroidsA.Enabled = True
            timerAsteroidsB.Enabled = True
            timerAsteroidsC.Enabled = True
            timerAsteroidsD.Enabled = True
            timerBoom.Enabled = True
            timerHit.Enabled = True
            timerS_Lasers.Enabled = False
        End If
        
        For t = 0 To 13
            S_Turret(t).Top = S_Turret(t).Top + 120
        Next t
    End If
End Sub

Private Sub timerAsteroidsA_Timer()
    Randomize
    a = Int((15 - 0 + 1) * Rnd + 0)
    
    fall (a)
End Sub

Private Sub timerAsteroidsB_Timer()
    Randomize
    b = Int((31 - 16 + 1) * Rnd + 16)
        
    fall (b)
End Sub

Private Sub timerAsteroidsC_Timer()
    Randomize
    c = Int((47 - 32 + 1) * Rnd + 32)

    fall (c)
End Sub

Private Sub timerAsteroidsD_Timer()
    Randomize
    d = Int((63 - 48 + 1) * Rnd + 48)

    fall (d)
End Sub

Private Sub timerLaser_Timer()
    Line1.Visible = False
    Line2.Visible = False
End Sub

Function fall(ByVal z As Integer) As Integer
    Asteroid(z).Top = Asteroid(z).Top + 100
End Function

Private Sub Form_Load()
    frmProtector.Left = 1800
    frmProtector.Top = 500
    
    timerAsteroidsA.Interval = SPEED
    timerAsteroidsB.Interval = SPEED
    timerAsteroidsC.Interval = SPEED
    
    lblHull = HULL
    
    Shots = -1
    Disap = -1
    Size = 1
    Pos = -480
    T_Hull = 25000
    T_Score = 15000
    Bonus = 10 * SCORE
    
    For e = 0 To 15
        Energy(e).Height = 100
        Energy(e).Top = 6360
    Next e
    
    For l = 1 To 15
        S_Laser(l).X1 = S_Laser(l - 1).X1 + 480
        S_Laser(l).X2 = S_Laser(l - 1).X2 + 480
    Next l
    
    For t = 0 To 13
        S_Turret(t).Top = 6960
        S_Turret(t).Picture = Turret_Y.Picture
    Next t

    For z = 0 To 63
        Asteroid(z).Picture = Asteroid_Y.Picture
    Next z

    For z = 0 To 15
        Asteroid(z).Top = -1000
    Next z
    For z = 16 To 31
        Asteroid(z).Top = -3000
    Next z
    For z = 32 To 47
        Asteroid(z).Top = -5000
    Next z
    For z = 48 To 63
        Asteroid(z).Top = -7000
    Next z
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
End Sub

