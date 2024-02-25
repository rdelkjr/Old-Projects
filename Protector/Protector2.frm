VERSION 5.00
Begin VB.Form frmProtector 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Protect the base!!!"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   Moveable        =   0   'False
   Picture         =   "Protector2.frx":0000
   ScaleHeight     =   6930
   ScaleWidth      =   7560
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&End"
      Height          =   495
      Left            =   3200
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer timerHull 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   4920
   End
   Begin VB.Timer timerDie 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   1560
      Top             =   2880
   End
   Begin VB.Timer timerBonus 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   4920
   End
   Begin VB.PictureBox mmLaser 
      Height          =   495
      Left            =   4200
      ScaleHeight     =   435
      ScaleWidth      =   3090
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.Timer timerImages 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Top             =   4920
   End
   Begin VB.Timer timerAsteroids 
      Interval        =   10000
      Left            =   720
      Top             =   2880
   End
   Begin VB.Timer timerLaser 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   3960
   End
   Begin VB.Timer timerHit 
      Interval        =   10
      Left            =   1200
      Top             =   3960
   End
   Begin VB.Timer timerBoom 
      Interval        =   250
      Left            =   1680
      Top             =   3960
   End
   Begin VB.Timer timerSpecial 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   2280
      Top             =   3960
   End
   Begin VB.Timer timerS_Lasers 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   2280
      Top             =   4440
   End
   Begin VB.PictureBox mmExplode 
      Height          =   495
      Left            =   4200
      ScaleHeight     =   435
      ScaleWidth      =   3090
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.Label lblHighScore 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "You Achieved the High Score"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   330
      Left            =   2040
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblTheEnd 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "The End"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1560
      Left            =   960
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   31
      Left            =   4560
      Picture         =   "Protector2.frx":C2EC2
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   30
      Left            =   4200
      Picture         =   "Protector2.frx":C33E4
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   29
      Left            =   3840
      Picture         =   "Protector2.frx":C3906
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   28
      Left            =   4560
      Picture         =   "Protector2.frx":C3E28
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   27
      Left            =   4200
      Picture         =   "Protector2.frx":C434A
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   26
      Left            =   3840
      Picture         =   "Protector2.frx":C486C
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   25
      Left            =   3480
      Picture         =   "Protector2.frx":C4D8E
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   24
      Left            =   3120
      Picture         =   "Protector2.frx":C52B0
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   23
      Left            =   2760
      Picture         =   "Protector2.frx":C57D2
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   22
      Left            =   3480
      Picture         =   "Protector2.frx":C5CF4
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   21
      Left            =   3120
      Picture         =   "Protector2.frx":C6216
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   20
      Left            =   2760
      Picture         =   "Protector2.frx":C6738
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   19
      Left            =   3840
      Picture         =   "Protector2.frx":C6C5A
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   18
      Left            =   3480
      Picture         =   "Protector2.frx":C717C
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   17
      Left            =   3120
      Picture         =   "Protector2.frx":C769E
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   16
      Left            =   2760
      Picture         =   "Protector2.frx":C7BC0
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   15
      Left            =   2760
      Picture         =   "Protector2.frx":C80E2
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   14
      Left            =   4560
      Picture         =   "Protector2.frx":C8604
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   13
      Left            =   4200
      Picture         =   "Protector2.frx":C8B26
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   12
      Left            =   3840
      Picture         =   "Protector2.frx":C9048
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   11
      Left            =   3480
      Picture         =   "Protector2.frx":C956A
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   10
      Left            =   3120
      Picture         =   "Protector2.frx":C9A8C
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   9
      Left            =   2760
      Picture         =   "Protector2.frx":C9FAE
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   8
      Left            =   4560
      Picture         =   "Protector2.frx":CA4D0
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   7
      Left            =   4200
      Picture         =   "Protector2.frx":CA9F2
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   6
      Left            =   3840
      Picture         =   "Protector2.frx":CAF14
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   5
      Left            =   3480
      Picture         =   "Protector2.frx":CB436
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   4
      Left            =   3120
      Picture         =   "Protector2.frx":CB958
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   3
      Left            =   2760
      Picture         =   "Protector2.frx":CBE7A
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   2
      Left            =   3840
      Picture         =   "Protector2.frx":CC39C
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   1
      Left            =   3480
      Picture         =   "Protector2.frx":CC8BE
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgExplosion 
      Height          =   450
      Index           =   0
      Left            =   3120
      Picture         =   "Protector2.frx":CCDE0
      Top             =   4920
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lblBonHull 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Hull:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   630
      Left            =   2040
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.Label lblBonus 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bonus:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   630
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.Image Turret_Y 
      Height          =   450
      Left            =   1680
      Picture         =   "Protector2.frx":CD302
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   5880
      Picture         =   "Protector2.frx":CD6E8
      Top             =   120
      Width           =   930
   End
   Begin VB.Image Image4 
      Height          =   225
      Left            =   120
      Picture         =   "Protector2.frx":CDAD6
      Top             =   120
      Width           =   735
   End
   Begin VB.Image imgSpecial 
      Height          =   270
      Left            =   1920
      Picture         =   "Protector2.frx":CDE9E
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
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   1
      Top             =   75
      Width           =   195
   End
   Begin VB.Label lblHull 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6960
      TabIndex        =   0
      Top             =   75
      Width           =   465
   End
   Begin VB.Image Asteroid_X 
      Height          =   315
      Left            =   720
      Picture         =   "Protector2.frx":CE3FC
      Top             =   3480
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Asteroid_Y 
      Height          =   315
      Left            =   1200
      Picture         =   "Protector2.frx":CE77B
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
      Picture         =   "Protector2.frx":CEB10
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image Turret 
      Height          =   450
      Index           =   1
      Left            =   4920
      Picture         =   "Protector2.frx":CEEF6
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   1
      Left            =   6840
      Picture         =   "Protector2.frx":CF2DC
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   2
      Left            =   7320
      Picture         =   "Protector2.frx":CF68E
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   5
      Left            =   5880
      Picture         =   "Protector2.frx":CFA40
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   7
      Left            =   5400
      Picture         =   "Protector2.frx":CFDF2
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   4
      Left            =   4440
      Picture         =   "Protector2.frx":D01A4
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   6
      Left            =   3960
      Picture         =   "Protector2.frx":D0556
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   0
      Left            =   120
      Picture         =   "Protector2.frx":D0908
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   8
      Left            =   3480
      Picture         =   "Protector2.frx":D0CBA
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   9
      Left            =   3000
      Picture         =   "Protector2.frx":D106C
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   3
      Left            =   6360
      Picture         =   "Protector2.frx":D141E
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   10
      Left            =   2040
      Picture         =   "Protector2.frx":D17D0
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   11
      Left            =   1560
      Picture         =   "Protector2.frx":D1B82
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   12
      Left            =   1080
      Picture         =   "Protector2.frx":D1F34
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S_Turret 
      Height          =   450
      Index           =   13
      Left            =   600
      Picture         =   "Protector2.frx":D22E6
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
      Picture         =   "Protector2.frx":D2698
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   62
      Left            =   6720
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D2CC2
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   61
      Left            =   6240
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D32EC
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   60
      Left            =   5760
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D3916
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   59
      Left            =   5280
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D3F40
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   58
      Left            =   4800
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D456A
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   57
      Left            =   4320
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D4B94
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   56
      Left            =   3840
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D51BE
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   55
      Left            =   3360
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D57E8
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   54
      Left            =   2880
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D5E12
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   53
      Left            =   2400
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D643C
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   52
      Left            =   1920
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D6A66
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   51
      Left            =   1440
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D7090
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   50
      Left            =   960
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D76BA
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   49
      Left            =   480
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D7CE4
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   48
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D830E
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   47
      Left            =   7200
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D8938
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   46
      Left            =   6720
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D8F62
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   45
      Left            =   6240
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D958C
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   44
      Left            =   5760
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":D9BB6
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   43
      Left            =   5280
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DA1E0
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   42
      Left            =   4800
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DA80A
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   41
      Left            =   4320
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DAE34
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   40
      Left            =   3840
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DB45E
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   39
      Left            =   3360
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DBA88
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   38
      Left            =   2880
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DC0B2
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   37
      Left            =   2400
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DC6DC
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   36
      Left            =   1920
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DCD06
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   35
      Left            =   1440
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DD330
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   34
      Left            =   960
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DD95A
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   33
      Left            =   480
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DDF84
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   32
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DE5AE
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   31
      Left            =   7200
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DEBD8
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   30
      Left            =   6720
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DF202
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   29
      Left            =   6240
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DF82C
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   28
      Left            =   5760
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":DFE56
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   27
      Left            =   5280
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E0480
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   26
      Left            =   4800
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E0AAA
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   25
      Left            =   4320
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E10D4
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   24
      Left            =   3840
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E16FE
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   23
      Left            =   3360
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E1D28
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   22
      Left            =   2880
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E2352
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   21
      Left            =   2400
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E297C
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   20
      Left            =   1920
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E2FA6
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   19
      Left            =   1440
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E35D0
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   18
      Left            =   960
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E3BFA
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   17
      Left            =   480
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E4224
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   16
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E484E
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   0
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E4E78
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   1
      Left            =   480
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E54A2
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   2
      Left            =   960
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E5ACC
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   3
      Left            =   1440
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E60F6
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   4
      Left            =   1920
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E6720
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   5
      Left            =   2400
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E6D4A
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   6
      Left            =   2880
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E7374
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   7
      Left            =   3360
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E799E
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   8
      Left            =   3840
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E7FC8
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   9
      Left            =   4320
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E85F2
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   10
      Left            =   4800
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E8C1C
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   11
      Left            =   5280
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E9246
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   12
      Left            =   5760
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E9870
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   13
      Left            =   6240
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":E9E9A
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   14
      Left            =   6720
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":EA4C4
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image Asteroid 
      Height          =   315
      Index           =   15
      Left            =   7200
      MousePointer    =   2  'Cross
      Picture         =   "Protector2.frx":EAAEE
      Top             =   1080
      Width           =   345
   End
End
Attribute VB_Name = "frmProtector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Action As Integer
Dim Shots As Integer
Dim Disap As Integer
Dim Size As Integer
Dim Pos As Integer
Dim Bonus As Integer
Dim Die As Integer
Dim Phase As Integer
Dim Boom As Integer
Dim TheHighScore As Integer

' Laser Fire

Private Sub Asteroid_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If timerBoom.Enabled = False Then
    Else:
            Line1.Visible = True
            Line2.Visible = True
            timerLaser.Enabled = True
    End If
    
    Asteroid(Index).Picture = Asteroid_X.Picture
    Line1.X1 = Asteroid(Index).Left + Asteroid(Index).Width / 2
    Line2.X1 = Asteroid(Index).Left + Asteroid(Index).Width / 2
    Line1.Y1 = Asteroid(Index).Top + Asteroid(Index).Height / 2
    Line2.Y1 = Asteroid(Index).Top + Asteroid(Index).Height / 2
End Sub

' End

Private Sub cmdEnd_Click()
    If lblScore.Caption > HIGHSCORE Then
        TheHighScore = 1
        HIGHSCORE = lblScore.Caption
        frmMain.lblHighScore.Caption = HIGHSCORE
    End If

    Unload Me
End Sub

' Mousestrokes

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

' Keystrokes

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If timerBoom.Enabled = False Then
    ElseIf lblHull.Caption <= 25 Then
        timerImages.Enabled = True
        imgSpecial.Visible = True
    ElseIf KeyCode = 32 Then
            timerAsteroids.Enabled = False
            timerBoom.Enabled = False
            timerHit.Enabled = False
            timerSpecial.Enabled = True
    End If
    
    If KeyCode = 48 Then
        lblHull.Caption = 0
    End If
    
    If KeyCode = 68 Then
        lblHull.Caption = 1
    End If
    
    If KeyCode = 83 Then
        lblScore.Caption = lblScore.Caption + 1000
    End If
End Sub

' Die

Private Sub timerDie_Timer()
    Select Case (Phase)
        Case 0:
                For t = 0 To 13
                    S_Turret(t).Top = S_Turret(t).Top - 120
                Next t
                    
                If S_Turret(Index).Top = 6480 Then
                    Phase = 1
                    timerDie.Interval = 25
                End If
        Case 1:
                If Action = 6 Then
                    For e = 0 To 31
                        imgExplosion(e).Visible = False
                    Next e
                    
                    timerDie.Interval = 1000
                    Phase = 2
                    Action = 1
                Else
                        For e = 0 To 31
                            imgExplosion(e).Visible = False
                        Next e
                    
                        For e = 0 To 31
                            Randomize
                            l = Int((7200 - 0 + 1) * Rnd + 0)
                                t = Int((6720 - 6240 + 1) * Rnd + 6240)
                            
                            imgExplosion(e).Visible = True
                            
                            imgExplosion(e).Left = l
                            imgExplosion(e).Top = t
                        Next e
                        
                    'mmExplode.Command = "open"
                    'mmExplode.Command = "play"
                    
                    Boom = 1
                    Action = Action + 1
                End If
        Case 2:
                Turret(0).Top = Turret(0).Top + 120
                Turret(1).Top = Turret(1).Top + 120
                
                For s = 0 To 13
                    S_Turret(s).Top = S_Turret(s).Top + 120
                Next s
                
                Action = Action + 1
                
                If Action = 5 Then
                    timerDie.Interval = 1000
                    Phase = 3
                End If
        Case 3:
                lblTheEnd.Visible = True
                cmdEnd.Visible = True
                
                If lblScore.Caption > HIGHSCORE Then
                    If lblHighScore.Visible = True Then
                        lblHighScore.Visible = False
                    Else:   lblHighScore.Visible = True
                    End If
                End If
    End Select
End Sub

' Hull Struck

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
            Exit For
        End If
    Next z
    
    If lblHull.Caption <= 0 Then
        timerAsteroids.Enabled = False
        timerHit.Enabled = False
        timerBonus.Enabled = False
        timerBoom.Enabled = False
        timerHull.Enabled = False
        timerImages.Enabled = False
        timerS_Lasers.Enabled = False
        timerLaser.Enabled = False
        
        Line1.Visible = False
        Line2.Visible = False
        
        timerDie.Enabled = True
    End If
End Sub

' Asteroid Struck

Private Sub timerBoom_Timer()
    For z = 0 To 63
    If Asteroid(z).Picture = Asteroid_X.Picture Then
        
        lblScore.Caption = lblScore.Caption + SCORE
        
        If lblScore.Caption Mod 25000 = 0 Then
            lblBonHull.Caption = "Hull:          50"
            lblBonHull.Visible = True
            timerHull.Enabled = True
            lblHull.Caption = lblHull.Caption + 50
        ElseIf lblScore.Caption Mod 15000 = 0 Then
            lblBonus.Caption = "Bonus:     " & Bonus
            lblBonus.Visible = True
            timerBonus.Enabled = True
            lblScore.Caption = lblScore.Caption + Bonus
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

' Special Fired

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
            'mmLaser.Command = "open"
            'mmLaser.Command = "play"
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
        
            timerAsteroids.Enabled = True
            timerBoom.Enabled = True
            timerHit.Enabled = True
            timerS_Lasers.Enabled = False
        End If
        
        For t = 0 To 13
            S_Turret(t).Top = S_Turret(t).Top + 120
        Next t
    End If
End Sub

' Asteroid Move

Private Sub timerasteroids_Timer()
    Randomize
    a = Int((15 - 0 + 1) * Rnd + 0)
    b = Int((31 - 16 + 1) * Rnd + 16)
    c = Int((47 - 32 + 1) * Rnd + 32)
    d = Int((63 - 48 + 1) * Rnd + 48)
    
    fall (a)
    fall (b)
    fall (c)
    fall (d)
End Sub

' Special Laser Animation

Private Sub timerSpecial_Timer()
    For t = 0 To 13
        S_Turret(t).Top = S_Turret(t).Top - 120
    Next t
        
    If S_Turret(Index).Top = 6480 Then
        timerSpecial.Enabled = False
        timerS_Lasers.Enabled = True
    End If
End Sub

' Point Bonus

Private Sub timerBonus_Timer()
    If lblBonus.Visible = True Then
        lblBonus.Visible = False
        timerBonus.Enabled = False
    End If
End Sub

' Hull Bonus

Private Sub timerHull_Timer()
    If lblBonHull.Visible = True Then
        lblBonHull.Visible = False
        timerHull.Enabled = False
    End If
End Sub

' Not Enough Hull

Private Sub timerImages_Timer()
    If imgSpecial.Visible = True Then
        imgSpecial.Visible = False
        timerImages.Enabled = False
    End If
End Sub

' Laser Visible

Private Sub timerLaser_Timer()
    Line1.Visible = False
    Line2.Visible = False
End Sub

Function fall(z) As Integer
    Asteroid(z).Top = Asteroid(z).Top + 100
End Function

'Private Sub mmExplode_Done(NotifyCode As Integer)
'    mmExplode.Command = "close"
'End Sub

'Private Sub mmLaser_Done(NotifyCode As Integer)
'    mmLaser.Command = "close"
'End Sub

Private Sub Form_Load()
    Me.Show
    
    timerAsteroids.Interval = SPEED
    
    lblHull = HULL
    
    Shots = -1
    Disap = -1
    Size = 1
    Pos = -480
    Bonus = 10 * SCORE
    Phase = 0
    Action = 1
    Boom = 1
    
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
    Load frmMain
End Sub
