VERSION 5.00
Begin VB.Form frmBreakout 
   Caption         =   "Breakout"
   ClientHeight    =   6300
   ClientLeft      =   1290
   ClientTop       =   1230
   ClientWidth     =   8385
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Breakout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6756.032
   ScaleMode       =   0  'User
   ScaleWidth      =   8525.055
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   4920
   End
   Begin VB.Timer tmrPreventReSize 
      Interval        =   1
      Left            =   240
      Top             =   3720
   End
   Begin VB.Timer tmrFrameSpeed 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1200
      Top             =   5760
   End
   Begin VB.PictureBox pixBackground 
      Height          =   6135
      Left            =   1080
      ScaleHeight     =   6000
      ScaleMode       =   0  'User
      ScaleWidth      =   7000
      TabIndex        =   0
      Top             =   120
      Width           =   7180
      Begin VB.VScrollBar vsbLevel 
         Height          =   255
         Left            =   3610
         TabIndex        =   67
         Top             =   1305
         Width           =   135
         Visible         =   0   'False
      End
      Begin VB.CommandButton cmdOKOptions 
         Caption         =   "OK"
         Height          =   375
         Left            =   2880
         TabIndex        =   63
         Top             =   4800
         Width           =   1335
         Visible         =   0   'False
      End
      Begin VB.OptionButton optHard 
         Caption         =   "Hard"
         Height          =   255
         Left            =   3480
         TabIndex        =   56
         Top             =   2400
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.OptionButton optMedium 
         Caption         =   "Medium"
         Height          =   255
         Left            =   3480
         TabIndex        =   55
         Top             =   2040
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.OptionButton optEasy 
         Caption         =   "Easy"
         Height          =   255
         Left            =   3480
         TabIndex        =   54
         Top             =   1680
         Width           =   855
         Visible         =   0   'False
      End
      Begin VB.CommandButton cmdOKCredits 
         Caption         =   "OK"
         Height          =   375
         Left            =   2880
         TabIndex        =   52
         Top             =   4920
         Width           =   1335
         Visible         =   0   'False
      End
      Begin VB.CommandButton cmdOKNewHighScoreName 
         Caption         =   "OK"
         Height          =   375
         Left            =   2880
         TabIndex        =   36
         Top             =   3960
         Width           =   1335
         Visible         =   0   'False
      End
      Begin VB.TextBox txtNewHighScoreName 
         Height          =   285
         Left            =   2040
         TabIndex        =   35
         Top             =   3480
         Width           =   2895
         Visible         =   0   'False
      End
      Begin VB.CommandButton cmdOKHighScore 
         Caption         =   "OK"
         Height          =   375
         Left            =   2880
         TabIndex        =   32
         Top             =   5160
         Width           =   1335
         Visible         =   0   'False
      End
      Begin VB.CommandButton cmdOKGameOver 
         Caption         =   "OK"
         Height          =   375
         Left            =   2880
         TabIndex        =   31
         Top             =   3120
         Width           =   1335
         Visible         =   0   'False
      End
      Begin VB.CommandButton cmdCredits 
         Caption         =   "Credits"
         Height          =   375
         Left            =   2880
         TabIndex        =   30
         Top             =   3600
         Width           =   1335
         Visible         =   0   'False
      End
      Begin VB.CommandButton cmdHighScore 
         Caption         =   "High Score"
         Height          =   375
         Left            =   2880
         TabIndex        =   29
         Top             =   3120
         Width           =   1335
         Visible         =   0   'False
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         Height          =   375
         Left            =   2880
         TabIndex        =   28
         Top             =   2640
         Width           =   1335
         Visible         =   0   'False
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   4080
         Width           =   1335
         Visible         =   0   'False
      End
      Begin VB.CommandButton cmdNewGame 
         Caption         =   "New Game"
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   2160
         Width           =   1323
         Visible         =   0   'False
      End
      Begin VB.Label lblLevel 
         Caption         =   "1"
         Height          =   255
         Left            =   3480
         TabIndex        =   68
         Top             =   1320
         Width           =   135
         Visible         =   0   'False
      End
      Begin VB.Label lblLevelText 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   195
         Left            =   2880
         TabIndex        =   66
         Top             =   1320
         Width           =   435
         Visible         =   0   'False
      End
      Begin VB.Label lblBrickColor 
         BackColor       =   &H8000000D&
         Height          =   495
         Left            =   3360
         TabIndex        =   62
         Top             =   3960
         Width           =   495
         Visible         =   0   'False
      End
      Begin VB.Label lblBrickColorText 
         AutoSize        =   -1  'True
         Caption         =   "Brick Color:"
         Height          =   195
         Left            =   2280
         TabIndex        =   61
         Top             =   4080
         Width           =   810
         Visible         =   0   'False
      End
      Begin VB.Label lblPaddleColor 
         BackColor       =   &H8000000D&
         Height          =   495
         Left            =   3360
         TabIndex        =   60
         Top             =   3360
         Width           =   495
         Visible         =   0   'False
      End
      Begin VB.Label lblPaddelColorText 
         AutoSize        =   -1  'True
         Caption         =   "Paddle Color:"
         Height          =   195
         Left            =   2280
         TabIndex        =   59
         Top             =   3480
         Width           =   945
         Visible         =   0   'False
      End
      Begin VB.Label lblBallColor 
         BackColor       =   &H8000000D&
         Height          =   495
         Left            =   3360
         TabIndex        =   58
         Top             =   2760
         Width           =   495
         Visible         =   0   'False
      End
      Begin VB.Label lblBallColorText 
         AutoSize        =   -1  'True
         Caption         =   "Ball Color:"
         Height          =   195
         Left            =   2280
         TabIndex        =   57
         Top             =   2880
         Width           =   705
         Visible         =   0   'False
      End
      Begin VB.Label lblDificultLevel 
         AutoSize        =   -1  'True
         Caption         =   "Dificult Level:"
         Height          =   255
         Left            =   2280
         TabIndex        =   53
         Top             =   2040
         Width           =   960
         Visible         =   0   'False
      End
      Begin VB.Label lblSpecialThanksToName 
         Alignment       =   2  'Center
         Caption         =   "Jarod Davis"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   51
         Top             =   4320
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblSpecialThanksToName 
         Alignment       =   2  'Center
         Caption         =   "Robin Swing"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   50
         Top             =   4080
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblSpecialThanksToName 
         Alignment       =   2  'Center
         Caption         =   "Gord Darkonian"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   49
         Top             =   3840
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblSpecialThanksToName 
         Alignment       =   2  'Center
         Caption         =   "Darksaber"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   48
         Top             =   3600
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblTestedByName 
         Alignment       =   2  'Center
         Caption         =   "Leon Ljunggren"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   44
         Top             =   2520
         Width           =   3735
         Visible         =   0   'False
      End
      Begin VB.Label lblSpecialThanksTo 
         Alignment       =   2  'Center
         Caption         =   "Special Thanks to:"
         Height          =   255
         Left            =   2520
         TabIndex        =   47
         Top             =   3360
         Width           =   2055
         Visible         =   0   'False
      End
      Begin VB.Label lblTestedByName 
         Alignment       =   2  'Center
         Caption         =   "Arso Slyth"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   46
         Top             =   3000
         Width           =   3735
         Visible         =   0   'False
      End
      Begin VB.Label lblTestedByName 
         Alignment       =   2  'Center
         Caption         =   "Gord Darkonian"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   45
         Top             =   2760
         Width           =   3735
         Visible         =   0   'False
      End
      Begin VB.Label lblTestedBy 
         Alignment       =   2  'Center
         Caption         =   "Tested by:"
         Height          =   255
         Left            =   1440
         TabIndex        =   43
         Top             =   2280
         Width           =   4215
         Visible         =   0   'False
      End
      Begin VB.Label lblDesignedByName 
         Alignment       =   2  'Center
         Caption         =   "Leon Ljunggren"
         Height          =   255
         Left            =   2280
         TabIndex        =   42
         Top             =   1920
         Width           =   2415
         Visible         =   0   'False
      End
      Begin VB.Label lblDesignedBy 
         Alignment       =   2  'Center
         Caption         =   "Designed by:"
         Height          =   255
         Left            =   2400
         TabIndex        =   41
         Top             =   1680
         Width           =   2295
         Visible         =   0   'False
      End
      Begin VB.Label lblProgramedByName 
         Alignment       =   2  'Center
         Caption         =   "Leon Ljunggren"
         Height          =   255
         Left            =   2280
         TabIndex        =   40
         Top             =   1320
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblProgramedBy 
         Alignment       =   2  'Center
         Caption         =   "Programed by:"
         Height          =   255
         Left            =   3000
         TabIndex        =   39
         Top             =   1080
         Width           =   1095
         Visible         =   0   'False
      End
      Begin VB.Label lblFinalScore 
         Caption         =   "You got a score of: "
         Height          =   255
         Left            =   2880
         TabIndex        =   38
         Top             =   3000
         Width           =   2055
         Visible         =   0   'False
      End
      Begin VB.Label lblEnterYourName 
         Caption         =   "Enter your name:"
         Height          =   255
         Left            =   3000
         TabIndex        =   34
         Top             =   3240
         Width           =   1815
         Visible         =   0   'False
      End
      Begin VB.Label lblNewHighScore 
         Caption         =   "Congratulations! You got a new High Score!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   33
         Top             =   2640
         Width           =   5535
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNameNr 
         Caption         =   "Name"
         Height          =   255
         Index           =   9
         Left            =   1320
         TabIndex        =   27
         Top             =   4800
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNameNr 
         Caption         =   "Name"
         Height          =   255
         Index           =   8
         Left            =   1320
         TabIndex        =   26
         Top             =   4440
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNameNr 
         Caption         =   "Name"
         Height          =   255
         Index           =   7
         Left            =   1320
         TabIndex        =   25
         Top             =   4080
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNameNr 
         Caption         =   "Name"
         Height          =   255
         Index           =   6
         Left            =   1320
         TabIndex        =   24
         Top             =   3720
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNameNr 
         Caption         =   "Name"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   23
         Top             =   3360
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNameNr 
         Caption         =   "Name"
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   22
         Top             =   3000
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNameNr 
         Caption         =   "Name"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   21
         Top             =   2640
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNameNr 
         Caption         =   "Name"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   20
         Top             =   2280
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNameNr 
         Caption         =   "Name"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   19
         Top             =   1920
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNameNr 
         Caption         =   "Name"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   18
         Top             =   1560
         Width           =   2535
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNr 
         Alignment       =   1  'Right Justify
         Caption         =   "Score"
         Height          =   255
         Index           =   9
         Left            =   4080
         TabIndex        =   17
         Top             =   4800
         Width           =   1815
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNr 
         Alignment       =   1  'Right Justify
         Caption         =   "Score"
         Height          =   255
         Index           =   8
         Left            =   4080
         TabIndex        =   16
         Top             =   4440
         Width           =   1815
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNr 
         Alignment       =   1  'Right Justify
         Caption         =   "Score"
         Height          =   255
         Index           =   7
         Left            =   4080
         TabIndex        =   15
         Top             =   4080
         Width           =   1815
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNr 
         Alignment       =   1  'Right Justify
         Caption         =   "Score"
         Height          =   255
         Index           =   6
         Left            =   4080
         TabIndex        =   14
         Top             =   3720
         Width           =   1815
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNr 
         Alignment       =   1  'Right Justify
         Caption         =   "Score"
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   13
         Top             =   3360
         Width           =   1815
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNr 
         Alignment       =   1  'Right Justify
         Caption         =   "Score"
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   12
         Top             =   3000
         Width           =   1815
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNr 
         Alignment       =   1  'Right Justify
         Caption         =   "Score"
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   11
         Top             =   2640
         Width           =   1815
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNr 
         Alignment       =   1  'Right Justify
         Caption         =   "Score"
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   10
         Top             =   2280
         Width           =   1815
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNr 
         Alignment       =   1  'Right Justify
         Caption         =   "Score"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   9
         Top             =   1920
         Width           =   1815
         Visible         =   0   'False
      End
      Begin VB.Label lblHighScoreNr 
         Alignment       =   1  'Right Justify
         Caption         =   "Score"
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   8
         Top             =   1560
         Width           =   1815
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   34
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   33
         Left            =   4800
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   32
         Left            =   3960
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   31
         Left            =   3120
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   30
         Left            =   2280
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   29
         Left            =   1440
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   28
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   27
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   26
         Left            =   4800
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   25
         Left            =   3960
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   24
         Left            =   3120
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   23
         Left            =   2280
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   22
         Left            =   1440
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   21
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   20
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   19
         Left            =   4800
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   18
         Left            =   3960
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   17
         Left            =   3120
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   16
         Left            =   2280
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   15
         Left            =   1440
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   14
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   13
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   12
         Left            =   4800
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   11
         Left            =   3960
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   10
         Left            =   3120
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   9
         Left            =   2280
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   8
         Left            =   1440
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   7
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   6
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   5
         Left            =   4800
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   4
         Left            =   3960
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   3120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   2280
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   1440
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Label lblGameOver 
         AutoSize        =   -1  'True
         Caption         =   "Game Over"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   5
         Top             =   2760
         Width           =   1605
         Visible         =   0   'False
      End
      Begin VB.Shape shpBall 
         FillStyle       =   0  'Solid
         Height          =   203
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   2280
         Width           =   204
         Visible         =   0   'False
      End
      Begin VB.Shape shpPaddle 
         FillStyle       =   0  'Solid
         Height          =   210
         Left            =   3000
         Shape           =   4  'Rounded Rectangle
         Top             =   5760
         Width           =   1425
         Visible         =   0   'False
      End
   End
   Begin VB.Label lblLevelStatus 
      Caption         =   "Level:"
      Height          =   255
      Left            =   0
      TabIndex        =   69
      Top             =   600
      Width           =   735
      Visible         =   0   'False
   End
   Begin VB.Label lblDif 
      Caption         =   "Dif:"
      Height          =   255
      Left            =   0
      TabIndex        =   65
      Top             =   240
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.Label lblFPS 
      Caption         =   "FPS:"
      Height          =   255
      Left            =   0
      TabIndex        =   64
      Top             =   2040
      Width           =   735
      Visible         =   0   'False
   End
   Begin VB.Label lblTime 
      Caption         =   "Time:"
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   1680
      Width           =   975
      Visible         =   0   'False
   End
   Begin VB.Label lblLives 
      Caption         =   "3"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   495
      Visible         =   0   'False
   End
   Begin VB.Label lblScore 
      Caption         =   "0"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   375
      Visible         =   0   'False
   End
   Begin VB.Label lblScoreText 
      Caption         =   "Score :"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.Label lblLivesText 
      Caption         =   "Lives:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   855
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmBreakout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'Ball variables
    Dim intBallSpeedX As Integer
    Dim intBallSpeedY As Integer
    Dim intBallPosX As Integer
    Dim intBallPosY As Integer
    
    'Paddle variables
    Dim intPaddlePosX As Integer
    Dim intPaddlePosY As Integer
    Dim intPaddleSpeed As Integer
    Dim intBallSize As Integer
    Dim intPaddleSize As Integer
    Dim intPaddleHeight As Integer
    
    'Brick variables
    Dim intBrickHeight As Integer
    Dim intBrickWidth As Integer
    Dim intBrickPosX(34) As Integer
    Dim intBrickPosY(34) As Integer
    
    'Score variables
    Dim intScore As Integer
    Dim intHighScore(9) As Integer
    Dim strHighScoreName(9) As String
    Dim bytPlace As Byte
    
    'Option variables
    Dim bytBallColor As Byte
    Dim bytPaddleColor As Byte
    Dim bytBrickColor As Byte
    Dim bytDificultLevel As Byte
    
    'Misc variables
    Dim intMaxX As Integer
    Dim intMaxY As Integer
    Dim bytLives As Byte
    Dim bytLevel As Byte
    Dim bytNumBricksDestroyed As Byte
    Dim intTime As Integer
    Dim intFPS As Integer
    Dim intFormHeight As Integer
    Dim intFormWidth As Integer
    
    'Constants
    Const cntMaxPaddleSpeed = 320
    Const cntMaxBallSpeed = 250
    Const cntNumberOfBricks = 35
    Const cntNrOfLevels = 5

Private Sub Form_Load()
    'Gets the height and width of the form for use in later frezing the size of it (should only be caled once, thus it's here)
    intFormHeight = frmBreakout.Height
    intFormWidth = frmBreakout.Width
    
    intMaxX = pixBackground.ScaleWidth
    intMaxY = pixBackground.ScaleHeight
    
    'Set which level to begin at (the default is always 1 but it can be chagned form game runing to game runing)
    bytLevel = 1
    
    'Get the options and high scores saved in the .shp file (Spearhawk Productions File)
    subReadOptions
    
    'Display the menue
    subMenu (True)
    
End Sub
    
Public Sub subMenu(blnShow As Boolean)
    If (blnShow = True) Then
        cmdNewGame.Visible = True
        cmdOptions.Visible = True
        cmdHighScore.Visible = True
        cmdCredits.Visible = True
        cmdQuit.Visible = True
    ElseIf (blnShow = False) Then
        cmdNewGame.Visible = False
        cmdOptions.Visible = False
        cmdHighScore.Visible = False
        cmdCredits.Visible = False
        cmdQuit.Visible = False
    End If
End Sub

'Show the options (where you can chose dificult levels, and colors on the ball, paddle and brick)
Private Sub cmdOptions_Click()
    
    'Hide the menu
    subMenu (False)
    
    'Read in the setings made the previus run
    If (bytDificultLevel = 100) Then
        optEasy.Value = True
        optMedium.Value = False
        optHard.Value = False
    ElseIf (bytDificultLevel = 150) Then
        optEasy.Value = False
        optMedium.Value = True
        optHard.Value = False
    Else
        optEasy.Value = False
        optMedium.Value = False
        optHard.Value = True
    End If
    
    'Set the level to it's default (1)
    vsbLevel.Value = 1
    
    lblBallColor.BackColor = QBColor(bytBallColor)
    lblPaddleColor.BackColor = QBColor(bytPaddleColor)
    lblBrickColor.BackColor = QBColor(bytBrickColor)
    
    'Make the option list visible
    lblBallColor.Visible = True
    lblBallColorText.Visible = True
    lblBrickColor.Visible = True
    lblBrickColorText.Visible = True
    lblPaddelColorText.Visible = True
    lblPaddleColor.Visible = True
    lblDificultLevel.Visible = True
    optEasy.Visible = True
    optMedium.Visible = True
    optHard.Visible = True
    cmdOKOptions.Visible = True
    lblLevel.Visible = True
    lblLevelText.Visible = True
    vsbLevel.Visible = True
  
    
End Sub

Private Sub cmdOKOptions_Click()
    
    'Set the options chosen
    If (optEasy.Value = True) Then
        bytDificultLevel = 100
    ElseIf (optMedium.Value = True) Then
       bytDificultLevel = 150
    ElseIf (optHard.Value = True) Then
        bytDificultLevel = 200
    End If
    
    'Hide the option list
    lblBallColor.Visible = False
    lblBallColorText.Visible = False
    lblBrickColor.Visible = False
    lblBrickColorText.Visible = False
    lblPaddelColorText.Visible = False
    lblPaddleColor.Visible = False
    lblDificultLevel.Visible = False
    optEasy.Visible = False
    optMedium.Visible = False
    optHard.Visible = False
    cmdOKOptions.Visible = False
    lblLevel.Visible = False
    lblLevelText.Visible = False
    vsbLevel.Visible = False
    
    'Display the menu
    subMenu (True)
    
End Sub

'Chage the levels
Private Sub vsbLevel_Change()
    
    lblLevel.Caption = vsbLevel.Value
    
    'Stop the value from overscreeding the nr of levels and from going belove it
    If (vsbLevel.Value < 1) Then
        vsbLevel.Value = 1
    ElseIf (vsbLevel.Value > cntNrOfLevels) Then
        vsbLevel.Value = cntNrOfLevels
    End If
    
    
    bytLevel = vsbLevel.Value
End Sub

Private Sub lblBallColor_Click()

    'Change the color of the ball
    bytBallColor = bytBallColor + 1
    
    'If colors run out, go back from the start
    If (bytBallColor > 15) Then
        bytBallColor = 0
        ElseIf (bytBallColor = 7) Then     'Prevent the user to choose the "invisible" color
        bytBallColor = 8
    End If
    
    'Redraw the lable in the new color
    lblBallColor.BackColor = QBColor(bytBallColor)
    
End Sub

Private Sub lblPaddleColor_Click()

    'Change the color of the paddle
    bytPaddleColor = bytPaddleColor + 1
    
    'If colors run out, go back from the start
    If (bytPaddleColor > 15) Then
        bytPaddleColor = 0
        ElseIf (bytPaddleColor = 7) Then     'Prevent the user to choose the "invisible" color
        bytPaddleColor = 8
    End If
    
    'Redraw the lable in the new color
    lblPaddleColor.BackColor = QBColor(bytPaddleColor)
    
End Sub

Private Sub lblBrickColor_Click()

    'Change the color of the paddle
    bytBrickColor = bytBrickColor + 1
    
    'If colors run out, go back from the start
    If (bytBrickColor > 15) Then
        bytBrickColor = 0
    ElseIf (bytBrickColor = 7) Then     'Prevent the user to choose the "invisible" color
        bytBrickColor = 8
    End If
    
    'Redraw the lable in the new color
    lblBrickColor.BackColor = QBColor(bytBrickColor)
    
End Sub

Private Sub optEasy_Click()
    optEasy.Value = True
End Sub

Private Sub optMedium_Click()
    optMedium.Value = True
End Sub

Private Sub optHard_Click()
    optHard.Value = True
End Sub

Private Sub cmdCredits_Click()
    
    'Hide the menu
    subMenu (False)
    
    'Display the credit list
    lblProgramedBy.Visible = True
    lblProgramedByName.Visible = True
    lblDesignedBy.Visible = True
    lblDesignedByName.Visible = True
    lblTestedBy.Visible = True
    lblTestedByName(0).Visible = True
    lblTestedByName(1).Visible = True
    lblTestedByName(2).Visible = True
    lblSpecialThanksTo.Visible = True
    lblSpecialThanksToName(0).Visible = True
    lblSpecialThanksToName(1).Visible = True
    lblSpecialThanksToName(2).Visible = True
    lblSpecialThanksToName(3).Visible = True
    
    'Display the ok button
    cmdOKCredits.Visible = True
    
    
End Sub

Private Sub cmdOKCredits_Click()

    'Hide the credit list
    lblProgramedBy.Visible = False
    lblProgramedByName.Visible = False
    lblDesignedBy.Visible = False
    lblDesignedByName.Visible = False
    lblTestedBy.Visible = False
    lblTestedByName(0).Visible = False
    lblTestedByName(1).Visible = False
    lblTestedByName(2).Visible = False
    lblSpecialThanksTo.Visible = False
    lblSpecialThanksToName(0).Visible = False
    lblSpecialThanksToName(1).Visible = False
    lblSpecialThanksToName(2).Visible = False
    lblSpecialThanksToName(3).Visible = False
    
    'Hide the ok button
    cmdOKCredits.Visible = False
    
    'Display the menu
    subMenu (True)
    
End Sub

'Section: Take care of all pre-game settings
Public Sub subInitalize()

    'The ball's settings
    intBallSpeedX = 0     'Starts with a X speed of 0 twip per frame (it's going straight down)
    intBallSpeedY = bytDificultLevel     'Starts with a Y speed of the chosen dificult (50, 75 or 100 twip per frame)
    intBallSize = shpPaddle.Height
    intBallPosY = intMaxY / 2 'Starts in the middle of the screen
    intBallPosX = intMaxX / 2 'Starts in the middle of the screen
    shpBall.FillColor = QBColor(bytBallColor)
    
    'The paddle's settings
    intPaddleSpeed = 0         'Starts with zero speed
    intPaddleSize = shpPaddle.Width
    intPaddleHeight = shpPaddle.Height
    intPaddlePosX = (intMaxX / 2) - (intPaddleSize / 2)  'Starts in the middle of the screen
    intPaddlePosY = intMaxY - intPaddleHeight            'The Paddle starts at the button of the play yard
    shpPaddle.Top = intPaddlePosY                        'Set the paddles intBallPosY, no need to change that later on is no need to have it in the time loop
    shpPaddle.FillColor = QBColor(bytPaddleColor)
    
    'The bricks settings
    intBrickHeight = shpBrick(0).Height
    intBrickWidth = shpBrick(0).Width
    bytNumBricksDestroyed = 0
    
    'Set all game elemets to be viisible
    For i = 0 To (cntNumberOfBricks - 1)
        shpBrick(i).Visible = True
        shpBrick(i).FillColor = QBColor(bytBrickColor)
    Next i
    
    shpBall.Visible = True
    shpPaddle.Visible = True
    lblLives.Visible = True
    lblLivesText.Visible = True
    lblScore.Visible = True
    lblScoreText.Visible = True
    lblTime.Visible = True
    lblFPS.Visible = True
    lblLevelStatus.Visible = True
    
    'Misc settings
    intScore = 1
    bytLives = 3
    intTime = 0
    
    'Load level (in this case level 1)
    subLoadLevel
    
    'Place the bricks on thier correct cordinates
    For i = 0 To (cntNumberOfBricks - 1)
        shpBrick(i).Left = intBrickPosX(i)
        shpBrick(i).Top = intBrickPosY(i)
    Next i
    
    'Display the current dificult level
    If (bytDificultLevel = 100) Then
        lblDif.Caption = "Dif: Easy"
    ElseIf (bytDificultLevel = 150) Then
        lblDif.Caption = "Dif: Medium"
    ElseIf (bytDificultLevel = 200) Then
        lblDif.Caption = "Dif: Hard"
    End If
    lblDif.Visible = True
    
End Sub


Private Sub cmdNewGame_Click()
    'Hide the menue
    subMenu (False)
    
    'Resets the game, making it ready to begin
    subInitalize
    
    '(Re)Start the game loop
    tmrFrameSpeed.Enabled = True
    
    '(Re)Star the time
    tmrTime.Enabled = True
    
End Sub

'Section: The main game loop
Private Sub tmrFrameSpeed_Timer()
    
    intFPS = intFPS + 1
    
    'Update the position of the paddle and ball
    intBallPosX = intBallPosX + intBallSpeedX
    intBallPosY = intBallPosY + intBallSpeedY
    intPaddlePosX = intPaddlePosX + intPaddleSpeed
    
    'Section: Game Logic
    'The balls bouncing algoritm
    If ((intBallPosX > intMaxX - intBallSize) And (intBallSpeedX > 0)) Then 'Check so that the x speed is positive, so the ball can't get stuck
        intBallSpeedX = -intBallSpeedX
    ElseIf ((intBallPosX <= 0) And (intBallSpeedX < 0)) Then    'Check so that the x speed is negative, so the ball can't get stuck
        intBallSpeedX = -intBallSpeedX
        
    'Does the ball hit the paddle?
    ElseIf ((intBallPosY > intPaddlePosY - intBallSize) And ((intBallPosX > intPaddlePosX - intBallSize) And (intBallPosX < intPaddlePosX + intPaddleSize))) Then
        
        'Make the ball bounce in diffrent angles depending on where on the paddle it hits
        If (intBallPosX < (intPaddlePosX + (intPaddleSize / 5))) Then
            intBallSpeedX = intBallSpeedX - 35
        ElseIf (intBallPosX < (intPaddlePosX + (2 * (intPaddleSize / 5)))) Then
            intBallSpeedX = intBallSpeedX - 17
        ElseIf (intBallPosX < (intPaddlePosX + (3 * (intPaddleSize / 5)))) Then
            intBallSpeedX = intBallSpeedX
        ElseIf (intBallPosX < (intPaddlePosX + (4 * (intPaddleSize / 5)))) Then
            intBallSpeedX = intBallSpeedX + 17
        ElseIf (intBallPosX < (intPaddlePosX + intPaddleSize)) Then
            intBallSpeedX = intBallSpeedX + 35
        End If
        
        'Limit the ball to its max speed
        If (intBallSpeedX >= cntMaxBallSpeed) Then
            intBallSpeedX = cntMaxBallSpeed
        ElseIf (intBallSpeedX <= -cntMaxBallSpeed) Then
            intBallSpeedX = -cntMaxBallSpeed
        End If
        
        'Check so that the ball is going down thorwards the paddle, if so bounce it (too prevent cheating)
        If (intBallSpeedY > 0) Then
            intBallSpeedY = -intBallSpeedY
        End If
        
        intScore = intScore + 1     'The player gains one pts for hitting the ball
    'If it don't then do it miss? If soo reduce a life
    ElseIf (intBallPosY > intMaxY - intBallSize) Then
        bytLives = bytLives - 1     'Reduce a life
        
        'Reset everything
        intBallSpeedX = 0    'Starts with a X speed of 0 twip per frame
        intBallSpeedY = bytDificultLevel     'Starts with a Y speed of 75 twip per frame
        intBallPosY = intMaxY / 2 'Starts in the middle of the screen
        intBallPosX = intMaxX / 2 'Starts in the middle of the screen
        intPaddleSpeed = 0         'Starts with zero speed
        intPaddlePosX = (intMaxX / 2) - (intPaddleSize / 2)  'Starts in the middle of the screen
    
        If (bytLives = 0) Then
            subEndGame ("Game over!")    'End the game when the players life hits zero
        End If
        
    ElseIf ((intBallPosY < 0) And (intBallSpeedY < 0)) Then 'Check so that the y speed is negative so the ball can't get stuck
        intBallSpeedY = -intBallSpeedY
    End If
    
    'Collision detection for the bricks
    For i = 0 To (cntNumberOfBricks - 1)
        'If the brick is hit, "delete" it and give the player his pts
        If ((intBallPosY >= shpBrick(i).Top - intBallSize) And (intBallPosY <= (shpBrick(i).Top + intBrickHeight)) And (intBallPosX >= (shpBrick(i).Left - intBallSize)) And (intBallPosX <= (shpBrick(i).Left + intBrickWidth)) And (shpBrick(i).Visible = True)) Then
            intBallSpeedY = -intBallSpeedY
            shpBrick(i).Visible = False
            bytNumBricksDestroyed = bytNumBricksDestroyed + 1
            intScore = intScore + 100
        End If
    Next i
    
    'If all bricks have been "destroyed" then load the next level
    If (bytNumBricksDestroyed = cntNumberOfBricks) Then
        'Load the next level
        subNextLevel
    End If
    
    'Prevent the paddle from exceding it's max speed
    If (intPaddleSpeed >= cntMaxPaddleSpeed) Then
        intPaddleSpeed = cntMaxPaddleSpeed
    ElseIf (intPaddleSpeed <= -cntMaxPaddleSpeed) Then
        intPaddleSpeed = -cntMaxPaddleSpeed
    End If
    
    'Prevent the paddle to get out of screen
    If (intPaddlePosX > intMaxX - intPaddleSize) Then
        intPaddleSpeed = 0
        intPaddlePosX = intMaxX - intPaddleSize
    ElseIf (intPaddlePosX < 0) Then
        intPaddleSpeed = 0
        intPaddlePosX = 1
    End If

    'Section: draw everything
    shpPaddle.Left = intPaddlePosX
    shpBall.Left = intBallPosX
    shpBall.Top = intBallPosY
    lblScore.Caption = intScore
    lblLives.Caption = bytLives
    lblLevelStatus = "Level: " & bytLevel
        
End Sub

'Section: Get player input data
Public Sub pixBackground_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyRight Then
        intPaddleSpeed = intPaddleSpeed + 40
    ElseIf KeyCode = vbKeyLeft Then
        intPaddleSpeed = intPaddleSpeed - 40
    ElseIf KeyCode = 27 Then
        If (tmrFrameSpeed.Enabled = True) Then  'Check if the game is runing
            subEndGame ("Game Over!")    'If Escape is hit, end the game
        End If
    End If
End Sub

Public Sub subNextLevel()
    
    bytLevel = bytLevel + 1
    intScore = intScore + 1000  'You get 1K pts for finishng a level
    
    'If all levels have been completed end the game
    If (bytLevel > cntNrOfLevels) Then
        subEndGame ("Congratulations! You finished the game!")
    Else
        'Load the new level
        subLoadLevel
        
        'Put the bricks in thier new position
        For i = 0 To (cntNumberOfBricks - 1)
            shpBrick(i).Left = intBrickPosX(i)
            shpBrick(i).Top = intBrickPosY(i)
        Next i
        
        'Reset everything
        intBallSpeedX = 0    'Starts with a X speed of 0 twip per frame
        intBallSpeedY = bytDificultLevel     'Starts with a Y speed of 75 twip per frame
        intBallPosY = intMaxY / 2 'Starts in the middle of the screen
        intBallPosX = intMaxX / 2 'Starts in the middle of the screen
        intPaddleSpeed = 0         'Starts with zero speed
        intPaddlePosX = (intMaxX / 2) - (intPaddleSize / 2)  'Starts in the middle of the screen
        bytNumBricksDestroyed = 0     'Zero bricks have been destroyed in the new level
        
        'Make the bricks visible once more
        For i = 0 To (cntNumberOfBricks - 1)
            shpBrick(i).Visible = True
        Next i
        
    End If
    
End Sub

Public Sub subEndGame(strEndingGame As String)
    tmrFrameSpeed.Enabled = False   'Disable the game loop
    tmrTime.Enabled = False         'Disable the time counter
    lblGameOver.Caption = strEndingGame     'Sets the game ending message to the content on the string
    lblGameOver.Left = (intMaxX / 2) - (lblGameOver.Width / 2)  'Place the message in the middle of the window
    lblGameOver.Visible = True      'Dislay the message Game Over
    cmdOKGameOver.Visible = True
    
    'The score is calcualted byt how many life you have left and how long time it have taken to complete it
    intScore = ((intScore * (bytLives + 1)) / ((intTime / 100) + 1))
    
    'Reset the levels so it starts on the first level
    bytLevel = 1
    
    'If not all the bricks have been destroyed when the user quits the game,
    'devide the score with the remaining bricks + 1 (e.g 1 brick remains, devide by 2)
    If (bytNumBricksDestroyed <> cntNumberOfBricks) Then
        intScore = intScore / ((cntNumberOfBricks - bytNumBricksDestroyed) + 1)
    End If
    
    'Hide all game elements (paddle, ball, bricks, lives, scores etc)
    For i = 0 To (cntNumberOfBricks - 1)
        shpBrick(i).Visible = False
    Next i
    
    shpBall.Visible = False
    shpPaddle.Visible = False
    lblLives.Visible = False
    lblLivesText.Visible = False
    lblScore.Visible = False
    lblScoreText.Visible = False
    lblTime.Visible = False
    lblDif.Visible = False
    lblFPS.Visible = False
    lblLevelStatus.Visible = False
    
End Sub

Private Sub cmdOKGameOver_Click()
    lblGameOver.Visible = False     'Hidds the message in the lable Game Over
    cmdOKGameOver.Visible = False   'Hidd the OK button
    
    'If there's a new high score get the players name for the high score list, otherwise show the menue
    If (fncShortHighScore(intScore) = True) Then
        subNewHighScoreName
    Else
        subMenu (True)     'Cals the menu and makes visible
    End If
    
End Sub

'The High Score function
Public Function fncShortHighScore(intHS As Integer)
    
    Dim blnNewHighScore
    Dim intHighScoreTemp(1) As Integer
    
    'Check if a high score have been atained
    If (intHS > intHighScore(9)) Then
            
        'If the player got a new high score, put it on its place, then move every other score under it down a step
        For i = 0 To 9
            If (intHS > intHighScore(i)) Then
                intHighScoreTemp(0) = intHighScore(i)
                intHighScore(i) = intHS
                intHS = 0
                bytPlace = i    'Used to know on which place to set the name
                
                For j = i To 8
                    intHighScoreTemp(1) = intHighScore(j + 1)
                    intHighScore(j + 1) = intHighScoreTemp(0)
                    intHighScoreTemp(0) = intHighScoreTemp(1)
                Next j
                
            End If
        Next i
        
        blnNewHighScore = True
        
    Else
        blnNewHighScore = False
    End If
    
    fncShortHighScore = blnNewHighScore
    
End Function

'Get the name of the new high scorer
Public Sub subNewHighScoreName()

    lblEnterYourName.Visible = True
    lblNewHighScore.Visible = True
    txtNewHighScoreName.Visible = True
    cmdOKNewHighScoreName.Visible = True
    lblFinalScore.Visible = True
    
    lblFinalScore.Caption = "You got a score of: " & intHighScore(bytPlace)
    
End Sub

Private Sub cmdOKNewHighScoreName_Click()

    Dim strNewName As String
    Dim strHighScoreNameTemp(1) As String
    
    strNewName = txtNewHighScoreName.Text
    txtNewHighScoreName.Text = ""   'Reset the text fild for the next HS
    
    'Put the new high score in its right position and put every other on step below
    'where they were before
    strHighScoreNameTemp(0) = strHighScoreName(bytPlace)
    strHighScoreName(bytPlace) = strNewName

    For j = bytPlace To 8
        strHighScoreNameTemp(1) = strHighScoreName(j + 1)
        strHighScoreName(j + 1) = strHighScoreNameTemp(0)
        strHighScoreNameTemp(0) = strHighScoreNameTemp(1)
    Next j
   
    
    lblEnterYourName.Visible = False
    lblNewHighScore.Visible = False
    txtNewHighScoreName.Visible = False
    cmdOKNewHighScoreName.Visible = False
    lblFinalScore.Visible = False
    
    'When the name have been entered corectly show the high score list
    subShowHighScore (True)
    
End Sub

'Display/Hide the high score list
Public Sub subShowHighScore(blnShowHS As Boolean)
    
    If (blnShowHS = True) Then
        For i = 0 To 9
            lblHighScoreNr(i).Caption = intHighScore(i)
            lblHighScoreNr(i).Visible = True
            
            lblHighScoreNameNr(i).Caption = strHighScoreName(i)
            lblHighScoreNameNr(i).Visible = True
            
            cmdOKHighScore.Visible = True
        Next i
    Else
        For i = 0 To 9
            lblHighScoreNr(i).Visible = False
            lblHighScoreNameNr(i).Visible = False
            cmdOKHighScore.Visible = False
        Next i
    End If
    
End Sub

Private Sub cmdOKHighScore_Click()
    
    'Hide the high score list
    subShowHighScore (False)
    
    'Show the menu
    subMenu (True)
    
End Sub

Private Sub cmdHighScore_Click()
    
    'Hide the menu
    subMenu (False)
    
    'Show the high score list
    subShowHighScore (True)
    
End Sub

Private Sub tmrPreventReSize_Timer()
    'Prevent the user from re-sizeing the window
    frmBreakout.Height = intFormHeight
    frmBreakout.Width = intFormWidth
End Sub

Public Sub subLoadLevel()

    'Decide which level file to open (.shl)
    If (bytLevel = 1) Then
        Open "level1.shl" For Input As #2
    ElseIf (bytLevel = 2) Then
        Open "level2.shl" For Input As #2
    ElseIf (bytLevel = 3) Then
        Open "level3.shl" For Input As #2
    ElseIf (bytLevel = 4) Then
        Open "level4.shl" For Input As #2
    ElseIf (bytLevel = 5) Then
        Open "level5.shl" For Input As #2
    End If
    
    'Read what's in the level file
    For i = 0 To 34
        Input #2, intBrickPosX(i)
        Input #2, intBrickPosY(i)
    Next i
    
    Close #2
    
End Sub

'The Loads options and High Scores from the .shp file (Spearhawk Productions File)
Public Sub subReadOptions()
    Open "options.shp" For Input As #1
    
    For i = 0 To 9
        Input #1, intHighScore(i)
        Input #1, strHighScoreName(i)
    Next i
    
    Input #1, bytBallColor
    Input #1, bytPaddleColor
    Input #1, bytBrickColor
    Input #1, bytDificultLevel
    
    Close #1
    
End Sub

'The Saves options and High Scores to the .shp file (Spearhawk Productions File)
Public Sub subSaveOptions()
    Open "options.shp" For Output As #1
    
    
    For i = 0 To 9
        Print #1, intHighScore(i)
        Print #1, strHighScoreName(i)
    Next i
    
    Print #1, bytBallColor
    Print #1, bytPaddleColor
    Print #1, bytBrickColor
    Print #1, bytDificultLevel
        
    Close #1
    
End Sub

'Quit the game
Private Sub cmdQuit_Click()
    
    'Save options and high scores to the .shp file
    subSaveOptions
    
    End
    
End Sub

'This function keeps track of how long a game is played (which is inportant in calculution of scores)
Private Sub tmrTime_Timer()
    intTime = intTime + 1
    lblTime = "Time: " & intTime
    lblFPS.Caption = "FPS: " & intFPS
    intFPS = 0
End Sub

