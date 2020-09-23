VERSION 5.00
Begin VB.Form frmBreakoutLevelEditor 
   Caption         =   "Breakout Level Editor"
   ClientHeight    =   6315
   ClientLeft      =   1290
   ClientTop       =   1230
   ClientWidth     =   8655
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6772.118
   ScaleMode       =   0  'User
   ScaleWidth      =   8799.564
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   1335
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdChageLevel 
      Caption         =   "Change Level"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   1335
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
      Visible         =   0   'False
   End
   Begin VB.PictureBox pixBackground 
      Height          =   6135
      Left            =   1440
      ScaleHeight     =   6000
      ScaleMode       =   0  'User
      ScaleWidth      =   7000
      TabIndex        =   0
      Top             =   120
      Width           =   7180
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lblWhichLevel 
         AutoSize        =   -1  'True
         Caption         =   "Which level do you want to edit (1-5)?"
         Height          =   195
         Left            =   2280
         TabIndex        =   1
         Top             =   2520
         Width           =   2685
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   34
         Left            =   4080
         Shape           =   4  'Rounded Rectangle
         Top             =   2640
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   33
         Left            =   4080
         Shape           =   4  'Rounded Rectangle
         Top             =   3120
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   32
         Left            =   720
         Shape           =   4  'Rounded Rectangle
         Top             =   1320
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   31
         Left            =   6360
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   30
         Left            =   3240
         Shape           =   4  'Rounded Rectangle
         Top             =   2760
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   29
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   2880
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   28
         Left            =   1440
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   27
         Left            =   4080
         Shape           =   4  'Rounded Rectangle
         Top             =   2400
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   26
         Left            =   4200
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   25
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   2400
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   24
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   1320
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   23
         Left            =   2160
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   22
         Left            =   1440
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   21
         Left            =   720
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   20
         Left            =   4920
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   19
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   3120
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   18
         Left            =   2760
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   17
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   16
         Left            =   3600
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   15
         Left            =   6360
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   14
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   13
         Left            =   2760
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   12
         Left            =   4920
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   11
         Left            =   2760
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   10
         Left            =   4200
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   9
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   2640
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   8
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   7
         Left            =   4080
         Shape           =   4  'Rounded Rectangle
         Top             =   2880
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   6
         Left            =   2160
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   5
         Left            =   3600
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   4
         Left            =   2760
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   2760
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   3600
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   3600
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Shape shpBrick 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   3600
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   735
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmBreakoutLevelEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'Moving variables
    Dim mouseOffsetX As Long, mouseOffsetY As Long

    'Brick variables
    Dim intBrickPosX(34) As Integer
    Dim intBrickPosY(34) As Integer
    
    'Constants
    Const cntNumberOfBricks = 35
    Const cntNrOfLevels = 5
    
    Dim bytLevel As Byte
    
    
Private Sub cmdChageLevel_Click()
    
    'Hide the bricks
    For i = 0 To (cntNumberOfBricks - 1)
        shpBrick(i).Visible = False
    Next i
    
    'Hide the menu
    cmdSave.Visible = False
    cmdReset.Visible = False
    cmdChageLevel.Visible = False
    cmdQuit.Visible = False
    
    'Display the choice object
    lblWhichLevel.Visible = True
    txtLevel.Visible = True
    cmdOK.Visible = True
    
End Sub

Private Sub cmdReset_Click()
    'Read the cordinates for the bricks so you can se how it looks
    subLoadLevel
    
    For i = 0 To (cntNumberOfBricks - 1)
        shpBrick(i).Left = intBrickPosX(i)
        shpBrick(i).Top = intBrickPosY(i)
    Next i
End Sub

Private Sub cmdOK_Click()
    
    bytLevel = Val(txtLevel.Text)
    txtLevel.Text = ""
    
    'If the level excists loade it, otherwise give a warning
    If ((bytLevel >= 1) And (bytLevel <= cntNrOfLevels)) Then
        
        lblWhichLevel.Visible = False
        txtLevel.Visible = False
        cmdOK.Visible = False
        
        For i = 0 To (cntNumberOfBricks - 1)
            shpBrick(i).Visible = True
        Next i
        
        'Load the level that's chosen
        'Read the cordinates for the bricks so you can se how it looks
        subLoadLevel
        
        For i = 0 To (cntNumberOfBricks - 1)
            shpBrick(i).Left = intBrickPosX(i)
            shpBrick(i).Top = intBrickPosY(i)
        Next i
        
        'Display the menue
        cmdSave.Visible = True
        cmdReset.Visible = True
        cmdChageLevel.Visible = True
        cmdQuit.Visible = True
        
    Else
        MsgBox ("Pleas chose a level betwen 1 and " & cntNrOfLevels)
    End If
    
End Sub

Private Sub cmdSave_Click()

    For i = 0 To (cntNumberOfBricks - 1)
        intBrickPosX(i) = shpBrick(i).Left
        intBrickPosY(i) = shpBrick(i).Top
    Next i
    
    'Write the curent cordinates of the bricks to the level file
    subSaveLevel
    
End Sub

'This loade the corect level form the diffrent level files (.shl)
Public Sub subLoadLevel()
    'Decide which level file to open
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
    
    'Read what's in the file
    For i = 0 To 34
        Input #2, intBrickPosX(i)
        Input #2, intBrickPosY(i)
    Next i
    
    Close #2
    
End Sub

'This saves the level to the corect level file (.shl)
Public Sub subSaveLevel()

    If (bytLevel = 1) Then
        Open "level1.shl" For Output As #2
    ElseIf (bytLevel = 2) Then
        Open "level2.shl" For Output As #2
    ElseIf (bytLevel = 3) Then
        Open "level3.shl" For Output As #2
    ElseIf (bytLevel = 4) Then
        Open "level4.shl" For Output As #2
    ElseIf (bytLevel = 5) Then
        Open "level5.shl" For Output As #2
    End If
    
    'Print the codrinates of the bricks to the .shl file (Spearhawk Levels)
    For i = 0 To 34
        Print #2, intBrickPosX(i)
        Print #2, intBrickPosY(i)
    Next i
        
    Close #2
    
End Sub

'The Folowing three functions takes care of the moving of the bricks
Private Sub pixBackground_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    For i = 0 To (cntNumberOfBricks - 1)
    
        'If the mouse putton is pressed while the pointer is on one of the brick, tag
        'that one as DRAG so the program knows which to move
        If X > shpBrick(i).Left And X < shpBrick(i).Left + shpBrick(i).Width And _
            Y > shpBrick(i).Top And Y < shpBrick(i).Top + shpBrick(i).Height Then
            
            mouseOffsetX = X - shpBrick(i).Left
            mouseOffsetY = Y - shpBrick(i).Top
            shpBrick(i).Tag = "DRAG"
        End If
    Next i
    
End Sub

Private Sub pixBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    For i = 0 To (cntNumberOfBricks - 1)
    
        'If the brick have the tag DRAG then move it as you move the mouse
        If shpBrick(i).Tag = "DRAG" Then
            shpBrick(i).Move X - mouseOffsetX, Y - mouseOffsetY
            
            'Make it so that two bricks can't be moved on the same time (if they overlap each other)
            For j = (i + 1) To (cntNumberOfBricks - 1)
                shpBrick(j).Tag = ""
            Next j
            
        End If
    Next i
    
End Sub

Private Sub pixBackground_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    For i = 0 To (cntNumberOfBricks - 1)
        'When the mouse button is released erase the tag so that the brick won't continue to move after the mouse
        shpBrick(i).Tag = ""
    Next i
    
End Sub

'Quit the editor
Private Sub cmdQuit_Click()
    End
End Sub
