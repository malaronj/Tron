VERSION 5.00
Begin VB.Form frmTron 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tron"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrGamePlay2 
      Interval        =   25
      Left            =   600
      Top             =   6600
   End
   Begin VB.Timer tmrSpeed 
      Interval        =   3000
      Left            =   1560
      Top             =   6600
   End
   Begin VB.PictureBox picSpeed 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   590
      Left            =   9120
      ScaleHeight     =   525
      ScaleWidth      =   675
      TabIndex        =   23
      Top             =   6960
      Width           =   730
   End
   Begin VB.Frame fraStages 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3235
      Left            =   6840
      TabIndex        =   18
      Top             =   3480
      Width           =   3015
      Begin VB.OptionButton optStage3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Multiplayer w/ Speed Increase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2280
         Width           =   2415
      End
      Begin VB.OptionButton optStage2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Single Player w/ Speed Increase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1320
         Width           =   2415
      End
      Begin VB.OptionButton optStage1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Single Player"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame fraControls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   6840
      TabIndex        =   5
      Top             =   240
      Width           =   3015
      Begin VB.Label lblRules 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "e - End Game"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   920
         Width           =   2775
      End
      Begin VB.Label lblRules2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "r - Reset Game"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   625
         Width           =   2775
      End
      Begin VB.Label lblPlayer2Controls4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "l - RIGHT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lblPlayer2Controls3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "j - LEFT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblPlayer2Controls2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "k - DOWN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblPlayer2Controls 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "i - UP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblPlayer1Controls4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "d - RIGHT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lblPlayer1Controls3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "a - LEFT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblPlayer2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Player 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblPlayer1Controls2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "s - DOWN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblPlayer1Controls 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "w - UP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblPlayer1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Player 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblRules1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "p/o - Pause/Unpause"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Timer tmrScores 
      Interval        =   40
      Left            =   1080
      Top             =   6600
   End
   Begin VB.PictureBox picScore2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   590
      Left            =   6480
      ScaleHeight     =   525
      ScaleWidth      =   945
      TabIndex        =   4
      Top             =   6960
      Width           =   1000
   End
   Begin VB.PictureBox picScore1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   590
      Left            =   2760
      ScaleHeight     =   525
      ScaleWidth      =   945
      TabIndex        =   2
      Top             =   6960
      Width           =   1000
   End
   Begin VB.Timer tmrGamePlay 
      Interval        =   25
      Left            =   120
      Top             =   6600
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   6500
      Left            =   120
      ScaleHeight     =   429
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   429
      TabIndex        =   0
      Top             =   240
      Width           =   6500
      Begin VB.Line Line6 
         X1              =   0
         X2              =   432
         Y1              =   428
         Y2              =   428
      End
      Begin VB.Line Line5 
         X1              =   428
         X2              =   428
         Y1              =   0
         Y2              =   432
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   432
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   432
      End
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   24
      Top             =   7005
      Width           =   1335
   End
   Begin VB.Label lblPlayer2Score 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Player 2 Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   7005
      Width           =   2295
   End
   Begin VB.Label lblPlayer1Score 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Player 1 Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      TabIndex        =   1
      Top             =   7000
      Width           =   2295
   End
End
Attribute VB_Name = "frmTron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Malaron Jeyakumar
'Date: May 16, 2022
'Purpose: Sandbox - Playing the game of tron, with 1-2 players.
'===========================================================================================================
Option Explicit

'Global Variables
Dim strUser1Direction As String
Dim intUser1X As Integer
Dim intUser1Y As Integer
Dim strUser2Direction As String
Dim intUser2X As Integer
Dim intUser2Y As Integer
Dim intPlayer1Score As Integer
Dim intPlayer2Score As Integer
Dim intSpeed As Integer

Private Sub Form_Load()
picOutput.Cls

'Initialize
strUser1Direction = 0
intUser1X = 0
intUser1Y = 0
strUser2Direction = 0
intUser2X = 0
intUser2Y = 0
intPlayer1Score = 0
intPlayer2Score = 0
intSpeed = 0

MsgBox "Choose Stage to Start the Game"
End Sub

Private Sub Form_Activate()
picSpeed.Print intSpeed
End Sub

Private Sub picOutput_KeyPress(KeyAscii As Integer)
'Player 1
If KeyAscii = 119 Then 'Key w
    strUser1Direction = "North"
ElseIf KeyAscii = 115 Then 'Key s
    strUser1Direction = "South"
ElseIf KeyAscii = 97 Then 'Key a
    strUser1Direction = "West"
ElseIf KeyAscii = 100 Then 'Key d
    strUser1Direction = "East"
ElseIf KeyAscii = 112 Then 'Key p
    tmrGamePlay.Enabled = False
    tmrScores.Enabled = False
    tmrSpeed.Enabled = False
    If optStage3.Value = True Then
        tmrGamePlay2.Enabled = False
    End If
ElseIf KeyAscii = 111 Then 'Key o
    tmrGamePlay.Enabled = True
    tmrScores.Enabled = True
    tmrSpeed.Enabled = True
    If optStage3.Value = True Then
        tmrGamePlay2.Enabled = True
    End If
ElseIf KeyAscii = 114 Then 'Key r
    picOutput.Cls
    If optStage1.Value = True Then
        tmrGamePlay.Enabled = True
        tmrGamePlay2.Enabled = False
        tmrScores.Enabled = True
        tmrSpeed.Enabled = True
        lblPlayer1Score.Left = 2000
        picScore1.Left = 4500
        lblSpeed.Left = 5560
        picSpeed.Left = 7120
        lblPlayer2Score.Visible = False
        picScore2.Visible = False
        tmrGamePlay.Interval = 75
        picSpeed.Cls
        intSpeed = 76
        intSpeed = intSpeed - tmrGamePlay.Interval
        picSpeed.Print intSpeed
        strUser1Direction = "North"
        intUser1X = 200
        intUser1Y = 200
        picOutput.PSet (intUser1X, intUser1Y), vbBlue
        intPlayer1Score = 0
        picScore1.Print intPlayer1Score
    ElseIf optStage2.Value = True Then
        tmrGamePlay.Enabled = True
        tmrGamePlay2.Enabled = False
        tmrScores.Enabled = True
        tmrSpeed.Enabled = True
        lblPlayer1Score.Left = 2000
        picScore1.Left = 4500
        lblSpeed.Left = 5560
        picSpeed.Left = 7120
        lblPlayer2Score.Visible = False
        picScore2.Visible = False
        tmrGamePlay.Interval = 40
        picSpeed.Cls
        intSpeed = 41
        intSpeed = intSpeed - tmrGamePlay.Interval
        picSpeed.Print intSpeed
        strUser1Direction = "North"
        intUser1X = 200
        intUser1Y = 200
        picOutput.PSet (intUser1X, intUser1Y), vbBlue
        intPlayer1Score = 0
        picScore1.Print intPlayer1Score
    ElseIf optStage3.Value = True Then
        tmrGamePlay.Enabled = True
        tmrGamePlay2.Enabled = True
        tmrScores.Enabled = True
        tmrSpeed.Enabled = True
        lblPlayer1Score.Left = 240
        picScore1.Left = 2760
        lblSpeed.Left = 7560
        picSpeed.Left = 9120
        lblPlayer2Score.Visible = True
        picScore2.Visible = True
        tmrGamePlay.Interval = 40
        tmrGamePlay2.Interval = 40
        picSpeed.Cls
        intSpeed = 41
        intSpeed = intSpeed - tmrGamePlay.Interval
        picSpeed.Print intSpeed
        strUser1Direction = "North"
        intUser1X = 150
        intUser1Y = 200
        picOutput.PSet (intUser1X, intUser1Y), vbBlue
        strUser2Direction = "North"
        intUser2X = 250
        intUser2Y = 200
        intPlayer1Score = 0
        intPlayer2Score = 0
        picOutput.PSet (intUser2X, intUser2Y), vbRed
        picScore1.Print intPlayer1Score
        picScore2.Print intPlayer2Score
    End If
ElseIf KeyAscii = 101 Then 'Key e
    End
End If

'Player 2
If KeyAscii = 105 Then 'Key i
    strUser2Direction = "North"
ElseIf KeyAscii = 107 Then 'Key k
    strUser2Direction = "South"
ElseIf KeyAscii = 106 Then 'Key j
    strUser2Direction = "West"
ElseIf KeyAscii = 108 Then 'Key l
    strUser2Direction = "East"
End If

End Sub

Private Sub tmrGamePlay_Timer()
picOutput.PSet (intUser1X, intUser1Y), vbBlue

'Player 1
If strUser1Direction = "North" Then
    intUser1Y = intUser1Y - 1
ElseIf strUser1Direction = "South" Then
    intUser1Y = intUser1Y + 1
ElseIf strUser1Direction = "West" Then
    intUser1X = intUser1X - 1
ElseIf strUser1Direction = "East" Then
    intUser1X = intUser1X + 1
End If

If (CStr(picOutput.Point(intUser1X, intUser1Y)) = vbBlue Or CStr(picOutput.Point(intUser1X, intUser1Y)) = vbRed Or CStr(picOutput.Point(intUser1X, intUser1Y)) = vbBlack) Then
    If optStage1.Value = True Then
        MsgBox "| Player 1 score was: " & intPlayer1Score & " |", vbOK, "You Lost!"
        optStage1.Value = False
        tmrGamePlay.Enabled = False
        tmrScores.Enabled = False
        tmrSpeed.Enabled = False
        lblPlayer1Score.Left = 2000
        picScore1.Left = 4500
        lblSpeed.Left = 5560
        picSpeed.Left = 7120
        lblPlayer2Score.Visible = False
        picScore2.Visible = False
        tmrGamePlay.Interval = 75
        picOutput.Cls
        picScore1.Cls
        picScore2.Cls
        picSpeed.Cls
        intSpeed = 76
        intSpeed = intSpeed - tmrGamePlay.Interval
        picSpeed.Print intSpeed
        strUser1Direction = "North"
        intUser1X = 200
        intUser1Y = 200
        picOutput.PSet (intUser1X, intUser1Y), vbBlue
        intPlayer1Score = 0
        picScore1.Print intPlayer1Score
    ElseIf optStage2.Value = True Then
        MsgBox "| Player 1 score was: " & intPlayer1Score & " |", vbOK, "You Lost!"
        optStage2.Value = False
        tmrGamePlay.Enabled = False
        tmrScores.Enabled = False
        tmrSpeed.Enabled = False
        lblPlayer1Score.Left = 2000
        picScore1.Left = 4500
        lblSpeed.Left = 5560
        picSpeed.Left = 7120
        lblPlayer2Score.Visible = False
        picScore2.Visible = False
        tmrGamePlay.Interval = 40
        picOutput.Cls
        picScore1.Cls
        picScore2.Cls
        picSpeed.Cls
        intSpeed = 41
        intSpeed = intSpeed - tmrGamePlay.Interval
        picSpeed.Print intSpeed
        strUser1Direction = "North"
        intUser1X = 200
        intUser1Y = 200
        picOutput.PSet (intUser1X, intUser1Y), vbBlue
        intPlayer1Score = 0
        picScore1.Print intPlayer1Score
    ElseIf optStage3.Value = True And intPlayer1Score > 1 Then
        tmrGamePlay.Enabled = False
    End If
ElseIf (CStr(picOutput.Point(intUser2X, intUser2Y)) = vbBlue Or CStr(picOutput.Point(intUser2X, intUser2Y)) = vbRed Or CStr(picOutput.Point(intUser2X, intUser2Y)) = vbBlack) Then
    If optStage3.Value = True And intPlayer2Score > 1 Then
        tmrGamePlay2.Enabled = False
    End If
End If

End Sub

Private Sub tmrGamePlay2_Timer()
If optStage3.Value = True Then
    picOutput.PSet (intUser2X, intUser2Y), vbRed
End If

'Player 2
If strUser2Direction = "North" Then
    intUser2Y = intUser2Y - 1
ElseIf strUser2Direction = "South" Then
    intUser2Y = intUser2Y + 1
ElseIf strUser2Direction = "West" Then
    intUser2X = intUser2X - 1
ElseIf strUser2Direction = "East" Then
    intUser2X = intUser2X + 1
End If

If (CStr(picOutput.Point(intUser1X, intUser1Y)) = vbBlue Or CStr(picOutput.Point(intUser1X, intUser1Y)) = vbRed Or CStr(picOutput.Point(intUser1X, intUser1Y)) = vbBlack) Then
    If optStage3.Value = True And intPlayer1Score > 1 Then
        tmrGamePlay.Enabled = False
    End If
End If

If (CStr(picOutput.Point(intUser2X, intUser2Y)) = vbBlue Or CStr(picOutput.Point(intUser2X, intUser2Y)) = vbRed Or CStr(picOutput.Point(intUser2X, intUser2Y)) = vbBlack) Then
    If optStage3.Value = True And intPlayer2Score > 1 Then
        tmrGamePlay2.Enabled = False
    End If
End If

End Sub

Private Sub tmrScores_Timer()
'Player 1
If tmrGamePlay.Enabled = True Then
    If strUser1Direction = "North" Then
        intPlayer1Score = intPlayer1Score + 1
    ElseIf strUser1Direction = "South" Then
        intPlayer1Score = intPlayer1Score + 1
    ElseIf strUser1Direction = "West" Then
        intPlayer1Score = intPlayer1Score + 1
    ElseIf strUser1Direction = "East" Then
        intPlayer1Score = intPlayer1Score + 1
    End If
End If

'Player 2
If tmrGamePlay2.Enabled = True Then
    If strUser2Direction = "North" Then
        intPlayer2Score = intPlayer2Score + 1
    ElseIf strUser2Direction = "South" Then
        intPlayer2Score = intPlayer2Score + 1
    ElseIf strUser2Direction = "West" Then
        intPlayer2Score = intPlayer2Score + 1
    ElseIf strUser2Direction = "East" Then
        intPlayer2Score = intPlayer2Score + 1
    End If
End If

If tmrGamePlay.Enabled = False And tmrGamePlay2.Enabled = False Then
    If intPlayer1Score > intPlayer2Score And intPlayer1Score > 1 Then
        MsgBox "| Player 1 score was: " & intPlayer1Score & " | Player 2 score was: " & intPlayer2Score & " |", vbOK, "Player 1 Won!"
    ElseIf intPlayer1Score < intPlayer2Score And intPlayer1Score > 1 Then
        MsgBox "| Player 1 score was: " & intPlayer1Score & " | Player 2 score was: " & intPlayer2Score & " |", vbOK, "Player 2 Won!"
    ElseIf intPlayer1Score = intPlayer2Score And intPlayer1Score > 1 Then
        MsgBox "| Player 1 score was: " & intPlayer1Score & " | Player 2 score was: " & intPlayer2Score & " |", vbOK, "Tie!"
    End If
    optStage3.Value = False
    tmrGamePlay.Enabled = False
    tmrGamePlay2.Enabled = False
    lblPlayer1Score.Left = 240
    picScore1.Left = 2760
    lblSpeed.Left = 7560
    picSpeed.Left = 9120
    lblPlayer2Score.Visible = True
    picScore2.Visible = True
    tmrGamePlay.Interval = 40
    tmrGamePlay2.Interval = 40
    picOutput.Cls
    picScore1.Cls
    picScore2.Cls
    picSpeed.Cls
    intSpeed = 41
    intSpeed = intSpeed - tmrGamePlay.Interval
    picSpeed.Print intSpeed
    strUser1Direction = "North"
    intUser1X = 150
    intUser1Y = 200
    picOutput.PSet (intUser1X, intUser1Y), vbBlue
    strUser2Direction = "North"
    intUser2X = 250
    intUser2Y = 200
    picOutput.PSet (intUser2X, intUser2Y), vbRed
    intPlayer1Score = 0
    intPlayer2Score = 0
    picScore1.Print intPlayer1Score
    picScore2.Print intPlayer2Score
End If

picScore1.Cls
picScore2.Cls
picScore1.Print intPlayer1Score
picScore2.Print intPlayer2Score

End Sub

Private Sub tmrSpeed_Timer()
If (optStage2.Value = True Or optStage3.Value = True) And tmrGamePlay.Interval > 1 Then
    tmrGamePlay.Interval = tmrGamePlay.Interval - 1
    If optStage3.Value = True Then
        tmrGamePlay2.Interval = tmrGamePlay2.Interval - 1
    End If
    intSpeed = intSpeed + 1
    picSpeed.Cls
    picSpeed.Print intSpeed
End If

End Sub

Private Sub optStage1_Click()
tmrGamePlay.Enabled = True
tmrScores.Enabled = True
tmrSpeed.Enabled = True
lblPlayer1Score.Left = 2000
picScore1.Left = 4500
lblSpeed.Left = 5560
picSpeed.Left = 7120
lblPlayer2Score.Visible = False
picScore2.Visible = False
tmrGamePlay.Interval = 75
picOutput.Cls
picScore1.Cls
picScore2.Cls
picSpeed.Cls
intSpeed = 76
intSpeed = intSpeed - tmrGamePlay.Interval
picSpeed.Print intSpeed
strUser1Direction = "North"
intUser1X = 200
intUser1Y = 200
picOutput.PSet (intUser1X, intUser1Y), vbBlue
intPlayer1Score = 0
picScore1.Print intPlayer1Score
    
End Sub

Private Sub optStage2_Click()
tmrGamePlay.Enabled = True
tmrScores.Enabled = True
tmrSpeed.Enabled = True
lblPlayer1Score.Left = 2000
picScore1.Left = 4500
lblSpeed.Left = 5560
picSpeed.Left = 7120
lblPlayer2Score.Visible = False
picScore2.Visible = False
tmrGamePlay.Interval = 40
picOutput.Cls
picScore1.Cls
picScore2.Cls
picSpeed.Cls
intSpeed = 41
intSpeed = intSpeed - tmrGamePlay.Interval
picSpeed.Print intSpeed
strUser1Direction = "North"
intUser1X = 200
intUser1Y = 200
picOutput.PSet (intUser1X, intUser1Y), vbBlue
intPlayer1Score = 0
picScore1.Print intPlayer1Score

End Sub

Private Sub optStage3_Click()
tmrGamePlay.Enabled = True
tmrGamePlay2.Enabled = True
tmrScores.Enabled = True
tmrSpeed.Enabled = True
lblPlayer1Score.Left = 240
picScore1.Left = 2760
lblSpeed.Left = 7560
picSpeed.Left = 9120
lblPlayer2Score.Visible = True
picScore2.Visible = True
tmrGamePlay.Interval = 40
tmrGamePlay2.Interval = 40
picOutput.Cls
picScore1.Cls
picScore2.Cls
picSpeed.Cls
intSpeed = 41
intSpeed = intSpeed - tmrGamePlay.Interval
picSpeed.Print intSpeed
strUser1Direction = "North"
intUser1X = 150
intUser1Y = 200
picOutput.PSet (intUser1X, intUser1Y), vbBlue
strUser2Direction = "North"
intUser2X = 250
intUser2Y = 200
picOutput.PSet (intUser2X, intUser2Y), vbRed
intPlayer1Score = 0
intPlayer2Score = 0
picScore1.Print intPlayer1Score
picScore2.Print intPlayer2Score

End Sub

