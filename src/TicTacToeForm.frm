VERSION 5.00
Begin VB.Form TicTacToe 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic-Tac-Toe"
   ClientHeight    =   5304
   ClientLeft      =   120
   ClientTop       =   768
   ClientWidth     =   8604
   DrawWidth       =   3
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   27.6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "TicTacToeForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5304
   ScaleWidth      =   8604
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar Intelligence 
      Height          =   255
      Left            =   600
      Max             =   5
      Min             =   1
      TabIndex        =   14
      Top             =   4080
      Value           =   1
      Width           =   3372
   End
   Begin VB.CheckBox HumanX 
      BackColor       =   &H00808080&
      Caption         =   "Human plays X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox HumanFirst 
      BackColor       =   &H00808080&
      Caption         =   "Human goes first"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton Quit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton PlayAgain 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.Label IntelligenceLevel 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Intelligence Level:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   732
      Left            =   600
      TabIndex        =   16
      Top             =   3360
      Width           =   1692
   End
   Begin VB.Label Level 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   492
      Left            =   2160
      TabIndex        =   15
      Top             =   3720
      Width           =   1812
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   7
      Visible         =   0   'False
      X1              =   7800
      X2              =   4200
      Y1              =   1080
      Y2              =   4680
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   6
      Visible         =   0   'False
      X1              =   4200
      X2              =   7800
      Y1              =   1080
      Y2              =   4680
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   5
      Visible         =   0   'False
      X1              =   7200
      X2              =   7200
      Y1              =   1080
      Y2              =   4680
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   4
      Visible         =   0   'False
      X1              =   6000
      X2              =   6000
      Y1              =   1080
      Y2              =   4680
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   3
      Visible         =   0   'False
      X1              =   4800
      X2              =   4800
      Y1              =   1080
      Y2              =   4680
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   2
      Visible         =   0   'False
      X1              =   4080
      X2              =   7920
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   1
      Visible         =   0   'False
      X1              =   4080
      X2              =   7920
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   0
      Visible         =   0   'False
      X1              =   4080
      X2              =   7920
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Index           =   8
      Left            =   6720
      TabIndex        =   13
      Top             =   3600
      Width           =   972
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Index           =   7
      Left            =   5520
      TabIndex        =   12
      Top             =   3600
      Width           =   972
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Index           =   6
      Left            =   4320
      TabIndex        =   11
      Top             =   3600
      Width           =   972
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Index           =   5
      Left            =   6720
      TabIndex        =   10
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Index           =   4
      Left            =   5520
      TabIndex        =   9
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Index           =   3
      Left            =   4320
      TabIndex        =   8
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Index           =   2
      Left            =   6720
      TabIndex        =   7
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Index           =   1
      Left            =   5520
      TabIndex        =   6
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Index           =   0
      Left            =   4320
      TabIndex        =   5
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Tic Tac Toe"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1452
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7692
   End
   Begin VB.Line Line4 
      X1              =   4200
      X2              =   7800
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      X1              =   4200
      X2              =   7800
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      X1              =   6600
      X2              =   6600
      Y1              =   1080
      Y2              =   4680
   End
   Begin VB.Line Line1 
      X1              =   5400
      X2              =   5400
      Y1              =   1080
      Y2              =   4680
   End
   Begin VB.Menu game 
      Caption         =   "Game"
      Begin VB.Menu newgame 
         Caption         =   " "
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu gamexit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu helptips 
         Caption         =   "Tips"
         Shortcut        =   {F1}
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu helpabout 
         Caption         =   "About"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "TicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StartSquare As Integer, EndSquare As Integer, Direction As Integer
Dim XO(2) As String
Dim Blank As Integer
Dim CaseCounter As Integer
Dim XsInARow As Integer, OsInARow As Integer
Dim HumanPlayingAsX As Boolean
Dim HumanMovedAlready As Boolean, CompMovedAlready As Boolean
Dim GameOver As Boolean
Dim SquareValue(8) As Integer
Dim RandChoice(8) As Integer
Dim WinLine As Integer
Dim TieGame As Boolean
Dim strCPU As String
Dim strHuman As String

Private Sub Form_Load()
Dim Index As Integer
Cls
Randomize
For Index = 0 To 8
Square(Index).Caption = ""
SquareValue(Index) = 0
Next Index

For Index = 0 To 7
Win(Index).Visible = False
Next Index

If HumanX.Value = 1 Then
    HumanPlayingAsX = True
    strCPU = "O"
    strHuman = "X"
Else
    HumanPlayingAsX = False
    strCPU = "X"
    strHuman = "O"
End If

HumanMovedAlready = False
CompMovedAlready = False

HumanX.Visible = True
HumanFirst.Visible = True
HumanFirst.Value = 1
Call IntelligenceLevelCase

GameOver = False
TieGame = False

PlayAgain.Caption = "New Game"
newgame.Caption = "New Game"

End Sub


Private Sub helpabout_Click()
MsgBox ("             Tic Tac Toe         " & vbCrLf & vbCrLf & _
              "             Version 1.0.8" & vbCrLf & vbCrLf & _
              "             Debugging And Logic Programmer       Mannan Sultanpuri" & vbCrLf & _
              "             Design And Form Programmer               Tarunbir Singh" & vbCrLf & _
              "             AI Programmer                                        Arvin Singh" & vbCrLf & vbCrLf & _
              "             Copyright Paradox Betalabs")
End Sub

Private Sub helptips_Click()
If MsgBox("Help and Tips!" & vbCrLf & vbCrLf & _
              "You can arrange the A-IQ to different levels. Maximum level is the toughest one." & vbCrLf & vbCrLf & _
              "This is A.I. based game. Repeating moves used to win the last game won't work everytime." & vbCrLf & _
              "Just enjoy!" & vbCrLf & vbCrLf & _
              "Continue?", vbYesNo + vbQuestion, "Help and Tips!") = vbYes Then
              Else
              End
              End If
                           
End Sub

Private Sub newgame_Click()
Dim Index As Integer
Dim OldColor As String

OldColor = Square(0).BackColor
    For Index = 0 To 7
        Square(Index).BackColor = Square(Index + 1).BackColor
    Next Index
Square(8).BackColor = OldColor

Form_Load
End Sub

Private Sub GamExit_Click()
End
End Sub

Private Sub Square_Click(Index As Integer)
Dim Flash100 As Integer

If Square(Index).Caption = "" And HumanMovedAlready = False Then
    HumanMovedAlready = True
    Square(Index).Caption = strHuman
    Call CheckForWin
    
    If GameOver = False Then
        Call CheckForCatsGame
        Call ComputerTurn
    End If
Else
    Beep
    For Flash100 = 1 To 100
        Me.BackColor = RGB(Int(256 * Rnd), Int(256 * Rnd), Int(256 * Rnd))
        Me.BackColor = "&H00808080"
    Next Flash100
End If
End Sub




Private Sub HumanFirst_Click()
If HumanFirst.Value = 1 Then
    HumanMovedAlready = False
Else
    HumanMovedAlready = True
    Cls
    Call ComputerTurn
End If
End Sub


Private Sub HumanX_Click()
If HumanX.Value = 1 Then
    HumanPlayingAsX = True
    strCPU = "O"
    strHuman = "X"
Else
    HumanPlayingAsX = False
    strCPU = "X"
    strHuman = "O"
End If

End Sub


Private Sub Intelligence_Change()
Dim Red As Integer, Green As Integer, Blue As Integer

Red = (Intelligence.Value / Intelligence.Max) * 255
Green = 0
Blue = (1 / Intelligence.Value) * 255
Level.ForeColor = RGB(Red, Green, Blue)
Call IntelligenceLevelCase
End Sub


Private Sub PlayAgain_Click()
Dim Index As Integer
Dim OldColor As String

OldColor = Square(0).BackColor
    For Index = 0 To 7
        Square(Index).BackColor = Square(Index + 1).BackColor
    Next Index
Square(8).BackColor = OldColor

Form_Load
End Sub


Private Sub Quit_Click()
Unload Me
End Sub


Private Sub XWins()
Call GameIsOver
MsgBox ("X wins!")
End Sub


Private Sub OWins()
GameOver = True
Call GameIsOver
MsgBox ("O Wins!")
End Sub


Private Sub CatsGame()
TieGame = True
Call GameIsOver
MsgBox ("Cat's Game!")
End Sub


Private Sub GameIsOver()
GameOver = True
PlayAgain.Caption = "Play Again?"
newgame.Caption = "Play Again?"
Beep
If TieGame = False Then Win(CaseCounter).Visible = True
End Sub


Private Sub BoardCheck()

Select Case CaseCounter
               

    Case 0
        StartSquare = 0
        EndSquare = 2
        Direction = 1
    Case 1
        StartSquare = 3
        EndSquare = 5
        Direction = 1
    Case 2
        StartSquare = 6
        EndSquare = 8
        Direction = 1
        
    Case 3
        StartSquare = 0
        EndSquare = 6
        Direction = 3
    Case 4
        StartSquare = 1
        EndSquare = 7
        Direction = 3
    Case 5
        StartSquare = 2
        EndSquare = 8
        Direction = 3

    Case 6
        StartSquare = 0
        EndSquare = 8
        Direction = 4
    Case 7
        StartSquare = 2
        EndSquare = 6
        Direction = 2
End Select
End Sub


Private Sub XOSelect(XOSquare)

Select Case XOSquare
    Case 1
        XO(0) = "X"
        XO(1) = "X"
        XO(2) = ""
        Blank = EndSquare
    Case 2
        XO(0) = "X"
        XO(1) = ""
        XO(2) = "X"
        Blank = StartSquare + Direction
    Case 3
        XO(0) = ""
        XO(1) = "X"
        XO(2) = "X"
        Blank = StartSquare
    Case 4
        XO(0) = "O"
        XO(1) = "O"
        XO(2) = ""
        Blank = EndSquare
    Case 5
        XO(0) = "O"
        XO(1) = ""
        XO(2) = "O"
        Blank = StartSquare + Direction
    Case 6
        XO(0) = ""
        XO(1) = "O"
        XO(2) = "O"
        Blank = StartSquare
End Select
End Sub


Private Sub IntelligenceLevelCase()
    Select Case Intelligence.Value
        Case 1
            Level.Caption = "Mindless"
        Case 2
            Level.Caption = "Poor"
        Case 3
            Level.Caption = "Average"
        Case 4
            Level.Caption = "Clever"
        Case 5
            Level.Caption = "Genius"
    End Select
End Sub


Private Sub CheckForCatsGame()
Dim Index As Integer
Dim Space As Integer

For Index = 0 To 8
    If Square(Index).Caption = "" Then Space = Space + 1
Next Index

If Space = 0 Then
    Call CatsGame
End If
End Sub


Private Sub SpecialAI()
Dim MiddleSideSquare As Integer
Dim Index As Integer

If (Square(0).Caption = strHuman And Square(8).Caption = strHuman) _
Or (Square(2).Caption = strHuman And Square(6).Caption = strHuman) Then
    
    For Index = 1 To 7 Step 2
        SquareValue(Index) = 10
    Next Index
    
End If

If Square(1).Caption = strHuman And Square(5).Caption = strHuman Then
    SquareValue(6) = 0
End If
    
If Square(5).Caption = strHuman And Square(7).Caption = strHuman Then
    SquareValue(0) = 0
End If
    
If Square(3).Caption = strHuman And Square(7).Caption = strHuman Then
    SquareValue(2) = 0
End If
    
If Square(1).Caption = strHuman And Square(3).Caption = strHuman Then
    SquareValue(8) = 0
End If


If Intelligence.Value > 4 Then
    For Index = 1 To 7 Step 2
        If Square(Index).Caption = strHuman Then
            MiddleSideSquare = MiddleSideSquare + 1
        End If
    Next Index
    
    If MiddleSideSquare = 1 Then
        If Square(1).Caption = strHuman Then
            SquareValue(0) = 10
            SquareValue(2) = 10
        End If
        
        If Square(3).Caption = strHuman Then
            SquareValue(0) = 10
            SquareValue(6) = 10
        End If
        
        If Square(5).Caption = strHuman Then
            SquareValue(2) = 10
            SquareValue(8) = 10
        End If
        
        If Square(7).Caption = strHuman Then
            SquareValue(6) = 10
            SquareValue(8) = 10
        End If
    End If
End If

End Sub


Private Sub CheckForWin()
Dim Index As Integer

If HumanX.Visible = True Then Call RemoveOptions

CaseCounter = 0
While CaseCounter < 8
                
    Call BoardCheck
        XsInARow = 0
        OsInARow = 0
    
    For Index = StartSquare To EndSquare Step Direction
        If Square(Index).Caption = "X" Then XsInARow = XsInARow + 1
        If Square(Index).Caption = "O" Then OsInARow = OsInARow + 1
    Next Index
    
If XsInARow = 3 Then Call XWins
If OsInARow = 3 Then Call OWins
      
CaseCounter = CaseCounter + 1
Wend
End Sub

    

Private Sub RemoveOptions()
    HumanX.Visible = False
    HumanFirst.Visible = False
End Sub

Private Sub CompAI()

Dim XOSquare As Integer
Dim Index As Integer
Dim Space As Integer

For Index = 0 To 8
    If Square(Index).Caption = "" Then
        If Index Mod 2 = 0 Then
            SquareValue(Index) = 5
                Else
            SquareValue(Index) = 1
        End If
    Else
        SquareValue(Index) = 0
    End If
Next Index

For Index = 0 To 8
    If Square(Index).Caption = "" Then Space = Space + 1
Next Index

If Space = 6 And Intelligence.Value > 3 Then
    Call SpecialAI
End If

CaseCounter = 0
XOSquare = 1

If Intelligence.Value > 1 Then
    If Square(4).Caption = "" Then
        SquareValue(4) = 10
    End If


    While CaseCounter < 8


    Call BoardCheck
    CaseCounter = CaseCounter + 1
        
        For XOSquare = 1 To 6
        Call XOSelect(XOSquare)

            If Square(StartSquare).Caption = XO(0) And _
            Square(StartSquare + Direction).Caption = XO(1) And _
            Square(EndSquare).Caption = XO(2) Then


                If HumanPlayingAsX = True Then
                    If XOSquare <= 3 Then
                        If Intelligence.Value > 2 And SquareValue(Blank) < 999 Then
                            SquareValue(Blank) = 500
                        End If
                    Else
                        If Intelligence.Value > 1 Then
                            SquareValue(Blank) = 999
                        End If
                    End If
                    

                Else
                    If XOSquare <= 3 Then
                        If Intelligence.Value > 1 Then
                            SquareValue(Blank) = 999
                        End If
                    Else
                        If Intelligence.Value > 2 And SquareValue(Blank) < 999 Then
                            SquareValue(Blank) = 500
                        End If
                    End If
                End If
            End If
        Next XOSquare
    Wend
End If


Debug.Print
For Index = 0 To 8
If Index Mod 3 = 0 Then Debug.Print
Debug.Print SquareValue(Index); " ";
If SquareValue(Index) < 100 Then Debug.Print " ";
If SquareValue(Index) < 10 Then Debug.Print " ";
Next Index

End Sub


Private Sub CompSquareSelection()
Dim Index As Integer
Dim HighVal As Integer
Dim Counter As Integer
Dim FinalChoice As Integer

Counter = 0
HighVal = 0
For Index = 0 To 8
If SquareValue(Index) = HighVal Then
    RandChoice(Counter) = Index
    Counter = Counter + 1
End If

If SquareValue(Index) > HighVal Then
    HighVal = SquareValue(Index)
    For Counter = 1 To 8
    RandChoice(Counter) = 0
    Next Counter
    RandChoice(0) = Index
    Counter = 1
End If
Next Index

Index = Int(Rnd * Counter)
Square(RandChoice(Index)).Caption = strCPU
HumanMovedAlready = False
End Sub


Private Sub ComputerTurn()

If GameOver = False Then
    Call CompAI
    Call CompSquareSelection
    Call CheckForWin
    Call CheckForCatsGame
End If

End Sub
