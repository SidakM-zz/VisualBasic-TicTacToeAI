VERSION 5.00
Begin VB.Form frmTicTacHuman 
   BackColor       =   &H0080FFFF&
   Caption         =   "Can You Beat Sidak?"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPlayAgain 
      Caption         =   "Play Again"
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CheckBox PlayX 
      Caption         =   "Play X"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblTied 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   15
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label lblCwins 
      BackColor       =   &H0080FFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   14
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label lblPwins 
      BackColor       =   &H0080FFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   13
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label lblComputer 
      BackColor       =   &H0080FFFF&
      Caption         =   "Computer Win(s):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label lblPlayer 
      BackColor       =   &H0080FFFF&
      Caption         =   "Player Win(s):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Line winDirection 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Index           =   8
      X1              =   840
      X2              =   4200
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line winDirection 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   6
      X1              =   360
      X2              =   4200
      Y1              =   240
      Y2              =   4800
   End
   Begin VB.Line winDirection 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   7
      X1              =   4200
      X2              =   360
      Y1              =   360
      Y2              =   4800
   End
   Begin VB.Line winDirection 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   5
      X1              =   3840
      X2              =   3840
      Y1              =   240
      Y2              =   4680
   End
   Begin VB.Line winDirection 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   4
      X1              =   2280
      X2              =   2280
      Y1              =   240
      Y2              =   4680
   End
   Begin VB.Line winDirection 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   3
      X1              =   720
      X2              =   720
      Y1              =   240
      Y2              =   4680
   End
   Begin VB.Line winDirection 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   2
      X1              =   120
      X2              =   4320
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line winDirection 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   4320
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line winDirection 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   120
      X2              =   4320
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   8
      Left            =   3240
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   7
      Left            =   1680
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   5
      Left            =   3240
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   4
      Left            =   1680
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmTicTacHuman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim filledorUnfilled(2) As String
Dim moveWorth(8) As Integer 'higher = better move SquareValue
Dim XisChecked As Boolean ' if human is x
Dim CPUGAME As String ' x or o
Dim HUMANGAME As String
Dim HumanMove As Boolean 'humanmovealready
Dim counter2 As Integer
Dim Begin, Ending, Direction, xtotal, ototal, vacant As Integer
Dim RANDOMNESS(8) As Integer
Dim gameTied As Boolean
Dim Pwins, Cwins As Integer
Dim tied_games As Integer

Private Sub cmdPlayAgain_Click()
Call Form_Load
End Sub

Private Sub Form_Load()
Dim Counter As Integer

For Counter = 0 To 8 'clear all the x and O's from the board
lblSquare(Counter).Caption = ""
moveWorth(Counter) = 0
Next Counter

For Counter = 0 To 7 'clear the line which tells the win direction from the board
winDirection(Counter).Visible = False
Next Counter


'human can choose to play as x or o  X BY DEFAULT
If PlayX.Value = 1 Then
    XisChecked = True
    CPUGAME = "O"
    HUMANGAME = "X"
Else
    XisChecked = False
    CPUGAME = "X"
    HUMANGAME = "O"
End If



End Sub
Private Sub XWON()
MsgBox ("X HAS WON")
Call GamesOver
If HUMANGAME = "X" Then
    Pwins = Pwins + 1
    lblPwins.Caption = Str(Pwins)
Else
    Cwins = Cwins + 1
    lblCwins.Caption = Str(Cwins)
    
End If
End Sub


Private Sub OWON()
Call GamesOver
MsgBox ("O HAS WON")
If HUMANGAME = "O" Then
    Pwins = Pwins + 1
    lblPwins.Caption = Str(Pwins)
Else
    Cwins = Cwins + 1
    lblCwins.Caption = Str(Cwins)
End If
End Sub


Private Sub CATGAME()
Call Form_Load
MsgBox ("CAT'S GAME")
tied_games = tied_games + 1
lblTied.Caption = Str(tied_games) + " games were tied."
End Sub


Private Sub GamesOver()
If gameTied = False Then
    winDirection(counter2).Visible = True
    Else
    End If
    

End Sub
Private Sub CheckCatGame()
Dim catcounter, block As Integer


For catcounter = 0 To 8 'will check every square for an empty square
    If lblSquare(catcounter).Caption = "" Then
        block = 1 'changing block to one will tell program that a space is empty
        Else
        End If
Next catcounter

If block = 0 Then
    Call CATGAME
End If
End Sub


Private Sub Label1_Click()

End Sub

Private Sub lblSquare_Click(Index As Integer)
If lblSquare(Index).Caption = "" And HumanMove = False Then
    HumanMove = True
    lblSquare(Index).ForeColor = vbBlue
    lblSquare(Index).Caption = HUMANGAME
    Call CheckForWin
    
        Call CheckCatGame
        Call ComputerTurn
    
    

Else
   
End If
End Sub
Private Sub CheckForWin()
Dim Counter As Integer

'check every possible direction one could
'win and triggers the end of game procedures if anyone gets 3 in a row.
counter2 = 0
Do While counter2 < 8
                
    Call ScanBoard
        xtotal = 0 'change the totals to 0 after checking last possible winning line
        ototal = 0
    
    For Counter = Begin To Ending Step Direction 'take values from board check and see if any are in a row all checked
        If lblSquare(Counter).Caption = "X" Then 'if the value is equal to x then add one to the x totalsame for o
            xtotal = xtotal + 1
        Else
        End If
        
        If lblSquare(Counter).Caption = "O" Then
            ototal = ototal + 1
        Else
      
        End If
        
        
    Next Counter
    
If xtotal = 3 Then ' if three x's are checked in a line say x wins same for o
Call XWON
Else
End If

If ototal = 3 Then
Call OWON
Else
End If

counter2 = counter2 + 1
Loop
End Sub
Private Sub ScanBoard()

Select Case counter2
               
'provides values for counter in checkforwin to check every line

    Case 0
        Begin = 0
        Ending = 2
        Direction = 1
    Case 1
        Begin = 3
        Ending = 5
        Direction = 1
    Case 2
        Begin = 6
        Ending = 8
        Direction = 1
            
    Case 3
        Begin = 0
        Ending = 6
        Direction = 3 ' direction is working by adding the set amount of times eaach time in the other functions to get the vertical or diagnol line so in this case its 0(first box),3(as direction 3 would be added in the function that is getting this value), 6
                        ' this would check for the left most vertical line
    Case 4
        Begin = 1
        Ending = 7
        Direction = 3
    Case 5
        Begin = 2
        Ending = 8
        Direction = 3
            
    Case 6
        Begin = 0
        Ending = 8
        Direction = 4
    Case 7
        Begin = 2
        Ending = 6
        Direction = 2
End Select
End Sub

Private Sub ComputerTurn()

    Call AIntelligence 'calls the ai
    Call selectSquareCOmputer
    Call CheckForWin
    Call CheckCatGame

End Sub


Private Sub AIntelligence()
    Dim SideinMiddle
    Dim Counter As Integer
    Dim block As Integer
    Dim CHANGESquare As Integer
    
    
    'values will be assigned to each square based on best move so higher value better move
    
    
    'The most important squares are the corners and middle so assigning more points to them would be better
    
    For Counter = 0 To 8
    If lblSquare(Counter).Caption = "" Then
        If Counter Mod 2 = 0 Then ' square 0,2,middle square and all other corners will have a 0 remainder when moded by 2
            moveWorth(Counter) = 5
        Else                'all other squares get 1 as value
            moveWorth(Counter) = 1
            
        End If
    Else 'ALL FILLED SQUARES ARE ZERO VALUED
    moveWorth(Counter) = 0
    End If
    Next Counter
    
    
    
    ' the third move of the game is the most imprtant as most techniques take place during this time
    For Counter = 0 To 8
        If lblSquare(Counter).Caption = "" Then
            block = block + 1 'will count number of blocks that are unfilled
        Else
        End If
    Next Counter
Special Maneuvers
    If block = 6 Then 'if 6 blocks are unfilled, meaining 3 are filled then it means that this is the third move
       
       'on the the third move the player could be putting his mark on the diagnols therefore the middle - side squares are the move the computer should pick
            If (lblSquare(0).Caption = HUMANGAME And lblSquare(8).Caption = HUMANGAME) Or (lblSquare(2).Caption = HUMANGAME And lblSquare(6).Caption = HUMANGAME) Then 'compares text in box to text set at starting of game
            For Counter = 1 To 7 Step 2
            moveWorth(Counter) = 10 'square 3 and 5 and 1 and 7 will get their worth increased exponentially
            Next Counter
            Else
            End If
            
        'Special Maneuvers TWO
        If lblSquare(1).Caption = HUMANGAME And lblSquare(5).Caption = HUMANGAME Then    'This maneuver is used so the computer can avoid being trapped if the player will put the marks on both sides of a corner
            moveWorth(6) = 0                                                             'If the computer then on the third move places its move on the opposite corner it will lose so if the conditions are met the opposing corner will get a zero value so it does not get chosen by the comp.
        ElseIf lblSquare(5).Caption = HUMANGAME And lblSquare(7).Caption = HUMANGAME Then
            moveWorth(0) = 0
        ElseIf lblSquare(3).Caption = HUMANGAME And lblSquare(7).Caption = HUMANGAME Then
            moveWorth(2) = 0
        ElseIf lblSquare(1).Caption = HUMANGAME And lblSquare(3).Caption = HUMANGAME Then
            moveWorth(8) = 0
        End If

        Else
        End If
'Finally SIDAK"S SUPER UNBEATABLE LOGIC aka SPECIAL MANEUVAR 3

    For Counter = 1 To 7 Step 2                       'PLAYER CAN TAKE a corner and the opposite side's middle square which will set them up for a double trap. The stop this the computer can place
        If lblSquare(Counter).Caption = HUMANGAME Then 'its next move on either side of the "opposite side's middle square" which will ruin the players plot
            SideinMiddle = SideinMiddle + 1            'This loop will identify if the player is setting up for such a plot as it starts at 1 which will be the top mid square and will test
        End If                                         'every 2 values as these would be a part of the player's plot. If one of these squares tests true the sidemiddle will become 1 inwhich case the following if statement will test true
    Next Counter
    
    If SideinMiddle = 1 Then                           ' this will then identify the plot square and give both squares on its side a high value which will get the computer to likely choose that square and ruin the player's plot
        If lblSquare(1).Caption = HUMANGAME Then
            moveWorth(0) = 10
            moveWorth(2) = 10
        
        ElseIf lblSquare(3).Caption = HUMANGAME Then
            moveWorth(0) = 10
            moveWorth(6) = 10
        
        ElseIf lblSquare(5).Caption = HUMANGAME Then
            moveWorth(2) = 10
            moveWorth(8) = 10
        
        ElseIf lblSquare(7).Caption = HUMANGAME Then
            moveWorth(6) = 10
            moveWorth(8) = 10
        End If
    End If

'SPECIAL MANEUVERS END SPECIAL MANEUVERS END SPECIAL MANEUVERS END SPECIAL MANEUVERS END SPECIAL MANEUVERS ENDSPECIAL MANEUVERS ENDSPECIAL MANEUVERS ENDSPECIAL MANEUVERS ENDSPECIAL MANEUVERS ENDSPECIAL MANEUVERS END


counter2 = 0 'value will be used for board scan
CHANGESquare = 1 ' value to be goven to method as parameter

 If lblSquare(4).Caption = "" Then ' if the middle square is empty give it a higher value
        moveWorth(4) = 10
 Else
 
 End If



Do While counter2 < 8 'used to reterate thriught the ScanBoard and check all lines for this peice of logic

    Call ScanBoard
    counter2 = counter2 + 1
        
        For CHANGESquare = 1 To 6
        Call CHANGESelect(CHANGESquare)

        If lblSquare(Begin).Caption = filledorUnfilled(0) And lblSquare(Begin + Direction).Caption = filledorUnfilled(1) And lblSquare(Ending).Caption = filledorUnfilled(2) Then 'checking if the current values from the select case in CHAGESELECT ARE TRUE
            
            If XisChecked = True Then 'if human is pllaying as x
                    If CHANGESquare <= 3 Then  'the cases less then 3 are all x's in the CHANGeSELECT function therefore this means that this move will be a winning move
                        
                        If moveWorth(vacant) < 999 Then 'this is here as a block is worse than an immediate win as a block could get a higher value if some of the condtions of the previous logic are met
                            moveWorth(vacant) = 500
                        Else
                        End If
                         
                    Else
                        moveWorth(vacant) = 999 'if the move is greater than 999 just give it a value of 999 this is for an immediate win
                    End If
                    
                    'next part is the same but for o reversed as if opponent is o the first value will be the immediate win
            Else
                    If CHANGESquare <= 3 Then
                       moveWorth(vacant) = 999
                    
                    ElseIf moveWorth(vacant) < 999 Then
                            moveWorth(vacant) = 500
                    End If
                    
                    
                    End If
            End If
    Next CHANGESquare
Loop
                    
                    
End Sub
Private Sub CHANGESelect(CHANGESquare) 'will take the parameter


Select Case CHANGESquare
    Case 1
        filledorUnfilled(0) = "X" 'will check if row is checked
        filledorUnfilled(1) = "X"
        filledorUnfilled(2) = ""
        vacant = Ending  'the vacant square will be given the positioning of the last square
    Case 2
        filledorUnfilled(0) = "X"
        filledorUnfilled(1) = ""
        filledorUnfilled(2) = "X"
        vacant = Begin + Direction 'current value from ScanBoard
    Case 3
        filledorUnfilled(0) = ""
        filledorUnfilled(1) = "X"
        filledorUnfilled(2) = "X"
        vacant = Begin
    Case 4
        filledorUnfilled(0) = "O"
        filledorUnfilled(1) = "O"
        filledorUnfilled(2) = ""
        vacant = Ending
    Case 5
        filledorUnfilled(0) = "O"
        filledorUnfilled(1) = ""
        filledorUnfilled(2) = "O"
        vacant = Begin + Direction
    Case 6
        filledorUnfilled(0) = ""
        filledorUnfilled(1) = "O"
        filledorUnfilled(2) = "O"
        vacant = Begin   'if the value is O and the computer is 0 it will gove it the highest value so it can win. If the computer isnt o it will gove it the second highest possible value
End Select
End Sub

Private Sub selectSquareCOmputer() 'select the square for the computer
''the upcoming loop will find the highest value move, store it in another array(random) and then see if the same value is present anywhere else in the x and o grid ,
'if it finds a higher value it will repeat this process until the main loop has checked all squares

Dim place As Integer
Dim Highest_Value As Integer
Dim Counter As Integer
Dim moveForComputer As Integer

Counter = 0
Highest_Value = 0
For place = 0 To 8
If moveWorth(place) = Highest_Value Then ' if the move worth is the current highest value
    RANDOMNESS(Counter) = place 'the index number of the square will be added to the random array which will than choose
    Counter = Counter + 1
Else
End If

If moveWorth(place) > Highest_Value Then
    Highest_Value = moveWorth(place)
    For Counter = 1 To 8
    RANDOMNESS(Counter) = 0
    Next Counter
    RANDOMNESS(0) = place
    Counter = 1
    Else
End If
Next place


place = Int(Rnd * Counter)
lblSquare(RANDOMNESS(place)).ForeColor = vbMagenta
lblSquare(RANDOMNESS(place)).Caption = CPUGAME
HumanMove = False
End Sub


