VERSION 5.00
Begin VB.UserControl Puzzle 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   LockControls    =   -1  'True
   PropertyPages   =   "Puzzle.ctx":0000
   ScaleHeight     =   1800
   ScaleWidth      =   2250
   ToolboxBitmap   =   "Puzzle.ctx":0015
   Begin VB.Timer tmrShowCorrect 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1260
      Top             =   660
   End
   Begin VB.PictureBox picShow 
      AutoRedraw      =   -1  'True
      Height          =   645
      Left            =   870
      ScaleHeight     =   585
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrShowFullPic 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1740
      Top             =   600
   End
   Begin VB.Timer tmrScore 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1740
      Top             =   1140
   End
   Begin VB.PictureBox picParts 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      FillStyle       =   0  'Solid
      Height          =   585
      Index           =   0
      Left            =   90
      ScaleHeight     =   585
      ScaleWidth      =   675
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox picSmall 
      AutoRedraw      =   -1  'True
      Height          =   795
      Left            =   1140
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   900
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Timer tmrShuffle 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1740
      Top             =   90
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   120
      ScaleHeight     =   630
      ScaleWidth      =   870
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   930
   End
End
Attribute VB_Name = "Puzzle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Note: for using drag and drop, I changed "DragMode" property of "picParts" to automatic.
'
Public Enum difficultyModes
    Easy = 0
    Standard = 1
    Hard = 2
End Enum
Public Difficulty As difficultyModes ' added in update

Public Rows As Integer, Columns As Integer

Private myPartsCollection As New Collection
Private itsReady As Boolean ' puzzle is ready for play or not
Private myTotalMoves As Long
Private myUndoPic1 As Integer, myUndoPic2 As Integer ' data for undo moves
Private myScore As Long
Private solveMe As Boolean ' added in update

Public Event Counter(ByVal NumberOfMoves As Long)
Public Event Resize(ByVal Width As Integer, ByVal Height As Integer)
Public Event Score(ByVal TotalScore As Long, ByVal LostTime As Long)
Public Event Finished()

' select new picture for puzzle
Public Sub mainPicture(ByVal newPic As Object)
    picSource.Picture = newPic
End Sub

' start
Public Sub Start(ByVal Width As Integer, ByVal Height As Integer, Optional partsBorder As Boolean = True)
    Call MakeObject(Width, Height, partsBorder) ' at the first, make picture at the correct size on multiple parts
    tmrShuffle.Enabled = True ' then change location of parts
End Sub

'shrink your picture to multi parts
Private Sub MakeObject(ByVal Width As Integer, ByVal Height As Integer, ByVal partsBorder As Boolean)
    On Error GoTo errHandler
    Dim i As Integer, j As Integer, k As Integer
    Dim myScale As Single, myWidth As Single, myHeight As Single
    
    Randomize Timer
    
    myTotalMoves = 0
    itsReady = False
    myScore = IIf(Difficulty = Easy, (Columns - 1) * (Rows - 1), Columns * Rows) ' in easy mod 1 column and 1 row are fix.
    myScore = myScore * 100 + myScore * 5 * Difficulty
    myScale = picSource.Height / picSource.Width
    
    myHeight = (Height - 25) / Rows
    myWidth = (Height - 25) / myScale / Columns
    If myWidth * Columns > Width Then
        myWidth = Width / Columns
        myHeight = Width * myScale / Rows
    End If
    RaiseEvent Resize(myWidth * Columns + 25, myHeight * Rows + 25)
    
    ' a small pic for status bar as help
    picSmall.PaintPicture picSource.Picture, 0, 0, IIf(myScale > 1, 500 / myScale, 600), IIf(myScale > 1, 700, myScale * 1000)
    
    DoEvents
    For i = 1 To Columns
        For j = 1 To Rows
            k = picParts.UBound + 1 ' get highest index of parts and add one
            Load picParts(k) ' make new part
            With picParts(k) ' set properties of new part
                .Width = myWidth * 1.5
                .Height = myHeight * 1.5
                .PaintPicture picSource, 0, 0, myWidth + 10, myHeight + 10, (picSource.Width - 60) / Columns * (i - 1), (picSource.Height - 60) / Rows * (j - 1), (picSource.Width - 60) / Columns, (picSource.Height - 60) / Rows
                .Top = 20 + (j - 1) * myHeight
                .Width = myWidth + 10
                .Height = myHeight + 10
                .Left = 10 + (i - 1) * myWidth
                .Tag = k
                .BorderStyle = IIf(partsBorder, 1, 0)
                .Visible = True
                .AutoRedraw = False
            End With
            myPartsCollection.Add picParts(k)
        Next
    Next
Exit Sub
errHandler:
    MsgBox "Please close some of other puzzles and try again.", vbOKOnly + vbCritical, "Error"
End Sub

' enabling or disabling all of puzzle
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

' enabling or disabling all of puzzle
Public Property Let Enabled(ByVal newStatus As Boolean)
    UserControl.Enabled() = newStatus
End Property

'for small pic on the left buttom of the status bar
Public Property Get Thumbnail() As PictureBox
    Set Thumbnail = picSmall
End Property

' when you move (shuffle) a part of puzzle
Private Sub picParts_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    On Error GoTo errHandler ' for preventing one run time error: when program is shuffling, user try to close form.
    Dim lastLeft As Integer, lastTop As Integer, lastTag As Integer

    If Index = Source.Index Then Exit Sub ' if you move a part to the same location
    
    If Difficulty = Easy Then ' in easy mode, borders of picture stay without changes
'        If (Index + Rows - 1) Mod Rows = 0 Or (Source.Index + Rows - 1) Mod Rows = 0 Then Exit Sub ' for top row
        If Index Mod Rows = 0 Or Source.Index Mod Rows = 0 Then Exit Sub ' for bottom row
        If Index <= Rows Or Source.Index <= Rows Then Exit Sub ' for left column
'        If Index / Rows > Columns - 1 Or Source.Index / Rows > Columns - 1 Then Exit Sub ' for right column
    End If
    
    myUndoPic1 = Index
    myUndoPic2 = Source.Index
    
    lastLeft = picParts(Index).Left
    lastTop = picParts(Index).Top
    lastTag = picParts(Index).Tag
    
    picParts(Index).Left = Source.Left
    picParts(Index).Top = Source.Top
    picParts(Index).Tag = Source.Tag
    
    Source.Left = lastLeft
    Source.Top = lastTop
    Source.Tag = lastTag
    
    If itsReady Then
        myTotalMoves = myTotalMoves + 1
        myScore = myScore - 10
        RaiseEvent Counter(myTotalMoves)
        If CheckOK Then ' you finished puzzle
            tmrScore.Enabled = False
            itsReady = False
            RaiseEvent Finished
        End If
        If myTotalMoves = 1 Then tmrScore.Enabled = True 'score timer only start after first move
    End If
errHandler: 'just go out
End Sub

' this timer shuffle picture parts at the start of game
Private Sub tmrShuffle_Timer()
    Static myPartNo As Integer, myMax As Integer
    Dim j As Integer
    
    myMax = Columns * Rows ' I use this variable for determaine how many parts must change
    myPartNo = myPartNo + 1
    
    ' if you have a few parts, speed is slow but if you have many parts speed must be fast
    If tmrShuffle.Interval >= 500 Then tmrShuffle.Interval = IIf(myMax < 50, 200 - (myMax * 4), 1)
    
    If myPartNo <= myMax Then
        For j = 0 To 1 + Difficulty * 2 ' more shuffle depends on difficulty mode
            picParts_DragDrop Int(Rnd(1) * myMax) + 1, picParts(myPartNo), 0, 0
            DoEvents
        Next
    Else
        tmrShuffle.Enabled = False
        itsReady = True ' user can play now
        RaiseEvent Counter(myTotalMoves) 'reset number of moves and also enable puzzle
    End If
End Sub

' you lose your score by the time after first move in hard mode, also set your lost time and last score.
Private Sub tmrScore_Timer()
    Static LostTime As Long

    LostTime = LostTime + 1
    If LostTime Mod 5 = 0 And Difficulty = Hard Then myScore = myScore - 10
    If myScore < 0 Then myScore = 0
    RaiseEvent Score(myScore, LostTime)
End Sub

'picture is complete or not
Private Function CheckOK() As Boolean
    Dim myPart As Control
    
    CheckOK = True
    For Each myPart In myPartsCollection
        If myPart.Index <> Val(myPart.Tag) Then
            CheckOK = False
        End If
    Next
End Function

' for showing or hiding border of puzzle parts
Public Sub CheckBorder(myNewBorder As Boolean)
    Dim myPart As Control
    
    For Each myPart In myPartsCollection
        myPart.BorderStyle = IIf(myNewBorder, 1, 0)
    Next
End Sub

' undo last move
Public Sub Undo()
    Call picParts_DragDrop(myUndoPic1, picParts(myUndoPic2), 0, 0)
End Sub

'when user wants computer to move a part to correct area ( with losing of score), added in update
Public Sub Solve(ByVal Random As Boolean)
    Dim picNo As Integer
    
    If Random Then
        Do
            picNo = Int(Rnd(1) * Rows * Columns)
        Loop While Val(picParts(picNo).Tag) = picNo
        
        picParts_DragDrop picNo, picParts(Val(picParts(picNo).Tag)), 0, 0
        myScore = myScore - 100 * (Difficulty + 1) + 10
    Else
        solveMe = True ' for manual solving by right click on part
    End If
End Sub

'when you choose Manual solve mode in menu.
Private Sub picParts_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If solveMe Then
        If Val(picParts(Index).Tag) <> Index Then ' this picture is not at correct location
            solveMe = False
            picParts_DragDrop Val(picParts(Index).Tag), picParts(Index), 0, 0
            myScore = myScore - 250 * (Difficulty + 1) + 10
        End If
    End If
End Sub

'when user wants to see full picture for limited time ( with losing of score), added in update
Public Sub showFullPic()
    picShow.Left = 0
    picShow.Top = 0
    picShow.Width = picParts(1).Width * Columns
    picShow.Height = picParts(1).Height * Rows
    picShow.PaintPicture picSource.Picture, 0, 0, picShow.Width, picShow.Height, 0, 0, picSource.Width, picSource.Height
    picShow.Visible = True
    tmrShowFullPic.Enabled = True
End Sub

'after a limited time of viewing full picture, you must back to game
Private Sub tmrShowFullPic_Timer()
    tmrShowFullPic.Enabled = False
    myScore = myScore - 250 * (Difficulty + 1)
    picShow.Visible = False
End Sub
' when user wants to see correct parts or uncorrect parts ( with losing of score), added in update
Public Sub showParts(ByVal correctType As Boolean)
    Dim myPart As Control
    
    For Each myPart In myPartsCollection
        myPart.Visible = IIf(Val(myPart.Tag) = myPart.Index, correctType, Not correctType)
    Next
    tmrShowCorrect.Enabled = True
End Sub

'after a limited time of viewing correct( or uncorrect) parts, you must back to game
Private Sub tmrShowCorrect_Timer()
    Dim myPart As Control
    
    tmrShowCorrect = False
    For Each myPart In myPartsCollection
        myPart.Visible = True
    Next
    myScore = myScore - 125 * (Difficulty + 1)
End Sub

' when you close form , this procedure remove objects from collection and unload them from memory
Public Sub UnloadAllObjects()
    Dim myPart As Control
    
    For Each myPart In myPartsCollection
        myPartsCollection.Remove 1
        Unload myPart
    Next
End Sub


