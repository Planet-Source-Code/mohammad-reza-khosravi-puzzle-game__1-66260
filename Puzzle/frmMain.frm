VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00400000&
   Caption         =   "Puzzle"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   5160
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   5160
   Begin MSComctlLib.StatusBar stbPuzzle 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   1
      Top             =   3990
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   926
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1076
            MinWidth        =   1076
            Key             =   "SmallPic"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Visible         =   0   'False
            Object.Width           =   952
            MinWidth        =   952
            Picture         =   "frmMain.frx":0442
            Key             =   "Finish"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "00:00:00"
            TextSave        =   "00:00:00"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Key             =   "Parts"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
            Text            =   "Moves :"
            TextSave        =   "Moves :"
            Key             =   "Moves"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "Score : "
            TextSave        =   "Score : "
            Key             =   "Score"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Key             =   "Difficulty"
            Object.ToolTipText     =   "Difficulty Mode"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Key             =   "Border"
            Object.ToolTipText     =   "Click for changing border of puzzles"
         EndProperty
      EndProperty
   End
   Begin PuzzleGame.Puzzle myPuzzle 
      Height          =   3285
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   5794
   End
   Begin VB.Menu mnuFileGroup 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuScores 
         Caption         =   "Scores"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuLine0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuEditAll 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuBorder 
         Caption         =   "Border"
         Checked         =   -1  'True
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnuMainSolve 
      Caption         =   "&Solve"
      Begin VB.Menu mnuSolve 
         Caption         =   "Random "
         Index           =   0
      End
      Begin VB.Menu mnuSolve 
         Caption         =   "Manual "
         Index           =   1
      End
      Begin VB.Menu mnuline 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show Full Picture"
         Index           =   0
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show Correct Parts"
         Index           =   1
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show Uncorrect Parts"
         Index           =   2
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub Form_Load()
    Me.Caption = frmSelection.imgChoose.Tag
End Sub

Private Sub Form_Activate()
    Static itsReady As Boolean
    
    If Not itsReady Then ' I want to run this part only one time at the beginning
        ' start of new game
        sndPlaySound App.Path & "\start.wav", 1 ' play starting sound
        myPuzzle.Start MDIPuzzle.Width * 3 / 4, MDIPuzzle.Height * 3 / 4, mnuBorder.Checked
        stbPuzzle.Panels("Border").Text = IIf(mnuBorder.Checked, "Border", "No Border")
        stbPuzzle.Panels("Parts") = myPuzzle.Rows & " X " & myPuzzle.Columns
        stbPuzzle.Panels("Difficulty").Text = Choose(myPuzzle.Difficulty + 1, "Easy", "Standard", "Hard")
        stbPuzzle.Panels("SmallPic").Picture = myPuzzle.Thumbnail.Image
        ' menu items are different, depends on difficulty mode
        If myPuzzle.Difficulty = Hard Then mnuShow(0).Enabled = False: mnuShow(1).Enabled = False: mnuShow(2).Enabled = False
        mnuSolve(0).Caption = "Random (-" & 100 * (myPuzzle.Difficulty + 1) & " score)"
        mnuSolve(1).Caption = "Manual (-" & 250 * (myPuzzle.Difficulty + 1) & " score) [use Right Click]"
        mnuShow(0).Caption = "Show Full Picture (-" & 250 * (myPuzzle.Difficulty + 1) & " score)"
        mnuShow(1).Caption = "Show Correct Parts (-" & 125 * (myPuzzle.Difficulty + 1) & " score)"
        mnuShow(2).Caption = "Show Uncorrect Parts (-" & 125 * (myPuzzle.Difficulty + 1) & " score)"
        itsReady = True
    End If
End Sub

Private Sub myPuzzle_Resize(ByVal Width As Integer, ByVal Height As Integer)
    myPuzzle.Width = Width
    myPuzzle.Height = Height
    Me.Width = Width + 230
    Me.Height = Height + 1150
End Sub

' show number of your moves
Private Sub myPuzzle_Counter(ByVal NumberOfMoves As Long)
    If NumberOfMoves = 0 Then myPuzzle.Enabled = True: Me.BackColor = vbButtonFace ' activate puzzle
    stbPuzzle.Panels("Moves").Text = "Moves : " & NumberOfMoves
End Sub

'score is depended on your moves and also lost time.
Private Sub myPuzzle_Score(ByVal TotalScore As Long, ByVal LostTime As Long)
    stbPuzzle.Panels("Time").Text = Format(LostTime \ 3600, "00") & ":" & Format((LostTime Mod 3600) \ 60, "00") & ":" & Format(LostTime Mod 60, "00")
    stbPuzzle.Panels("Score").Text = "Score : " & TotalScore
End Sub

' when you finish a puzzle, this procedure activate.
Private Sub myPuzzle_Finished()
    stbPuzzle.Panels("SmallPic").Visible = False
    stbPuzzle.Panels("Finish").Visible = True
    sndPlaySound App.Path & "\Finish.wav", 1 ' play finish sound
    frmFinish.LblResult(0).Caption = stbPuzzle.Panels("Parts").Text
    frmFinish.LblResult(1).Caption = stbPuzzle.Panels("Difficulty").Text
    frmFinish.LblResult(2).Caption = stbPuzzle.Panels("Time").Text
    frmFinish.LblResult(3).Caption = Mid(stbPuzzle.Panels("Moves").Text, 9)
    frmFinish.LblResult(4).Caption = Mid(stbPuzzle.Panels("Score").Text, 9)
    frmFinish.Show vbModal 'for showing statistics and saving your score in database
    myPuzzle.Enabled = False 'lock puzzle
End Sub

' you can show or hide border of each part of puzzle, by clicking on panel ( also by menu).
Private Sub StbPuzzle_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "Border" Then Call mnuBorder_Click
End Sub

Private Sub mnuBorder_Click()
    mnuBorder.Checked = Not mnuBorder.Checked
    myPuzzle.CheckBorder mnuBorder.Checked
    stbPuzzle.Panels("Border").Text = IIf(mnuBorder.Checked, "Border", "No Border")
End Sub

Private Sub mnuNew_Click()
    frmSelection.Show vbModal
End Sub

Private Sub mnuScores_Click()
    frmScores.Show vbModal
End Sub

Private Sub mnuUndo_Click()
    If myPuzzle.Enabled Then myPuzzle.Undo
End Sub

'added in update
Private Sub mnuSolve_Click(Index As Integer)
    If myPuzzle.Enabled Then myPuzzle.Solve (IIf(Index = 0, True, False))
End Sub

'added in update
Private Sub mnuShow_Click(Index As Integer)
    If myPuzzle.Enabled Then
        Select Case Index
            Case 0
                myPuzzle.showFullPic
            Case 1, 2
                myPuzzle.showParts (IIf(Index = 1, True, False))
        End Select
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    myPuzzle.UnloadAllObjects
End Sub

