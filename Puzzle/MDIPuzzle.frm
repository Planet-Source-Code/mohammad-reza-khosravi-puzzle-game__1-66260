VERSION 5.00
Begin VB.MDIForm MDIPuzzle 
   BackColor       =   &H8000000C&
   Caption         =   "Puzzle Game"
   ClientHeight    =   5085
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7620
   Icon            =   "MDIPuzzle.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFileGroup 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuScores 
         Caption         =   "Scores"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
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
Attribute VB_Name = "MDIPuzzle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'       Puzzle Game
'       Copyright (c) 2006 Mohammad Reza Khosravi ( Khosravi2500@yahoo.com )
'
'
'       Update September 30:
'           Added Difficulty Mode
'           Added Solve Menu
'

Private Sub MDIForm_Activate()
    Static newGame As Boolean
    If Not newGame Then ' when activated for the first time, show a selection form for new game
        newGame = True
        Call mnuNew_Click
    End If
End Sub

' when you double click on MDI form, new puzzle selection appear ( like Photoshop! )
Private Sub MDIForm_DblClick()
    frmSelection.Show vbModal
End Sub

Private Sub mnuNew_Click()
    frmSelection.Show vbModal
End Sub

Private Sub mnuScores_Click()
    frmScores.Show vbModal
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
    End
End Sub


