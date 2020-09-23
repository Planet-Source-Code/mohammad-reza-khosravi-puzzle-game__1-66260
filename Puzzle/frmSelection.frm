VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Your Favorite"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2895
   Icon            =   "frmSelection.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   193
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.UpDown updColumns 
      Height          =   315
      Left            =   2490
      TabIndex        =   10
      Top             =   1860
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Value           =   4
      BuddyControl    =   "txtColumns"
      BuddyDispid     =   196609
      OrigLeft        =   166
      OrigTop         =   110
      OrigRight       =   183
      OrigBottom      =   141
      Max             =   30
      Min             =   2
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtColumns 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2070
      TabIndex        =   9
      Text            =   "4"
      Top             =   1860
      Width           =   405
   End
   Begin VB.TextBox txtRows 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1140
      TabIndex        =   8
      Text            =   "3"
      Top             =   1860
      Width           =   405
   End
   Begin MSComCtl2.UpDown updRows 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   1860
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Value           =   3
      BuddyControl    =   "txtRows"
      BuddyDispid     =   196610
      OrigLeft        =   44
      OrigTop         =   118
      OrigRight       =   61
      OrigBottom      =   147
      Max             =   30
      Min             =   2
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.ComboBox cmbDifficulty 
      Height          =   315
      ItemData        =   "frmSelection.frx":0442
      Left            =   1140
      List            =   "frmSelection.frx":044F
      TabIndex        =   6
      Text            =   "Standard"
      Top             =   2250
      Width           =   1635
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2520
      Top             =   570
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   450
      Left            =   180
      TabIndex        =   1
      Top             =   2670
      Width           =   1125
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   450
      Left            =   1590
      TabIndex        =   0
      Top             =   2670
      Width           =   1125
   End
   Begin VB.PictureBox PicTemp 
      AutoSize        =   -1  'True
      Height          =   390
      Left            =   2550
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label Label1 
      Caption         =   "Difficulty :"
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   2280
      Width           =   1035
   End
   Begin VB.Image imgChoose 
      BorderStyle     =   1  'Fixed Single
      Height          =   1545
      Left            =   345
      Picture         =   "frmSelection.frx":0469
      Stretch         =   -1  'True
      Tag             =   "Puzzle"
      Top             =   150
      Width           =   2205
   End
   Begin VB.Label Label2 
      Caption         =   "X"
      Height          =   255
      Left            =   1890
      TabIndex        =   3
      Top             =   1920
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "Puzzle Parts :"
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   1890
      Width           =   1035
   End
End
Attribute VB_Name = "frmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'change picture from file
Private Sub cmdBrowse_Click()
    On Error GoTo errHandler
    Dim myScale As Single
    
    CD1.Filter = "Pictures (*.jpg,*.gif,*.bmp)|*.jpg;*.gif;*.bmp"
    CD1.ShowOpen
    If CD1.FileName <> "" Then
        With PicTemp
            .Picture = LoadPicture(CD1.FileName)
            myScale = .Width / .Height
        End With
        With imgChoose ' setting size and position of new image on Form
            .Picture = PicTemp.Picture
            .Tag = CD1.FileName
            If myScale > 1 Then
                .Width = Me.ScaleWidth - Me.ScaleWidth / 4
                .Height = .Width / myScale
                .Left = 24
                .Top = (Me.ScaleHeight / 1.8 - .Height) / 2
            Else
                .Height = Me.ScaleHeight / 2
                .Width = .Height * myScale
                .Top = 7
                .Left = (Me.ScaleWidth - .Width) / 2
            End If
        End With
    End If
Exit Sub
errHandler:
    MsgBox "Error in reading File " & CD1.FileName, vbExclamation, "Error"
End Sub

'make a new copy of frmMain and set it for playing new puzzle
Private Sub cmdPlay_Click()
    Dim myNewForm As New frmMain
    
    Me.Hide
    DoEvents
    With myNewForm.myPuzzle
        .Enabled = False
        .mainPicture imgChoose.Picture
        .Columns = updColumns.Value
        .Rows = updRows.Value
        .Difficulty = Abs(cmbDifficulty.ListIndex) 'must be positive
    End With
    myNewForm.Show
End Sub

