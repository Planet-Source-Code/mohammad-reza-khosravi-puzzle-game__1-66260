VERSION 5.00
Begin VB.Form frmFinish 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "You Win !"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
   Icon            =   "frmFinish.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2250
      TabIndex        =   12
      Top             =   3000
      Width           =   1185
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1590
      MaxLength       =   30
      TabIndex        =   0
      Top             =   2550
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   690
      TabIndex        =   1
      Top             =   3000
      Width           =   1185
   End
   Begin VB.Label lblFinish 
      Caption         =   "Difficulty :"
      Height          =   285
      Index           =   6
      Left            =   180
      TabIndex        =   14
      Top             =   1150
      Width           =   1335
   End
   Begin VB.Label LblResult 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   1590
      TabIndex        =   13
      Top             =   2200
      Width           =   2295
   End
   Begin VB.Label LblResult 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   1590
      TabIndex        =   11
      Top             =   1850
      Width           =   2295
   End
   Begin VB.Label LblResult 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   1590
      TabIndex        =   10
      Top             =   1500
      Width           =   2295
   End
   Begin VB.Label LblResult 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1590
      TabIndex        =   9
      Top             =   1150
      Width           =   2295
   End
   Begin VB.Label LblResult 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1590
      TabIndex        =   8
      Top             =   800
      Width           =   2295
   End
   Begin VB.Label lblFinish 
      Caption         =   "Puzzle :"
      Height          =   285
      Index           =   5
      Left            =   180
      TabIndex        =   7
      Top             =   800
      Width           =   1335
   End
   Begin VB.Label lblFinish 
      Caption         =   "Enter Your Name :"
      Height          =   285
      Index           =   4
      Left            =   150
      TabIndex        =   6
      Top             =   2550
      Width           =   1335
   End
   Begin VB.Label lblFinish 
      Caption         =   "Your Score :"
      Height          =   345
      Index           =   3
      Left            =   150
      TabIndex        =   5
      Top             =   2200
      Width           =   1335
   End
   Begin VB.Label lblFinish 
      Caption         =   "Your Moves :"
      Height          =   285
      Index           =   2
      Left            =   150
      TabIndex        =   4
      Top             =   1850
      Width           =   1335
   End
   Begin VB.Label lblFinish 
      Caption         =   "Your Time :"
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   1500
      Width           =   1335
   End
   Begin VB.Label lblFinish 
      Alignment       =   2  'Center
      Caption         =   "Congratulation !"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   645
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   120
      Width           =   3915
   End
End
Attribute VB_Name = "frmFinish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' I used another way for connecting to database and saving data

Private Sub cmdOk_Click()
    On Error GoTo errHandler
    Dim myCnn As New Connection
    Dim myRst As New Recordset
    
    myCnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Puzzle.mdb"
    myCnn.Open
    
    With myRst
        .Open "Select * from Main", myCnn, adOpenKeyset, adLockOptimistic
        .AddNew
        !Puzzle = LblResult(0).Caption & ""
        !Difficulty = LblResult(1).Caption & ""
        !Time = LblResult(2).Caption & ""
        !Moves = Val(LblResult(3).Caption & "")
        !Score = Val(LblResult(4).Caption & "")
        !Name = txtName.Text & ""
        .Update
        .Close
    End With
    
    myCnn.Close
    Set myRst = Nothing
    Set myCnn = Nothing
    
    Unload Me
Exit Sub
errHandler:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


