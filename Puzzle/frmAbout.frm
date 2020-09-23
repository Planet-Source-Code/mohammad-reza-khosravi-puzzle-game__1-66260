VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1553
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      Caption         =   "khosravi2500@yahoo.com"
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   1230
      Width           =   3885
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      Caption         =   "Copyright 2006 Mohammad Reza Khosravi"
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   900
      Width           =   3885
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ver 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      Top             =   390
      Width           =   1335
   End
   Begin VB.Label LblAbout 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Puzzle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3945
   End
   Begin VB.Shape Shape1 
      Height          =   2145
      Left            =   30
      Top             =   30
      Width           =   4005
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub
