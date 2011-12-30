VERSION 5.00
Begin VB.Form FPracticeScreen 
   BorderStyle     =   0  'None
   Caption         =   "PracticeScreen"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "FPracticeScreen.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPracticeSession 
      Caption         =   "Practice Session"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5513
      TabIndex        =   1
      Top             =   5520
      Width           =   4335
   End
   Begin VB.CommandButton cmdRealSessions 
      Caption         =   "Real Sessions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5513
      TabIndex        =   2
      Top             =   7440
      Width           =   4335
   End
   Begin VB.CommandButton cmdScreenCheck 
      Caption         =   "Screen Check"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5513
      TabIndex        =   0
      Top             =   3600
      Width           =   4335
   End
End
Attribute VB_Name = "FPracticeScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.BackColor = vbBlack
    If bShowMouse Then Me.MousePointer = 0
'    cmdPracticeSession.Enabled = False
'    cmdRealSessions.Enabled = False
End Sub

Private Sub cmdScreenCheck_Click()
    FPracticeTouch.prepareForm
    FPracticeTouch.Show
    cmdPracticeSession.Enabled = True
    Me.Hide
End Sub

Private Sub cmdPracticeSession_Click()
    'Reset all counters used on other practice forms now
    '(just in case user aborted the trials and needs to start again)
    FPracticeTaskInfoScreen.iPracticeItemSet = 0
'    FPracticeTaskScans.iChoicePage = 1
    FPracticeMainScreen.prepareForm
    FPracticeMainScreen.Show
    Me.Hide
    cmdRealSessions.Enabled = True
End Sub

Private Sub cmdRealSessions_Click()
    'load the MAIN expt forms
    Load FMainScreen
    Load FTaskInfoScreen
    Load FHolidayTaskScans
    'unload Practice Forms
    Unload FPracticeMainScreen
    Unload FPracticeTouch
    
    'Reset all counters used on other forms now
    '(just in case user aborted the trials and needs to start again)
    FTaskInfoScreen.iItemSet = 0
    FHolidayTaskScans.iChoicePage = 1
    FMainScreen.prepareForm
    FMainScreen.Show
    Me.Hide
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = asciiBackToFirstForm Then End
End Sub

