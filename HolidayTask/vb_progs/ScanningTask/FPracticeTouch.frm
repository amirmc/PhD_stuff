VERSION 5.00
Begin VB.Form FPracticeTouch 
   BorderStyle     =   0  'None
   Caption         =   "TouchScreen Practice"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MouseIcon       =   "FPracticeTouch.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   15
      Left            =   11040
      Picture         =   "FPracticeTouch.frx":0152
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   14
      Left            =   13800
      Picture         =   "FPracticeTouch.frx":A710
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   13
      Left            =   10800
      Picture         =   "FPracticeTouch.frx":14CCE
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   12
      Left            =   11640
      Picture         =   "FPracticeTouch.frx":1F28C
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   11
      Left            =   1320
      Picture         =   "FPracticeTouch.frx":2984A
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   10
      Left            =   840
      Picture         =   "FPracticeTouch.frx":33E08
      Stretch         =   -1  'True
      Top             =   600
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   9
      Left            =   4680
      Picture         =   "FPracticeTouch.frx":3E3C6
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   8
      Left            =   5280
      Picture         =   "FPracticeTouch.frx":48984
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   7
      Left            =   8520
      Picture         =   "FPracticeTouch.frx":52F42
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   6
      Left            =   8400
      Picture         =   "FPracticeTouch.frx":5D500
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   5
      Left            =   4680
      Picture         =   "FPracticeTouch.frx":67ABE
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   4
      Left            =   10800
      Picture         =   "FPracticeTouch.frx":7207C
      Stretch         =   -1  'True
      Top             =   360
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   3
      Left            =   720
      Picture         =   "FPracticeTouch.frx":7C63A
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   2
      Left            =   5760
      Picture         =   "FPracticeTouch.frx":86BF8
      Stretch         =   -1  'True
      Top             =   480
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   1
      Left            =   13680
      Picture         =   "FPracticeTouch.frx":911B6
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   735
   End
   Begin VB.Image imgTarget 
      Height          =   735
      Index           =   0
      Left            =   8640
      Picture         =   "FPracticeTouch.frx":9B774
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   735
   End
End
Attribute VB_Name = "FPracticeTouch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sStartTime As Single

Private Sub Form_Load()
    Me.BackColor = vbBlack
    If bShowMouse Then Me.MousePointer = 0
End Sub

Public Sub prepareForm()
    Dim v As Variant
    For Each v In imgTarget
        v.Visible = False
    Next
    imgTarget(imgTarget.LBound).Visible = True
'    sStartTime = Timer
End Sub

Private Sub imgTarget_Click(Index As Integer)
    If (Index = imgTarget.UBound) Then 'Or (Timer - sStartTime > 45) Then
        FPracticeScreen.Show
        Me.Hide
        Exit Sub
    End If
    imgTarget(Index).Visible = False
    imgTarget(Index + 1).Visible = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'mimic completion of the trial in order to go back to First Form
    If KeyAscii = asciiBackToFirstForm Then imgTarget_Click (imgTarget.UBound)
End Sub
