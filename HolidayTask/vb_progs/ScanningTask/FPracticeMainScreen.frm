VERSION 5.00
Begin VB.Form FPracticeMainScreen 
   BorderStyle     =   0  'None
   Caption         =   "Practice Screen"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MouseIcon       =   "FPracticeMainScreen.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "(Practice Sessions)"
      BeginProperty Font 
         Name            =   "BernhardMod BT"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   5220
      TabIndex        =   2
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Image imgPalm 
      Height          =   1920
      Left            =   10320
      Picture         =   "FPracticeMainScreen.frx":0152
      Top             =   6360
      Width           =   1650
   End
   Begin VB.Image imgScuba 
      Height          =   2595
      Left            =   9000
      Picture         =   "FPracticeMainScreen.frx":0C3D
      Top             =   8520
      Width           =   1290
   End
   Begin VB.Image imgIsland 
      Height          =   2925
      Left            =   120
      Picture         =   "FPracticeMainScreen.frx":45F1
      Top             =   8520
      Width           =   3750
   End
   Begin VB.Image imgCartoon 
      Height          =   3195
      Left            =   4680
      Picture         =   "FPracticeMainScreen.frx":77B4
      Top             =   6240
      Width           =   3000
   End
   Begin VB.Image imgHike 
      Height          =   3000
      Left            =   720
      Picture         =   "FPracticeMainScreen.frx":9243
      Top             =   4800
      Width           =   2160
   End
   Begin VB.Label lblCompanyTagLine 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Helping you find your ideal break"
      BeginProperty Font 
         Name            =   "BernhardMod BT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   4740
      Width           =   6495
   End
   Begin VB.Image imgCompanyLogo 
      Height          =   2490
      Left            =   6120
      Picture         =   "FPracticeMainScreen.frx":A3F6
      Top             =   2220
      Width           =   3150
   End
   Begin VB.Image imgSurf 
      Height          =   2400
      Left            =   840
      Picture         =   "FPracticeMainScreen.frx":B580
      Top             =   1800
      Width           =   3765
   End
   Begin VB.Image imgSun 
      Height          =   3510
      Left            =   11160
      Picture         =   "FPracticeMainScreen.frx":C703
      Top             =   1320
      Width           =   3750
   End
   Begin VB.Label lblHolidayTaskTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "The Holiday Store"
      BeginProperty Font 
         Name            =   "BernhardMod BT"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   11055
   End
   Begin VB.Image imgPoolside 
      Height          =   2745
      Left            =   11640
      Picture         =   "FPracticeMainScreen.frx":E4AE
      Top             =   6240
      Width           =   3375
   End
End
Attribute VB_Name = "FPracticeMainScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'internal flag
Private bGotSpacePress As Boolean

Private Sub Form_Load()
    Me.BackColor = vbBlack  'in case it is not the default
    If bShowMouse Then Me.MousePointer = 0
    prepareForm
End Sub

Public Sub prepareForm()
'NB this Sub is Public so that other forms can use it
    bGotSpacePress = False
    lblCompanyTagLine.Caption = strCompanyTagLine
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If bGotSpacePress Then Exit Sub
    If KeyAscii = asciiSpaceBar Then
        bGotSpacePress = True
        If (FPracticeTaskInfoScreen.iPracticeItemSet = UBound(PracticeMenuSet)) Then
            FPracticeScreen.Show
            Me.Hide
        Else
            lblCompanyTagLine.Caption = strGetReady
            Call Wait(mainScreenDelayTime)
            FPracticeTaskInfoScreen.prepareForm
            FPracticeTaskInfoScreen.Show
            Me.Hide
        End If
    End If
    'If user hits the "Escape" key then jump right back to FPracticeScreen
    If KeyAscii = asciiBackToFirstForm Then
        FPracticeScreen.Show
        Me.Hide
    End If
End Sub
