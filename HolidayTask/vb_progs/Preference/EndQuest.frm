VERSION 5.00
Begin VB.Form FEndQuest 
   BorderStyle     =   0  'None
   Caption         =   "End Questions"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   Icon            =   "EndQuest.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7073
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7073
      TabIndex        =   2
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox txtAnswer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4673
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "EndQuest.frx":030A
      ToolTipText     =   "Please type your answer here"
      Top             =   4920
      Width           =   6015
   End
   Begin VB.Image imgQuit 
      Height          =   975
      Left            =   14160
      Top             =   10440
      Width           =   1095
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Question in Here"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   855
      Left            =   3720
      TabIndex        =   0
      Top             =   3840
      Width           =   7935
   End
End
Attribute VB_Name = "FEndQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this is the file where all the questions are stored
Private Const NameOfQuestionFile = "questionList"

Private AnswerFile, h_AnswerFile, questionList, h_questionList As String
Private iQuestionNumber, iQuestionArrayMax As Integer

Private Sub Form_Load()
    Me.BackColor = vbBlack      'in case it's not the default
    
    questionList = NameOfQuestionFile
    h_questionList = FreeFile
    Open questionList For Input Access Read As #h_questionList
    Do While Not EOF(h_questionList)
        iQuestionArrayMax = iQuestionArrayMax + 1
        ReDim Preserve questionArray(1 To iQuestionArrayMax) 'resize array to input new item
        Input #h_questionList, questionArray(iQuestionArrayMax)
    Loop
    iQuestionNumber = LBound(questionArray)
    
    AnswerFile = Replace(FStartScreen.outputFile, "Ratings.csv", "Answers")
    h_AnswerFile = FreeFile
    Open AnswerFile For Output Access Write Lock Read Write As #h_AnswerFile
    Print #h_AnswerFile, Format(Now, "hh:mm:ss, dd mmmm yyyy")
    
    nextQ
End Sub

Private Sub nextQ()
    txtAnswer.Text = ""
    lblQuestion.Caption = questionArray(iQuestionNumber)
End Sub

Private Sub cmdNext_Click()
    Print #h_AnswerFile, lblQuestion.Caption
    Print #h_AnswerFile, txtAnswer.Text
    Print #h_AnswerFile, ""
    
    iQuestionNumber = iQuestionNumber + 1
    If (iQuestionNumber > UBound(questionArray)) Then finalPage Else nextQ
        
End Sub

Private Sub finalPage()
    txtAnswer.Visible = False
    cmdNext.Visible = False
    cmdClose.Visible = True
    Print #h_AnswerFile, "End Time", Format(Now, "hh:mm:ss, dd mmmm yyyy")
    Close #h_AnswerFile
    lblQuestion.Caption = "Thank you for your participation"
End Sub

Private Sub cmdClose_Click()
    End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If user hits the "Escape" key then give choice to exit
    If KeyAscii = 27 Then checkBeforeEnding    ' 27="Escape" key
End Sub

'Private Sub imgQuit_Click()
'    End
'End Sub
