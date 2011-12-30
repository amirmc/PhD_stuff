VERSION 5.00
Begin VB.Form FPracticeTaskInfoScreen 
   BorderStyle     =   0  'None
   Caption         =   "Task Info Screen"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MouseIcon       =   "FPracticeTaskInfoScreen.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Instructions"
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
      Height          =   1335
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   11055
   End
   Begin VB.Image imgSun 
      Height          =   3510
      Left            =   10680
      Picture         =   "FPracticeTaskInfoScreen.frx":0152
      Top             =   4485
      Width           =   3750
   End
   Begin VB.Label lblTaskInstructions 
      BackColor       =   &H00000000&
      Caption         =   "lblTaskInstructions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   5415
      Left            =   1080
      TabIndex        =   0
      Top             =   3780
      Width           =   8895
   End
End
Attribute VB_Name = "FPracticeTaskInfoScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''
''  There are a set of Text Strings used in this form
''  that can be found in PETScanningModule
''

Public iPracticeItemSet As Integer  'this is public because the writeToFile
'subroutine in FPracticeHolidayTaskScans needs it to print the choices to file
'and it is used in If statements in FPracticeHolidayTaskScans and FPracticeMainScreen

'internal flags
Private bGotSpacePress As Boolean
Public bDecisionTrial As Boolean
Public bNonAffectiveTrial As Boolean
'the last two are public because FPracticeTaskScans might need them.

Private Sub Form_Load()
    Me.BackColor = vbBlack  'in case it is not the default
    If bShowMouse Then Me.MousePointer = 0
End Sub

Public Sub prepareForm()
'NB this Sub is Public so that other forms can use it
    bGotSpacePress = False
    Call showAll(False)
    Call setupNextTrial
    Call showAll(True)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If bGotSpacePress = True Then Exit Sub
    If KeyAscii = asciiSpaceBar Then
        bGotSpacePress = True
        Call showAll(False)
        Call Wait(betweenPageDelayTime)
        FPracticeTaskScans.prepareForm
        FPracticeTaskScans.Show
        Me.Hide
    End If
    'If user hits the "Escape" key then jump right back to FPracticeScreen
    If KeyAscii = asciiBackToFirstForm Then
        FPracticeScreen.Show
        Me.Hide
    End If
End Sub

Private Sub setupNextTrial()
'ALL forms should have been loaded at end of FStartScreen
'so calling other forms should not be a problem

    ' this increment is important and should ONLY
    ' be altered by the program as THIS point
    iPracticeItemSet = iPracticeItemSet + 1
    
    Call setBooleanDecisionValues
    Call setTaskInstructions
End Sub

Private Sub setBooleanDecisionValues()
'This Sub looks at the info taken from the input file and then sets
'the boolean values used within this form
    
    Dim iFirstLine As Integer, strCondition As String
    iFirstLine = LBound(PracticeMenuSet(iPracticeItemSet).Title)
    
    'to make the following code in this module easier to read
    strCondition = PracticeMenuSet(iPracticeItemSet).Condition(iFirstLine)
    
    If (strCondition = strHighDecision) Or (strCondition = strLowDecision) Then
        bDecisionTrial = True
        bNonAffectiveTrial = False
    ElseIf (strCondition = strHighNoDecision) Or (strCondition = strLowNoDecision) Then
        bDecisionTrial = False
        bNonAffectiveTrial = False
    ElseIf (strCondition = strHighNonAffective) Or (strCondition = strLowNonAffective) Then
        bDecisionTrial = True
        bNonAffectiveTrial = True
    Else
        MsgBox ("Having trouble with InputFile in 'setBooleanDecisionValues'")
        End
    End If
End Sub

Private Sub setTaskInstructions()
'This bit uses the boolean values from setBooleanDecisionValues() to
'put the appropriate instructions for the subject onto the screen
'
'This bit of the code also sets "FpracticeTaskScans.lblTaskInstruction"
    
    Dim v As Variant
    If bDecisionTrial Then
        For Each v In FPracticeTaskScans.lblLink
            v.Caption = strDLink
        Next
        If bNonAffectiveTrial Then
            lblTaskInstructions.Caption = strNAinstr
            FPracticeTaskScans.lblTaskInstruction.Caption = strNAtrial
        Else
            lblTaskInstructions.Caption = strDinstr
            FPracticeTaskScans.lblTaskInstruction.Caption = strDtrial
        End If
    Else
        For Each v In FPracticeTaskScans.lblLink
            v.Caption = strNDLink
        Next
        lblTaskInstructions.Caption = strNDinstr
        FPracticeTaskScans.lblTaskInstruction.Caption = strNDtrial
    End If
End Sub

Private Sub showAll(ByVal bShow As Boolean)
    imgSun.Visible = bShow
    lblTitle.Visible = bShow
    lblTaskInstructions.Visible = bShow
End Sub

