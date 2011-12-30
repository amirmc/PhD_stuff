VERSION 5.00
Begin VB.Form FTaskInfoScreen 
   BorderStyle     =   0  'None
   Caption         =   "Task Instructions"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "FTaskInfoScreen.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      TabIndex        =   1
      Top             =   3780
      Width           =   8895
   End
   Begin VB.Image imgSun 
      Height          =   3510
      Left            =   10680
      Picture         =   "FTaskInfoScreen.frx":0152
      Top             =   4485
      Width           =   3750
   End
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
      TabIndex        =   0
      Top             =   360
      Width           =   11055
   End
End
Attribute VB_Name = "FTaskInfoScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''
''  There are a set of Text Strings used in this form
''  that can be found in PETScanningModule
''

Public iItemSet As Integer  'this is public because the writeToFile
'subroutine in FHolidayTaskScans needs it to print the choices to file
'and it is used in If statements in FHolidayTaskScans and FMainScreen

'internal flags
Private bGotSpacePress As Boolean
Public bDecisionTrial As Boolean
Public bNonAffectiveTrial As Boolean
'the last two are public because FHolidayTaskScans might need them.

Private Sub Form_Load()
    Me.BackColor = vbBlack  'in case it is not the default
    If bShowMouse Then Me.MousePointer = 0
'    iItemSet = FStartScreen.iStartMenu - 1
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
        FHolidayTaskScans.prepareForm
        FHolidayTaskScans.Show
        Me.Hide
    End If
    'If user hits the "Escape" key then jump right back to FPracticeScreen
    If KeyAscii = asciiBackToFirstForm Then
        Load FPracticeScreen
        FPracticeScreen.Show
        Me.Hide
        Unload FHolidayTaskScans
        Unload FTaskInfoScreen
        Unload Me
    End If
End Sub

Private Sub setupNextTrial()
'ALL forms should have been loaded at end of FStartScreen
'so calling other forms should not be a problem

    ' this increment is important and should ONLY
    ' be altered by the program as THIS point.
    '    NB if the user needs to change the trial number on
    '    the fly, then they can alter it via FMainScreen
    '    (see cmdSetStartMenu_Click and lblProgress_Click)
    iItemSet = iItemSet + 1
    
    ' this is to 'reset' the ChoicePage number so that each iItemSet
    ' starts from the first page (see code on FHolidayTaskScans)
    FHolidayTaskScans.iChoicePage = 0
    
    Call setBooleanDecisionValues
    Call setTaskInstructions
End Sub

Private Sub setBooleanDecisionValues()
'This Sub looks at the info taken from the input file and then sets
'the boolean values used within this form
    
    Dim iFirstPage As Integer, iFirstLine As Integer, strCondition As String
    iFirstPage = LBound(MenuSet(iItemSet).MenuPage)
    iFirstLine = LBound(MenuSet(iItemSet).MenuPage(iFirstPage).Title)
    
    'to make the following code in this module easier to read
    strCondition = MenuSet(iItemSet).MenuPage(iFirstPage).Condition(iFirstLine)
    
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
        MsgBox ("Having trouble with InputFile. Check values in 'Condition' Column")
        End
    End If
End Sub

Private Sub setTaskInstructions()
'This bit uses the boolean values from setBooleanDecisionValues() to
'put the appropriate instructions for the subject onto the screen
'
'This bit of the code also sets "FHolidayTaskScans.lblTaskInstruction"
    
    Dim v As Variant
    If bDecisionTrial Then
        For Each v In FHolidayTaskScans.lblLink
            v.Caption = strDLink
        Next
        If bNonAffectiveTrial Then
            lblTaskInstructions.Caption = strNAinstr
            FHolidayTaskScans.lblTaskInstruction.Caption = strNAtrial
        Else
            lblTaskInstructions.Caption = strDinstr
            FHolidayTaskScans.lblTaskInstruction.Caption = strDtrial
        End If
    Else
        For Each v In FHolidayTaskScans.lblLink
            v.Caption = strNDLink
        Next
        lblTaskInstructions.Caption = strNDinstr
        FHolidayTaskScans.lblTaskInstruction.Caption = strNDtrial
    End If
End Sub

Private Sub showAll(ByVal bShow As Boolean)
    imgSun.Visible = bShow
    lblTitle.Visible = bShow
    lblTaskInstructions.Visible = bShow
End Sub
