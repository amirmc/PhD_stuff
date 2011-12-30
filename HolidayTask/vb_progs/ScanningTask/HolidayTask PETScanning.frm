VERSION 5.00
Begin VB.Form FHolidayTaskScans1 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Holiday Task Scans"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1"
   WindowState     =   2  'Maximized
   Begin VB.Image imgQuit 
      Height          =   975
      Left            =   14280
      Top             =   10440
      Width           =   975
   End
   Begin VB.Label lblTaskInstruction 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Task Instructions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   2100
      TabIndex        =   6
      Top             =   360
      Width           =   11160
   End
   Begin VB.Label lblHolidayTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 1 (Title)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   0
      Left            =   780
      TabIndex        =   5
      Top             =   2475
      Visible         =   0   'False
      Width           =   13800
   End
   Begin VB.Label lblHolidayTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 2 (Title)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Index           =   1
      Left            =   780
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   13800
   End
   Begin VB.Label lblHolidayTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 3 (Title)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   2
      Left            =   780
      TabIndex        =   3
      Top             =   8340
      Visible         =   0   'False
      Width           =   13800
   End
   Begin VB.Label lblHolidayDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 1 (Description)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1050
      Index           =   0
      Left            =   4583
      TabIndex        =   2
      Top             =   3105
      Visible         =   0   'False
      Width           =   6195
   End
   Begin VB.Label lblHolidayDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 2 (Description)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1050
      Index           =   1
      Left            =   4583
      TabIndex        =   1
      Top             =   6030
      Visible         =   0   'False
      Width           =   6195
   End
   Begin VB.Label lblHolidayDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 3 (Description)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1050
      Index           =   2
      Left            =   4583
      TabIndex        =   0
      Top             =   8955
      Visible         =   0   'False
      Width           =   6195
   End
End
Attribute VB_Name = "FHolidayTaskScans1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'to measure times ... time functions should be (Timer - xTime)
Private dChoiceTime As Double

Private strErrorBuffer As String

'internal flags
Private bGotChoice As Boolean
Private bGotSpacePress As Boolean

'used when highlighting the subjects choice
Private Const colTitleHighlight = &HFFFF00      'light blue
Private Const colDescHighlight = &HFFFF00       'light blue
Private Const colTitleDefault = &H80FFFF        'yellowish (dark)
Private Const colDescDefault = &HC0FFFF         'yellowish (light)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' These subroutines MUST BE REMOVED before compiling for the final time! ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub imgQuit_Click()
    End
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Remove the above subroutines before compiling!
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_Load()
    Me.BackColor = vbBlack  'in case default is not
'    prepareForm
End Sub

Public Sub prepareForm()
'NB this Sub is Public so that other forms can use it
    bGotChoice = False
    bGotSpacePress = False
    Call resetColours
    startNextTrial
End Sub

Private Sub startNextTrial()
    strErrorBuffer = "n/a"
    Call showAll(True)
    dChoiceTime = Timer  'this will be used to calculate the time for a choice
End Sub

Private Sub lblHolidayTitle_Click(iHolidayChoice As Integer)
    gotChoice (iHolidayChoice)
End Sub
Private Sub lblHolidayDesc_Click(iHolidayChoice As Integer)
    gotChoice (iHolidayChoice)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Not bGotChoice Then Exit Sub
    If bGotSpacePress Then Exit Sub
    bGotChoice = False
    If KeyAscii = Asc(" ") Then
        bGotSpacePress = True
        FMainScreen.prepareForm
        If (FTaskInfoScreen.iItemSet = UBound(MenuSet)) Then
            Close #h_OutputFile
            FMainScreen.lblCompanyTagLine = "End of Sessions"
        End If
        FMainScreen.Show
        Me.Hide
    End If
End Sub

Private Sub gotChoice(iChoice As Integer)
    If (FTaskInfoScreen.bDecisionTrial = False) And _
                                    (iChoice <> (nDisplay - 1)) Then
        strErrorBuffer = strErrorBuffer & " " & iChoice & "-" & (Timer - dChoiceTime)
        Exit Sub
    End If
    If bGotChoice Then Exit Sub
    bGotChoice = True
    Call writeToFile(Me.Tag, iChoice, (Timer - dChoiceTime))
    Call highlightChoice(iChoice)       'includes a delay wait time
    Call showAll(False)
End Sub

Private Sub highlightChoice(iChoice As Integer)
    '
    '   need to make the selected choice become
    '   'highlighted' in some way and the rest
    '   of the possible options remain the same
    '
    '   Am still thinking of how best to do this
    '   (the following will do for now)
    '
    lblHolidayTitle(iChoice).ForeColor = colTitleHighlight
    lblHolidayDesc(iChoice).ForeColor = colDescHighlight
    Call Wait(choiceDelayTime)
    Call resetColours
End Sub
Private Sub resetColours()
    Dim v As Variant
    For Each v In lblHolidayTitle
        v.ForeColor = colTitleDefault
    Next
    For Each v In lblHolidayDesc
        v.ForeColor = colDescDefault
    Next
End Sub

Private Sub showAll(ByVal bShow As Boolean)
    lblTaskInstruction.Visible = bShow
    Dim v As Variant
    For Each v In lblHolidayTitle
        v.Visible = bShow
    Next
    For Each v In lblHolidayDesc
        v.Visible = bShow
    Next
End Sub

'' I think this bit makes sure the form is always in Focus. Useful?
Private Sub Form_LostFocus()
    On Error Resume Next 'dunno what this bit means
    Me.SetFocus
End Sub
