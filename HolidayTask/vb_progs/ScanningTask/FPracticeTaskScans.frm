VERSION 5.00
Begin VB.Form FPracticeTaskScans 
   BorderStyle     =   0  'None
   Caption         =   "PracticeScans"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MouseIcon       =   "FPracticeTaskScans.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label lblTaskInstruction 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "lblTaskInstruction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   1898
      TabIndex        =   8
      Top             =   360
      Width           =   11565
   End
   Begin VB.Label lblLink 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "and/or"
      BeginProperty Font 
         Name            =   "BernhardMod BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   0
      Left            =   7193
      TabIndex        =   7
      Top             =   4530
      Width           =   975
   End
   Begin VB.Label lblLink 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "and/or"
      BeginProperty Font 
         Name            =   "BernhardMod BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   1
      Left            =   7193
      TabIndex        =   6
      Top             =   7462
      Width           =   975
   End
   Begin VB.Label lblHolidayTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 1 (Title)"
      BeginProperty Font 
         Name            =   "BernhardMod BT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   630
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
         Name            =   "BernhardMod BT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   630
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
         Name            =   "BernhardMod BT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   630
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
         Name            =   "BernhardMod BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1050
      Index           =   0
      Left            =   1785
      TabIndex        =   2
      Top             =   3105
      Visible         =   0   'False
      Width           =   11805
   End
   Begin VB.Label lblHolidayDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 2 (Description)"
      BeginProperty Font 
         Name            =   "BernhardMod BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1050
      Index           =   1
      Left            =   1785
      TabIndex        =   1
      Top             =   6030
      Visible         =   0   'False
      Width           =   11805
   End
   Begin VB.Label lblHolidayDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 3 (Description)"
      BeginProperty Font 
         Name            =   "BernhardMod BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1050
      Index           =   2
      Left            =   1785
      TabIndex        =   0
      Top             =   8970
      Visible         =   0   'False
      Width           =   11805
   End
End
Attribute VB_Name = "FPracticeTaskScans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'internal flags
Private bGotChoice As Boolean
Private bGotSpacePress As Boolean

Private Sub Form_Load()
    Me.BackColor = vbBlack  'in case default is not
    If bShowMouse Then Me.MousePointer = 0
End Sub

Public Sub prepareForm()
'NB this Sub is Public so that other forms can use it
    Call startNextTrial
End Sub

Private Sub startNextTrial()
    bGotChoice = False
    bGotSpacePress = False
    Call showAll(False)
    Call resetColours
    Call setChoices
    Call showAll(True)
End Sub

Private Sub setChoices()
    Call showAll(False)     'just in case it's not been done already
    
    ' the value of iItemSet is on a different form so need FPracticeTaskInfoScreen.iPracticeItemSet
    With FPracticeTaskInfoScreen
        'Put the Holiday Titles and Descriptions in place
        Dim i As Integer
        For i = lblHolidayTitle.LBound To lblHolidayTitle.UBound
            lblHolidayTitle(i).Caption = PracticeMenuSet(.iPracticeItemSet).Title(i)
            lblHolidayDesc(i).Caption = PracticeMenuSet(.iPracticeItemSet).Desc(i)
        Next
    End With
End Sub

Private Sub lblHolidayTitle_Click(iHolidayChoice As Integer)
    gotChoice iHolidayChoice
End Sub
Private Sub lblHolidayDesc_Click(iHolidayChoice As Integer)
    gotChoice iHolidayChoice
End Sub

Private Sub gotChoice(iChosen As Integer)
    If bGotChoice Then Exit Sub
    If (FPracticeTaskInfoScreen.bDecisionTrial = False) And _
                (iChosen <> lblHolidayTitle.UBound) Then Exit Sub
    bGotChoice = True
    Call highlightChoice(lblHolidayTitle(iChosen), lblHolidayDesc(iChosen))  'includes a delay wait time
    Call showAll(False)
    Call Wait(betweenPageDelayTime)
    lblTaskInstruction.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If bGotSpacePress Then Exit Sub
    If KeyAscii = asciiSpaceBar Then
        bGotSpacePress = True
        gotoMainScreen
    End If
End Sub

Private Sub gotoMainScreen()
    FPracticeMainScreen.prepareForm
    If (FPracticeTaskInfoScreen.iPracticeItemSet = UBound(PracticeMenuSet)) Then _
                            FPracticeMainScreen.lblCompanyTagLine = strEndOfSessions
    FPracticeMainScreen.Show
    Me.Hide
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
    Dim v As Variant
    For Each v In lblHolidayTitle
        v.Visible = bShow
    Next
    For Each v In lblHolidayDesc
        v.Visible = bShow
    Next
    For Each v In lblLink
        v.Visible = bShow
    Next
    lblTaskInstruction.Visible = True 'always show this
End Sub

'' I think this bit makes sure the form is always in Focus. Useful?
Private Sub Form_LostFocus()
    On Error Resume Next 'dunno what this bit means
    Me.SetFocus
End Sub
