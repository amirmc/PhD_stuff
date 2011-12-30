VERSION 5.00
Begin VB.Form FHolidayTaskScans 
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
   MouseIcon       =   "FHolidayTaskScans.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1"
   WindowState     =   2  'Maximized
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
      TabIndex        =   8
      Top             =   7462
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
      Index           =   0
      Left            =   7200
      TabIndex        =   7
      Top             =   4530
      Width           =   975
   End
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
      TabIndex        =   6
      Top             =   360
      Width           =   11565
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
      Left            =   1778
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
      Left            =   1778
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
      Left            =   1778
      TabIndex        =   0
      Top             =   8970
      Visible         =   0   'False
      Width           =   11805
   End
End
Attribute VB_Name = "FHolidayTaskScans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'to measure times ... time functions should be (Timer - xTime)
Private dChoiceTime As Double

Public iChoicePage As Integer 'this is public because the writeToFile
'subroutine in FHolidayTaskScans needs it to print the choices to file

Private strErrorBuffer As String

'internal flags
Private bGotChoice As Boolean
Private bGotSpacePress As Boolean

Private Sub Form_Load()
    Me.BackColor = vbBlack  'in case default is not
    If bShowMouse Then Me.MousePointer = 0
'    prepareForm
End Sub

Public Sub prepareForm()
'NB this Sub is Public so that other forms can use it
    Call startNextTrial
End Sub

Private Sub startNextTrial()
    bGotChoice = False
    bGotSpacePress = False
    strErrorBuffer = "n/a"
    Call showAll(False)
    Call resetColours
    Call setChoices
    Call showAll(True)
    dChoiceTime = Timer  'this will be used to calculate the time for a choice
End Sub

Private Sub setChoices()
    Call showAll(False)     'just in case it's not been done already
    
    ' if last page of the previous iItemSet has been displayed,
    ' then 'reset' iChoicePage to zero.  this line of code should
    ' be redundant because iChoicePage is reset everytime iItemSet
    ' is incremented on FTaskInfoScreen
    If iChoicePage = UBound(MenuSet(FTaskInfoScreen.iItemSet).MenuPage) _
                        Then iChoicePage = 0
    
    iChoicePage = iChoicePage + 1 'increment through MenuSet(iItemSet).MenuPage(iChoicePage)
    
    ' the value of iItemSet is on a different form so need FTaskInfoScreen.iItemSet
    With FTaskInfoScreen
        'Put the Holiday Titles and Descriptions in place
        Dim i As Integer
        For i = lblHolidayTitle.LBound To lblHolidayTitle.UBound
            lblHolidayTitle(i).Caption = MenuSet(.iItemSet).MenuPage(iChoicePage).Title(i)
            lblHolidayDesc(i).Caption = MenuSet(.iItemSet).MenuPage(iChoicePage).Desc(i)
        Next
    End With
End Sub

Private Sub lblHolidayTitle_Click(iHolidayChoice As Integer)
    gotChoice iHolidayChoice, (Timer - dChoiceTime)
End Sub
Private Sub lblHolidayDesc_Click(iHolidayChoice As Integer)
    gotChoice iHolidayChoice, (Timer - dChoiceTime)
End Sub

Private Sub gotChoice(iChosen As Integer, dLatency As Double)
    If (FTaskInfoScreen.bDecisionTrial = False) And _
                                    (iChosen <> (nDisplay - 1)) Then
        strErrorBuffer = strErrorBuffer & "," & iChosen & " " & dLatency
        Exit Sub
    End If
    If bGotChoice Then Exit Sub
    bGotChoice = True
    Call writeToFile(iChoicePage, iChosen, dLatency, strErrorBuffer)
    Call highlightChoice(lblHolidayTitle(iChosen), lblHolidayDesc(iChosen))  'includes a delay wait time
    Call showAll(False)
    Call Wait(betweenPageDelayTime)
    
    'set up for next 'page' unless operator has finished the trial
    If bGotSpacePress Then Exit Sub
    If iChoicePage = UBound(MenuSet(FTaskInfoScreen.iItemSet).MenuPage) Then
        lblTaskInstruction.Visible = False
        Exit Sub
    Else
        startNextTrial
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If bGotSpacePress Then Exit Sub
    If KeyAscii = asciiSpaceBar Then
        bGotSpacePress = True
        gotoMainScreen
    End If
End Sub

Private Sub gotoMainScreen()
    FMainScreen.prepareForm
    If (FTaskInfoScreen.iItemSet = UBound(MenuSet)) Then _
                            FMainScreen.lblCompanyTagLine = strEndOfSessions
    FMainScreen.Show
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
