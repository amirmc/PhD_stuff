VERSION 5.00
Begin VB.Form FDisplayHolidayMenus 
   BorderStyle     =   0  'None
   Caption         =   "PracticeScans"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSlash 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000C0C0&
      Height          =   195
      Left            =   14880
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "/"
      Top             =   120
      Width           =   135
   End
   Begin VB.TextBox txtPage 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   195
      Left            =   14640
      MaxLength       =   2
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "xx"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtPageTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   195
      Left            =   15000
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "xx"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtHolidayDesc 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Index           =   2
      Left            =   3923
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "FDisplayHolidayMenus.frx":0000
      Top             =   10080
      Width           =   7515
   End
   Begin VB.TextBox txtHolidayDesc 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Index           =   0
      Left            =   3923
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "FDisplayHolidayMenus.frx":0018
      Top             =   4200
      Width           =   7515
   End
   Begin VB.TextBox txtHolidayDesc 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Index           =   1
      Left            =   3923
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "FDisplayHolidayMenus.frx":0030
      Top             =   7140
      Width           =   7515
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   7455
      Width           =   975
   End
   Begin VB.Label lblDescWordCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "lblDescWordCount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   12000
      TabIndex        =   14
      Top             =   10320
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDescWordCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "lblDescWordCount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   12000
      TabIndex        =   13
      Top             =   7440
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDescWordCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "lblDescWordCount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   12000
      TabIndex        =   12
      Top             =   4440
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgQuit 
      Height          =   1335
      Left            =   240
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image imgNext 
      Height          =   1215
      Left            =   14040
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblHolidayDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
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
      TabIndex        =   5
      Top             =   8970
      Visible         =   0   'False
      Width           =   11805
   End
   Begin VB.Label lblHolidayDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
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
      TabIndex        =   4
      Top             =   6030
      Visible         =   0   'False
      Width           =   11805
   End
   Begin VB.Label lblHolidayDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
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
      TabIndex        =   3
      Top             =   3105
      Visible         =   0   'False
      Width           =   11805
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
      TabIndex        =   2
      Top             =   8340
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
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   13800
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
      TabIndex        =   0
      Top             =   2475
      Visible         =   0   'False
      Width           =   13800
   End
End
Attribute VB_Name = "FDisplayHolidayMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iItemSet As Integer 'to count through array

'internal flags
Private bDecisionTrial As Boolean
Private bNonAffectiveTrial As Boolean

Private Sub Form_Load()
    Me.BackColor = vbBlack  'in case default is not
'    iItemSet = 1 'so it starts counting from 1
    txtPage.Text = iItemSet 'at this point, both equal zero
    txtPageTotal.Text = FStartScreen.lblArrayMax.Caption
End Sub

Public Sub prepareForm()
'NB this Sub is Public so that other forms can use it
    Call startNextTrial
End Sub

Private Sub startNextTrial()
    Call showAll(False)
    
    'close the program if all menu items have been displayed
    If (iItemSet = UBound(EditedMenuSet)) Then imgQuit_Click
    
    'increment the Page iff the user has not entered a Page number
    If txtPage.Text = iItemSet Then
        iItemSet = iItemSet + 1 'should start counting from 1
    Else
        iItemSet = txtPage.Text
    End If
    txtPage.Text = iItemSet
    
    Call showAll(False)
    Call setChoices
    Call showAll(True)
End Sub

Private Sub setChoices()
    Call showAll(False)     'just in case it's not been done already
    Call setBooleanDecisionValues
    Call setTaskInstructions
    
    'Put the Holiday Titles and Descriptions in place
    Dim i As Integer
    For i = lblHolidayTitle.LBound To lblHolidayTitle.UBound
        lblHolidayTitle(i).Caption = EditedMenuSet(iItemSet).Title(i)
        lblHolidayDesc(i).Caption = EditedMenuSet(iItemSet).Desc(i)
        txtHolidayDesc(i).Text = EditedMenuSet(iItemSet).Desc(i)
        Call setWordCount(EditedMenuSet(iItemSet).Desc(i), i)
    Next
End Sub

Private Sub lblHolidayTitle_Click(iHolidayIndex As Integer)
    Dim strTitleChangeBuffer As String
    strTitleChangeBuffer = InputBox("Please make any changes then click OK", _
        "Holiday Title " & iHolidayIndex, lblHolidayTitle(iHolidayIndex).Caption)
    lblHolidayTitle(iHolidayIndex).Caption = strTitleChangeBuffer
End Sub

Private Sub txtHolidayDesc_KeyDown(iItem As Integer, KeyCode As Integer, iShift As Integer)
    'change the contents of the label if the F6 key is pressed
    If KeyCode = vbKeyF6 Then lblHolidayDesc(iItem).Caption = txtHolidayDesc(iItem).Text
    'Restore original label if the F9 key is pressed
    If KeyCode = vbKeyF9 Then
        lblHolidayDesc(iItem).Caption = EditedMenuSet(iItemSet).Desc(iItem)
        txtHolidayDesc(iItem).Text = EditedMenuSet(iItemSet).Desc(iItem)
    End If
    Call setWordCount(lblHolidayDesc(iItem).Caption, iItem)
End Sub

Private Sub imgNext_Click()
    Call saveEditing
    Call writeToFile
    Call showAll(False)
    lblTaskInstruction.Visible = False
    Call Wait(betweenPageDelayTime)
    startNextTrial
End Sub

Private Sub lblTaskInstruction_Click()
    'used to toggle the text boxes etc
    Dim v As Variant
    For Each v In txtHolidayDesc
        v.Visible = Not v.Visible
    Next
    For Each v In lblDescWordCount
        v.Visible = Not v.Visible
    Next
    Dim i As Integer
    For i = lblHolidayDesc.LBound To lblHolidayDesc.UBound
        If (txtHolidayDesc(i).Visible = True) Then
            lblHolidayDesc(i).BackColor = &H404040 'darkgrey
        ElseIf (txtHolidayDesc(i).Visible = False) Then
            lblHolidayDesc(i).BackColor = vbBlack
        End If
    Next i
    txtPage.Visible = Not txtPage.Visible
    txtSlash.Visible = Not txtSlash.Visible
    txtPageTotal.Visible = Not txtPageTotal.Visible
End Sub

Private Sub saveEditing()
    Dim i As Integer
    For i = lblHolidayDesc.LBound To lblHolidayDesc.UBound
        EditedMenuSet(iItemSet).Desc(i) = lblHolidayDesc(i).Caption
        EditedMenuSet(iItemSet).Title(i) = lblHolidayTitle(i).Caption
    Next i
End Sub

Private Sub setWordCount(strDescription As String, iHolidayDesc As Integer)
    Dim iChar As Integer, iSpaceCount As Integer, strChar As String
    For iChar = 1 To Len(strDescription)
        strChar = Mid(strDescription, iChar, 1)
        If strChar = " " Then iSpaceCount = iSpaceCount + 1
    Next iChar
    lblDescWordCount(iHolidayDesc).Caption = iSpaceCount + 1 & " / " & Len(strDescription)
End Sub

Private Sub setBooleanDecisionValues()
'This Sub looks at the info taken from the input file and then sets
'the boolean values used within this form
    
    Dim iFirstLine As Integer, strCondition As String
    iFirstLine = LBound(EditedMenuSet(iItemSet).Title)
    
    'to make the following code in this module easier to read
    strCondition = EditedMenuSet(iItemSet).Condition(iFirstLine)
    
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
    
    Dim v As Variant
    If bDecisionTrial Then
        For Each v In lblLink
            v.Caption = strDLink
        Next
        If bNonAffectiveTrial Then
            lblTaskInstruction.Caption = strNAtrial
        Else
            lblTaskInstruction.Caption = strDtrial
        End If
    Else
        For Each v In lblLink
            v.Caption = strNDLink
        Next
        lblTaskInstruction.Caption = strNDtrial
    End If
End Sub

Private Sub showAll(ByVal bShow As Boolean)
    Dim v As Variant
    For Each v In lblHolidayTitle
        v.Visible = bShow
    Next
    For Each v In lblHolidayDesc
        v.Visible = bShow
        v.BackColor = &H404040 'darkgrey
        'v.BackColor = vbBlack
    Next
    For Each v In lblLink
        v.Visible = bShow
    Next
    For Each v In txtHolidayDesc
        v.Visible = bShow
        'v.Visible = False 'bShow
    Next
    For Each v In lblDescWordCount
        v.Visible = bShow
        'v.Visible = False
    Next
    txtPage.Visible = bShow
    txtSlash.Visible = bShow
    txtPageTotal.Visible = bShow
    'txtPage.Visible = False
    'txtSlash.Visible = False
    'txtPageTotal.Visible = False
        
    lblTaskInstruction.Visible = True 'always show this
End Sub

Private Sub imgQuit_Click()
    Call writeToFile
    End
End Sub

'' I think this bit makes sure the form is always in Focus. Useful?
Private Sub Form_LostFocus()
    On Error Resume Next 'dunno what this bit means
    Me.SetFocus
End Sub

