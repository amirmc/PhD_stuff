VERSION 5.00
Begin VB.Form frmStimuli 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MouseIcon       =   "affective feedback.frx":0000
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   855
      Left            =   3120
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtNegStim 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "Neg"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtPosStim 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   360
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "Pos"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtStartBlock 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      MaxLength       =   2
      TabIndex        =   9
      Text            =   "BlockNumber"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.OptionButton opExptType 
      Caption         =   " fMRI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.OptionButton opExptType 
      Caption         =   " GSR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.OptionButton opExptType 
      Caption         =   " Neither"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblPulseCount 
      Alignment       =   2  'Center
      Caption         =   "(waiting for pulse)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   5520
      Width           =   15375
   End
   Begin VB.Label lblDebug 
      Caption         =   "Use this to output stuff to help debug"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6240
      TabIndex        =   7
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "lblInfo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   4800
      Width           =   15375
   End
   Begin VB.Image imgStimulus 
      Height          =   4935
      Index           =   3
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   4035
   End
   Begin VB.Image imgStimulus 
      Height          =   4935
      Index           =   2
      Left            =   5640
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   4035
   End
   Begin VB.Image imgStimulus 
      Height          =   4935
      Index           =   1
      Left            =   600
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   4035
   End
   Begin VB.Image imgFeedback 
      Height          =   6795
      Left            =   3480
      Top             =   480
      Width           =   6660
   End
End
Attribute VB_Name = "frmStimuli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'output ActiveX
Public Active_Excel As Excel.Application
Public Active_Workbook As Excel.Workbook
Public Active_Worksheet As Excel.Worksheet

'activeX data input
Public WkbObj As Workbook
Public wkbFeedbackFile As Workbook

'counters
Public intResponseCount As Integer
Public intPositiveFeedbackCount As Integer
Public intNegativeFeedbackCount As Integer
Public intFeedbackType As Integer
Public iStartBlock As Integer
'Private iPosStimPosn As Integer
'Private iNegStimPosn As Integer
Public bRun1 As Boolean
Public bRun2 As Boolean

Private Sub Form_Load()
    bFMRI_Expt = True
'    bFMRI_Expt = False
    
    bRun1 = False
    bRun2 = False
    
    If bGSR_Expt Then pllOut (iClearSignal) 'to clear any signals to the printer port
    imgFeedback.Visible = False
    lblInfo.Visible = False
    lblPulseCount.Visible = False
    lblDebug.Visible = False
    lblDebug.Caption = ""
    Me.MousePointer = 0
    opExptType(2).Value = True ' default to 'fMRI'
    txtStartBlock.Text = 1 'default value
    txtPosStim.Text = 1 'default
    txtNegStim.Text = 1 'default
    Call showStimuli(False)
    Call setActiveX_Names
    Call centreStimuli
    Call centreFeedback
    
    If bFMRI_Expt Then
        'Will only wait for first scanner pulse
        'for 2 mins, and then cut out
        objSS.SetTimeout (120000)
    
        PumpUpTheThreadPriority
        'The following line will only work when
        'there is a pio board installed on the PC
        If (objSS.Initialize("") <> 0) Then End
    
        'This is the alternative to pretend that the pio board really
        'is there and the scanner really is producing pulses
        'objSS.SetPretendMode True
        
        RestoreThreadPriority
    End If
End Sub


''
''  Image Handlers
''
Private Sub cmdStart_Click()
    
    'return current scannersync version
    If bFMRI_Expt Then MsgBox "Current ScannerSync version is " & objSS.GetVersion
    
    Call checkExptType
    Dim v As Variant
    For Each v In opExptType
        v.Visible = False
    Next
    Active_Workbook.SaveAs (App.Path & "\subjects\" & frmStimuli.Text1.Text)
    cmdStart.Visible = False
    cmdClose.Visible = True 'in case you need to quit at this stage
    Text1.Visible = False
    
    Call prepCounters
    
    StartTimer
    
    bRun1 = True
    cmdRun.Caption = "Run 1"
    cmdRun.Visible = True
End Sub
    
Private Sub cmdRun_Click()
    lblInfo.Visible = False
    cmdRun.Visible = False
    cmdClose.Visible = False

    If bRun2 Then Call printRunStart
    
    iStartBlock = txtStartBlock.Text
    txtStartBlock.Visible = False
    
    intPositiveFeedbackCount = txtPosStim.Text
    intNegativeFeedbackCount = txtNegStim.Text
    txtPosStim.Visible = False
    txtNegStim.Visible = False
    Me.MousePointer = 99 'custom pointer (ie transparent)

    ''''''''''''''''''''''''''''''''''''
    'added to make it work with scanner
    If bFMRI_Expt Then
        lblInfo.Caption = strWaitForScanner
        lblInfo.Visible = True
        lblPulseCount.Visible = True
        Call SS_waitForScanner  ' discards first 18 pulses
        lblInfo.Visible = False
        lblPulseCount.Visible = False
    End If
    ''''''''''''''''''''''''''''''''''''

    If bGSR_Expt Then Call pllOut(iStimulusSignal) 'to denote that stimuli are displayed
    Call EXP_RUN
    Call showStimuli(True)
    timeStimShown = GetTimer / 1000
    If bFMRI_Expt Then
        dblCalcPulseTime_StimOn = objSS.GetLastPulseTime(False)
        intCalcPulseNum_StimOn = objSS.GetLastPulseNum(False)
        dblLastPulseTime_StimOn = objSS.GetLastPulseTime(True)
        intLastPulseNum_StimOn = objSS.GetLastPulseNum(True)
    End If
    
    ''''''''''''''''''''''''''''''''''''
    'added to make it work with scanner
    If bFMRI_Expt Then Call SS_waitForButtonBox ' waits for input from button box
    ''''''''''''''''''''''''''''''''''''
    
    frmStimuli.SetFocus
End Sub

Private Sub cmdClose_Click()
    If bGSR_Expt Then Call pllOut(iClearSignal)
    If bFMRI_Expt Then objSS.Terminate
    Active_Workbook.Save
    End
End Sub

''
''  Code begins here
''
Private Sub Form_KeyUp(keyCode As Integer, Shift As Integer)
    If allowResponse = False Or bFMRI_Expt = True Then
        Exit Sub
    ElseIf allowResponse = True Then
        buttonKeyResponse (keyCode)
    End If
End Sub

Public Sub loadStimuli(iFirstRow As Integer, iColumn As Integer)
    With WkbObj.Worksheets(1)
        imgStimulus(1).Picture = LoadPicture(App.Path & "\stimuli\" & .Cells(iFirstRow, iColumn).Value & ".bmp")
        imgStimulus(2).Picture = LoadPicture(App.Path & "\stimuli\" & .Cells(iFirstRow + 1, iColumn).Value & ".bmp")
        imgStimulus(3).Picture = LoadPicture(App.Path & "\stimuli\" & .Cells(iFirstRow + 2, iColumn).Value & ".bmp")

        currStimOrder(1)(1) = .Cells(iFirstRow, iColumn).Value
        currStimOrder(1)(2) = .Cells(iFirstRow + 1, iColumn).Value
        currStimOrder(1)(3) = .Cells(iFirstRow + 2, iColumn).Value
        currStimOrder(2)(1) = .Cells(iFirstRow + 3, iColumn).Value
        currStimOrder(2)(2) = .Cells(iFirstRow + 4, iColumn).Value
        currStimOrder(2)(3) = .Cells(iFirstRow + 5, iColumn).Value
    End With
    Call checkActiveStimuli   '(iFirstRow, iColumn)
    Call resetLoadedContingencies
End Sub
Private Sub checkActiveStimuli() 'iFirstRow As Integer, iColumn As Integer)
    'make sure values are reset before checking
    Dim i As Integer
    For i = LBound(bKeyBlocked) To UBound(bKeyBlocked)
        bKeyBlocked(i) = False
    Next i
    
    'to ignore keypresses if on a two stimuli trial
    '
    If Right(currStimOrder(1)(1), 1) = "X" Then 'for left arrow key
        bKeyBlocked(37) = True
        iStimulusSignal = iTwoStimuliSignal
    ElseIf Right(currStimOrder(1)(2), 1) = "X" Then 'for centre arrow keys
        bKeyBlocked(38) = True
        iStimulusSignal = iTwoStimuliSignal
    ElseIf Right(currStimOrder(1)(3), 1) = "X" Then 'for right arrow key
        bKeyBlocked(39) = True
        iStimulusSignal = iTwoStimuliSignal
    Else
        For i = LBound(bKeyBlocked) To UBound(bKeyBlocked)
            bKeyBlocked(i) = False
        Next i
        iStimulusSignal = iThreeStimuliSignal
    End If
    
End Sub

Private Sub checkExptType()
    If opExptType(0).Value = True Then ' Neither
        bGSR_Expt = False
        bFMRI_Expt = False
    ElseIf opExptType(1).Value = True Then ' GSR
        bGSR_Expt = True
        bFMRI_Expt = False
    ElseIf opExptType(2).Value = True Then ' fMRI
        bGSR_Expt = False
        bFMRI_Expt = True
    Else
        MsgBox ("Something wrong in 'checkExptType()' Code")
    End If
    
    'just in case something silly happens
    If bGSR_Expt And bFMRI_Expt Then
        MsgBox ("Cannot be both fMRI & GSR Expt simultaneously")
        End
    End If
End Sub

''
''
''
Private Sub setActiveX_Names()
    Set Active_Excel = CreateObject("Excel.Application")
    Set Active_Workbook = Active_Excel.Workbooks.Add
    Set Active_Worksheet = Active_Workbook.Worksheets.Add
    Active_Excel.Visible = True
    Active_Excel.WindowState = 2
    Set WkbObj = GetObject(App.Path & "\bmpsequence.xls")
    Set Active_Worksheet = Active_Workbook.Worksheets.Add
    Set wkbFeedbackFile = GetObject(App.Path & "\feedbackFile.xls")
End Sub

Private Sub centreStimuli()
   'to arrange stimuli evenly along the width of the screen
    Dim sSpace As Single, v As Variant
    sSpace = (Me.ScaleWidth - (imgStimulus(1).Width + imgStimulus(2).Width + imgStimulus(3).Width)) / 4
    imgStimulus(1).Left = Me.ScaleLeft + sSpace
    imgStimulus(2).Left = (imgStimulus(1).Left + imgStimulus(1).Width) + sSpace
    imgStimulus(3).Left = (imgStimulus(2).Left + imgStimulus(2).Width) + sSpace
    'to arrange the stimuli along the centre of the screen
    For Each v In imgStimulus
        v.Top = (Me.ScaleHeight - v.Height) / 2
    Next
End Sub

Public Sub centreFeedback()
'    If (imgFeedback.Height > Me.ScaleHeight) Or (imgFeedback.Width > Me.ScaleWidth) Then
'        MsgBox ("Feedback image too big to fit on screen - resize your images.  Program will terminate")
'        End
'    End If
    imgFeedback.Top = (Me.ScaleHeight - imgFeedback.Height) / 2
    imgFeedback.Left = (Me.ScaleWidth - imgFeedback.Width) / 2
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If user hits the "Escape" key then give choice to exit
    ' 27="Escape" key
    If KeyAscii = 27 Then checkBeforeEnding Else Exit Sub
End Sub
Public Sub checkBeforeEnding()
    Dim vButtonChoice As VbMsgBoxResult
    vButtonChoice = MsgBox("Are your really SURE you meant to exit the program?", vbYesNo, "Escape Program")
    If vButtonChoice = vbYes Then
        If bGSR_Expt Then Call pllOut(iClearSignal)
        If bFMRI_Expt Then objSS.Terminate
        Active_Workbook.Save
        End
    ElseIf vButtonChoice = vbNo Then
        Exit Sub
    Else
        MsgBox ("Something screwy in checkBeforeEnding for exiting program")
    End If
End Sub

