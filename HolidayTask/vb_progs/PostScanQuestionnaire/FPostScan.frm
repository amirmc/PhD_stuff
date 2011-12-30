VERSION 5.00
Begin VB.Form FPostScanQuestionnaire 
   BorderStyle     =   0  'None
   Caption         =   "PostScan Questions"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10050
      TabIndex        =   16
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton opNonAffectiveAnswer 
      BackColor       =   &H00000000&
      Caption         =   "   Yes, I decided on a person"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   0
      Left            =   8490
      TabIndex        =   15
      Top             =   7485
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.OptionButton opNonAffectiveAnswer 
      BackColor       =   &H00000000&
      Caption         =   "   No, I could not think of anyone"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   1
      Left            =   8490
      TabIndex        =   14
      Top             =   8085
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7073
      TabIndex        =   7
      Top             =   5513
      Width           =   1215
   End
   Begin VB.Image imgQuit 
      Height          =   975
      Left            =   120
      Top             =   10440
      Width           =   1095
   End
   Begin VB.Shape shpPackageRateCursor 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      FillColor       =   &H0080FFFF&
      Height          =   345
      Left            =   10680
      Top             =   8535
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblPackageRating 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "lblPackageRating"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   9570
      TabIndex        =   26
      Top             =   8895
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblPackageDislike 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Would never do this"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   6360
      TabIndex        =   25
      Top             =   7920
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblPackageLike 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Would really like to do this"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   480
      Left            =   13680
      TabIndex        =   24
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblPackageIndifferent 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Would not mind doing this"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   9877
      TabIndex        =   23
      Top             =   7920
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label lblNoDecision 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Please rate the above Package Holiday"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   855
      Left            =   1080
      TabIndex        =   22
      Top             =   8160
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Shape shpRateCursor 
      BackColor       =   &H0080FFFF&
      BorderColor     =   &H0080FFFF&
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   2
      Left            =   10680
      Top             =   5738
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape shpRateCursor 
      BackColor       =   &H0080FFFF&
      BorderColor     =   &H0080FFFF&
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   1
      Left            =   10680
      Top             =   3848
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape shpRateCursor 
      BackColor       =   &H0080FFFF&
      BorderColor     =   &H0080FFFF&
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   0
      Left            =   10680
      Top             =   2048
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblRating 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "lblRating"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   9570
      TabIndex        =   21
      Top             =   2408
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblRating 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "lblRating"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   9570
      TabIndex        =   20
      Top             =   4208
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblRating 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "lblRating"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Index           =   2
      Left            =   9570
      TabIndex        =   19
      Top             =   6098
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblDifficultyRating 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "lblDifficultyRating"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   9570
      TabIndex        =   18
      Top             =   9420
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblIndifferent 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Would not mind doing this"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   9877
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Shape shpDifficultyCursor 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      FillColor       =   &H0080FFFF&
      Height          =   345
      Left            =   10680
      Top             =   9060
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Line linDivide 
      BorderColor     =   &H00404040&
      BorderWidth     =   8
      Visible         =   0   'False
      X1              =   525
      X2              =   14820
      Y1              =   6960
      Y2              =   6975
   End
   Begin VB.Label lblEasy 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7845
      TabIndex        =   13
      Top             =   9105
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblDifficult 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Hard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   12960
      TabIndex        =   12
      Top             =   9105
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblTaskInstruction 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Task Instruction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   1073
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   13215
   End
   Begin VB.Label lblNonAffectiveDecision 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Did you decide on someone for this question?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1215
      Left            =   1073
      TabIndex        =   10
      Top             =   7440
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblLike 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Would really like to do this"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   480
      Left            =   13680
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDislike 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Would never do this"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblDecisionDifficulty 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "How difficult was this decision for you?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   855
      Left            =   1073
      TabIndex        =   6
      Top             =   9000
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblHolidayDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 3 (Description)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   795
      Index           =   2
      Left            =   743
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   4995
   End
   Begin VB.Label lblHolidayDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 2 (Description)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   795
      Index           =   1
      Left            =   743
      TabIndex        =   4
      Top             =   3990
      Visible         =   0   'False
      Width           =   4995
   End
   Begin VB.Label lblHolidayDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 1 (Description)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   795
      Index           =   0
      Left            =   743
      TabIndex        =   3
      Top             =   2190
      Visible         =   0   'False
      Width           =   4995
   End
   Begin VB.Label lblHolidayTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 3 (Title)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   885
      Index           =   2
      Left            =   390
      TabIndex        =   2
      Top             =   5235
      Visible         =   0   'False
      Width           =   5700
   End
   Begin VB.Label lblHolidayTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 2 (Title)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   885
      Index           =   1
      Left            =   390
      TabIndex        =   1
      Top             =   3345
      Visible         =   0   'False
      Width           =   5700
   End
   Begin VB.Label lblHolidayTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Holiday 1 (Title)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   885
      Index           =   0
      Left            =   390
      TabIndex        =   0
      Top             =   1545
      Visible         =   0   'False
      Width           =   5700
   End
   Begin VB.Image imgRateScale 
      Height          =   360
      Index           =   0
      Left            =   6720
      Picture         =   "FPostScan.frx":0000
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   7995
   End
   Begin VB.Image imgRateScale 
      Height          =   360
      Index           =   1
      Left            =   6720
      Picture         =   "FPostScan.frx":7572
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   7995
   End
   Begin VB.Image imgRateScale 
      Height          =   360
      Index           =   2
      Left            =   6720
      Picture         =   "FPostScan.frx":EAE4
      Stretch         =   -1  'True
      Top             =   5730
      Visible         =   0   'False
      Width           =   7995
   End
   Begin VB.Image imgDifficultyScale 
      Height          =   360
      Left            =   8497
      Picture         =   "FPostScan.frx":16056
      Stretch         =   -1  'True
      Top             =   9052
      Visible         =   0   'False
      Width           =   4440
   End
   Begin VB.Image imgPackageRateScale 
      Height          =   360
      Left            =   6720
      Picture         =   "FPostScan.frx":1D5C8
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   7995
   End
End
Attribute VB_Name = "FPostScanQuestionnaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'internal stuff for counting and scoring
Private iItemSet As Integer
Private m_sRateBin As Single                'to denote which rateBin a give rating belongs in
Private sBinWidth As Single                 'defines the width of the rateBins
Private m_sRating As Single                 'rating (0-1) (used with all rate scales)
Private tempHolidayRating(0 To 2) As Single 'to hold values of holiday ratings
Private tempHolidayRateBin(0 To 2) As Single
Private tempDifficultyRating As Single      'to hold values of difficulty ratings
Private tempDifficultyRateBin As Single
Private tempPackageRating As Single         'to hold values of Package Holiday rating
Private tempPackageRateBin As Single
Private strNonAffectiveOutput As String     'to hold output for opButtons

'internal flags
Private bDecisionTrial As Boolean
Private bNonAffectiveTrial As Boolean
Private bNonDecisionTrial As Boolean
Private continueToNext As checkBeforeContinue
' This last array of Boolean values is to check if trial should continue
' The first three values in the above array will be for the Holiday Rating Bars
' and the last two are for the Decision trials (opButtons and Difficulty Bar)

'default numbers
Private Const sDefaultNumber = -47      'For whenever I need dummy data

'
' text strings
'
    'for 'Condition' Column in input file
Private Const strHighDecision = "HD"
Private Const strHighNoDecision = "HND"
Private Const strHighNonAffective = "HNA"
Private Const strLowDecision = "LD"
Private Const strLowNoDecision = "LND"
Private Const strLowNonAffective = "LNA"
    'caption underneath each rating bar
Private Const strPleaseRate = "Please Rate"
    '(mini) trial instructions (as a reminder of the task)
Private Const strDtrial = "Read and consider each item and then select your preferred holiday option"
Private Const strNDtrial = "Read and consider each item and then select the last option"
Private Const strNAtrial = "Read and consider each item and then select the option that is most similar to a holiday taken by someone you know (Otherwise select the last option)"
    'for output to file depending on subject's choice
Private Const strNonAffectiveDefault = "n/a"        'These three will be written
Private Const strNonAffectiveSomeone = "Someone"    'into strNonAffectiveOutput
Private Const strNonAffectiveNoone = "No-one"       'for output to File

Private Sub Form_Load()
    Me.BackColor = vbBlack  'in case default is not
    showAll (False)     'in case it's not already been done
    sBinWidth = 1 / iNumberOfRatingBins
End Sub

Private Sub cmdStart_Click()
    cmdStart.Visible = False    'hide the start button
    linDivide.Visible = True
    StartRating
End Sub

Private Sub StartRating()
'''''''''''''''''''''''''''''''''''''''
'' Used repeatedly to update the     ''
'' labels with the latest MenuScreen    ''
'''''''''''''''''''''''''''''''''''''''
    Call setDefaults    'centres all cursors and sets default values
    Call setupNextMenu    'iItemSet should get incremented here
    
'    lblTaskInstruction.Visible = True
'    linDivide.Visible = True    'show the dividing line across the page
    Call showAll(True)
End Sub

Private Sub setDefaults()
    'Making sure that all cursors are where they're supposed to be
    Dim k As Integer
    For k = shpRateCursor.lbound To shpRateCursor.ubound
        Call reCentreRateCursor(imgRateScale(k), shpRateCursor(k))
    Next
    
    'Put some default values in the temporary ratings variables
    Dim j As Integer
    For j = LBound(tempHolidayRating) To UBound(tempHolidayRating)
        tempHolidayRating(j) = sDefaultNumber   'put an obviously stupid value in there
        tempHolidayRateBin(j) = sDefaultNumber   'put an obviously stupid value in there
    Next
    tempDifficultyRating = sDefaultNumber  'put an obviously stupid value in there
    tempDifficultyRateBin = sDefaultNumber  'put an obviously stupid value in there
    tempPackageRating = sDefaultNumber      'put an obviously stupid value in there
    tempPackageRateBin = sDefaultNumber     'put an obviously stupid value in there
End Sub

Private Sub setupNextMenu()
    ' prepare the next set of Holidays before display
    ' showAll() should be FALSE when this sub is called
    Call showAll(False) 'just to make sure
    
    iItemSet = iItemSet + 1
       
    Call setBooleanDecisionValues   'the new iItemSet is used to set values
    Call setTaskInstructions    'the new boolean values are needed for this bit
            
    'to prepare the label captions for the next trial
    Dim i As Integer
    For i = lblHolidayTitle.lbound To lblHolidayTitle.ubound
        lblHolidayTitle(i).Caption = MenuScreen(iItemSet).Title(i)
        lblHolidayDesc(i).Caption = MenuScreen(iItemSet).Desc(i)
    Next
End Sub

Private Sub cmdNext_Click()
    'if there is something the subject has forgotten to rate
    'then do not save to file and do not move to next trial
    If (continueToNext.DifficultyCheck = False) Or _
        (continueToNext.nonAffectiveCheck = False) Or _
        (continueToNext.PackageHolidayCheck = False) _
        Then Exit Sub
    
    Dim i As Integer
    For i = LBound(continueToNext.HolidayCheck) To UBound(continueToNext.HolidayCheck)
        If continueToNext.HolidayCheck(i) = False Then Exit Sub
        'continueToNext.HolidayCheck(i) = False  ' see ***
    Next
    '*** can't have this line here because it means subject can
    'click the Next button a second time to bypass the check
      
    'To reset the values in continueToNext.
    Dim j As Integer
    For j = LBound(continueToNext.HolidayCheck) To UBound(continueToNext.HolidayCheck)
        continueToNext.HolidayCheck(j) = False  'to reset the values
    Next
    continueToNext.nonAffectiveCheck = False
    continueToNext.DifficultyCheck = False
    continueToNext.PackageHolidayCheck = False
    
    Call writeToFile    'now write the data from this set of Menus to the file
    
    Call showAll(False) 'this is purely aesthetic
    Call Wait(delayBetweenRatingPage)  'program adapted to make it work this way
    
    'Start the Form again unless we've finished all the items
    If (iItemSet = UBound(MenuScreen)) Then Final Else StartRating
End Sub

Private Sub setBooleanDecisionValues()
'This Sub looks at the info taken from the input file and then sets
'the boolean values used within this form
    
    Dim iFirstPage As Integer, iFirstLine As Integer, strConditionBuffer As String
    iFirstLine = LBound(MenuScreen(iItemSet).Title)
    
    'to make the following code in this procedure easier to read
    strConditionBuffer = MenuScreen(iItemSet).Condition(iFirstLine)
    
    If (strConditionBuffer = strHighDecision) Or (strConditionBuffer = strLowDecision) Then
        bDecisionTrial = True
        bNonAffectiveTrial = False
        bNonDecisionTrial = False
        ' No need to wait for input from NonAffective question or NoDecision rating
        continueToNext.nonAffectiveCheck = True
        continueToNext.DifficultyCheck = False
        continueToNext.PackageHolidayCheck = True
    ElseIf (strConditionBuffer = strHighNoDecision) Or (strConditionBuffer = strLowNoDecision) Then
        bDecisionTrial = False
        bNonAffectiveTrial = False
        bNonDecisionTrial = True
        ' No need to wait for input from NonAffective or Decision questions
        continueToNext.nonAffectiveCheck = True
        continueToNext.DifficultyCheck = True
        continueToNext.PackageHolidayCheck = False
    ElseIf (strConditionBuffer = strHighNonAffective) Or (strConditionBuffer = strLowNonAffective) Then
        bDecisionTrial = True
        bNonAffectiveTrial = True
        bNonDecisionTrial = False
        'Must wait for input from NonAffective and Decision questions NOT NoDecision rating
        continueToNext.nonAffectiveCheck = False
        continueToNext.DifficultyCheck = False
        continueToNext.PackageHolidayCheck = True
        'not getting holiday ratings here so don't need check
        Dim i As Integer
        For i = LBound(continueToNext.HolidayCheck) To UBound(continueToNext.HolidayCheck)
            continueToNext.HolidayCheck(i) = True
        Next i
    Else
        MsgBox ("Having trouble with InputFile. Check values in 'Condition' Column")
        End
    End If
End Sub

Private Sub setTaskInstructions()
'This bit uses the boolean values from setBooleanDecisionValues() to
'put the appropriate instructions for the trial onto the screen
    If bDecisionTrial Then
        If bNonAffectiveTrial Then
            lblTaskInstruction.Caption = strNAtrial
        Else
            lblTaskInstruction.Caption = strDtrial
        End If
    Else
        lblTaskInstruction.Caption = strNDtrial
    End If
End Sub

Private Sub imgRateScale_MouseDown(iHolidayRating As Integer, Button As Integer, Shift As Integer, Xcoord As Single, Ycoord As Single)
    m_sRating = Xcoord / imgRateScale(iHolidayRating).Width
    Call setRateBin(m_sRating)
    tempHolidayRating(iHolidayRating) = m_sRating
    tempHolidayRateBin(iHolidayRating) = m_sRateBin
    continueToNext.HolidayCheck(iHolidayRating) = True
    lblRating(iHolidayRating).Caption = ""
    
    shpRateCursor(iHolidayRating).Visible = False
    Call setRateCursor(m_sRating, imgRateScale(iHolidayRating), _
                        shpRateCursor(iHolidayRating))
End Sub

Private Sub imgDifficultyScale_MouseDown(Button As Integer, Shift As Integer, Xcoord As Single, Ycoord As Single)
    m_sRating = Xcoord / imgDifficultyScale.Width
    Call setRateBin(m_sRating)
    tempDifficultyRating = m_sRating
    tempDifficultyRateBin = m_sRateBin
    continueToNext.DifficultyCheck = True
    lblDifficultyRating.Caption = ""
    
    shpDifficultyCursor.Visible = False
    Call setRateCursor(m_sRating, imgDifficultyScale, shpDifficultyCursor)
End Sub

Private Sub imgPackageRateScale_MouseDown(Button As Integer, Shift As Integer, Xcoord As Single, Ycoord As Single)
    m_sRating = Xcoord / imgPackageRateScale.Width
    Call setRateBin(m_sRating)
    tempPackageRating = m_sRating
    tempPackageRateBin = m_sRateBin
    continueToNext.PackageHolidayCheck = True
    lblPackageRating.Caption = ""
    
    shpPackageRateCursor.Visible = False
    Call setRateCursor(m_sRating, imgPackageRateScale, shpPackageRateCursor)
End Sub

Private Sub opNonAffectiveAnswer_Click(Choice As Integer)
    continueToNext.nonAffectiveCheck = True
    
    ' must make sure the captions are the right way round on the form
    ' for this bit to work without any confusion
    If Choice = 0 Then strNonAffectiveOutput = strNonAffectiveSomeone
    If Choice = 1 Then strNonAffectiveOutput = strNonAffectiveNoone
    
    ' Or you could just save the caption of
    ' the respective radio-button instead
End Sub

Private Sub showAll(ByVal bShow As Boolean)
    lblTaskInstruction.Visible = bShow
    cmdNext.Visible = bShow
    Dim v As Variant
    For Each v In Me.lblHolidayTitle
        v.Visible = bShow
    Next
    For Each v In Me.lblHolidayDesc
        v.Visible = bShow
    Next
    'The following should only ever 'hide' the cursor,
    'never 'show' it when this procedure is called
    If Not bShow Then
        For Each v In Me.shpRateCursor
            v.Visible = bShow
        Next
        shpDifficultyCursor.Visible = bShow
        shpPackageRateCursor.Visible = bShow
    End If
    'Only show the rating scales if there is rating to be done
    If bNonAffectiveTrial Then
        lblDislike.Visible = False
        lblIndifferent.Visible = False
        lblLike.Visible = False
        For Each v In Me.lblRating
            v.Caption = strPleaseRate
            v.Visible = False
        Next
        For Each v In Me.imgRateScale
            v.Visible = False
        Next
    ElseIf Not bNonAffectiveTrial Then
        lblDislike.Visible = bShow
        lblIndifferent.Visible = bShow
        lblLike.Visible = bShow
        For Each v In Me.lblRating
            v.Caption = strPleaseRate
            v.Visible = bShow
        Next
        For Each v In Me.imgRateScale
            v.Visible = bShow
        Next
    End If
    
    'The following makes sure that the right things are shown/hidden
    If bDecisionTrial Then
        Call reCentreRateCursor(imgDifficultyScale, shpDifficultyCursor)
        If bNonAffectiveTrial Then
            Call DecisionSettings(bShow, bShow, False)
        Else
            Call DecisionSettings(bShow, False, False)
        End If
    ElseIf bNonDecisionTrial Then
        Call DecisionSettings(False, False, bShow)
    ElseIf Not bDecisionTrial And Not bNonAffectiveTrial And Not bNonDecisionTrial Then
        'ie if all boolean values are false
        Call DecisionSettings(False, False, False)
    Else
        MsgBox ("something wrong with Boolean bits in showAll()")
    End If
End Sub

Private Sub DecisionSettings(ByVal bShowDecision As Boolean, ByVal bShowNonAffective As Boolean, ByVal bShowNonDecision As Boolean)
    
        'for objects relating to non-decision trials
    lblPackageRating.Caption = strPleaseRate
    lblPackageRating.Visible = bShowNonDecision
    lblNoDecision.Visible = bShowNonDecision
    imgPackageRateScale.Visible = bShowNonDecision
    lblPackageDislike.Visible = bShowNonDecision
    lblPackageIndifferent.Visible = bShowNonDecision
    lblPackageLike.Visible = bShowNonDecision
    
        'for the objects relating to decisions
    lblDifficultyRating.Caption = strPleaseRate
    lblDifficultyRating.Visible = bShowDecision
    lblDecisionDifficulty.Visible = bShowDecision
    imgDifficultyScale.Visible = bShowDecision
    lblEasy.Visible = bShowDecision
    lblDifficult.Visible = bShowDecision
    
    'for the objects relating to the nonAffective decisions
    strNonAffectiveOutput = strNonAffectiveDefault
    lblNonAffectiveDecision.Visible = bShowNonAffective
    Dim v As Variant
    For Each v In opNonAffectiveAnswer
        v.Value = False
        v.Visible = bShowNonAffective
    Next

End Sub

Private Sub writeToFile()
    'Write the data to a csv string, then store string in file
    Dim i As Integer, strOutputBuffer As String
    With MenuScreen(iItemSet)
        For i = LBound(.Title) To UBound(.Title)
            strOutputBuffer = .Title(i) & "," _
                            & .Desc(i) & "," _
                            & .Condition(i) & "," _
                            & .Incentive(i) & "," _
                            & .Response(i) & "," _
                            & .Trial(i) & "," _
                            & .Page(i) & "," _
                            & .ItemOrder(i) & "," _
                            & tempHolidayRating(i) & "," _
                            & tempHolidayRateBin(i) & "," _
                            & tempDifficultyRating & "," _
                            & tempDifficultyRateBin & "," _
                            & tempPackageRating & "," _
                            & tempPackageRateBin & "," _
                            & strNonAffectiveOutput
            Print #h_OutputFile, strOutputBuffer
        Next
    End With
End Sub

Private Sub setRateBin(sRatingValue As Single)
''
'' There are two ways I could do this and I decided to use the For
'' loop instead of the Do While loop.
'' The variable m_sRateBin should be overwritten with a new value
'' before leaving this subroutine, either way
''
    Dim n As Integer
    For n = 1 To iNumberOfRatingBins
        If (sRatingValue < (n * sBinWidth)) Then
            m_sRateBin = n * sBinWidth
            Exit Sub
        End If
    Next n
    
    'if the above doesn't overwrite m_sRateBin, then it should be because
    'sRatingValue = 1
    If sRatingValue = 1 Then
        m_sRateBin = 111   'to make it really obvious in the output file
    Else
        MsgBox ("Something wrong with rating?")
    End If
 
    
''''''''''''''''''''''''''''''''''''''''''''''
'' This bit could replace the For-loop above
''''''''''''''''''''''''''''''''''''''''''''''
'    m_sRateBin = 0
'    Do While m_sRateBin = 0
'        n = n + 1
'        If (sRatingValue >= ((n - 1) * sBinWidth)) _
'                            And (sRatingValue < (n * sBinWidth)) _
'                            Then m_sRateBin = n * sBinWidth
'    Loop
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub setRateCursor(ByVal sRating As Single, objRateScale As Object, _
                            objRateCursor As Object)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' This is to set the location of the given rate cursor (objRateCursor)
'' to the sRating value on the given rating scale (objRateScale)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    objRateCursor.Visible = False   'just in case it's not already been done
    Call reCentreRateCursor(objRateScale, objRateCursor) 'reset the cursor
    objRateCursor.Left = objRateScale.Left + sRating _
                        * objRateScale.Width - objRateCursor.Width / 2
    objRateCursor.ZOrder '***
    objRateCursor.Visible = True
    
    ' *** VB does a BLOODY IRRITATING thing where it messes about with
    ' order of the 'layers' on the form. ZOrder is basically the
    ' "Bring to Front" command in case I don't notice at design time
End Sub

Private Sub reCentreRateCursor(ByVal objWhichScale As Object, _
                                objWhichCursor As Object)
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'' This moves the given cursor (objWhichCursor)
'' to the centre of the given scale (objWhichScale)
'''''''''''''''''''''''''''''''''''''''''''''''''''''
    objWhichCursor.Visible = False  'just in case it's not already been done
    objWhichCursor.Move objWhichScale.Left + (objWhichScale.Width _
                        - objWhichCursor.Width) / 2, objWhichScale.Top _
                        + (objWhichScale.Height - objWhichCursor.Height) / 2
End Sub

Private Sub Wait(delay_sec As Single)
    Dim sEndWait As Single
    sEndWait = Timer + delay_sec
    Do
        DoEvents
    Loop Until Timer > sEndWait
End Sub

Private Sub Final()
    Print #h_OutputFile, Format(Now, "hh:mm:ss  dd mmmm yyyy")
    Close #h_OutputFile
    End
End Sub

Private Sub imgQuit_Click()
    End
End Sub

'' I think this bit makes sure the form is always in Focus. Useful?
Private Sub Form_LostFocus()
    On Error Resume Next 'dunno what this bit means
    Me.SetFocus
End Sub
