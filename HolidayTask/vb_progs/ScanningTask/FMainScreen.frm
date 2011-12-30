VERSION 5.00
Begin VB.Form FMainScreen 
   BorderStyle     =   0  'None
   Caption         =   "The Holiday Store"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   FillColor       =   &H80000008&
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "FMainScreen.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSetStartMenu 
      Caption         =   "Set"
      Height          =   495
      Left            =   14760
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtStartMenu 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   15000
      MaxLength       =   2
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "xx"
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblStartingMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Next Menu ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   13920
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "xxxxx"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   14760
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Image imgCartoon 
      Height          =   3195
      Left            =   4680
      Picture         =   "FMainScreen.frx":0152
      Top             =   6240
      Width           =   3000
   End
   Begin VB.Image imgPalm 
      Height          =   1920
      Left            =   10320
      Picture         =   "FMainScreen.frx":1BE1
      Top             =   6360
      Width           =   1650
   End
   Begin VB.Image imgCompanyLogo 
      Height          =   2490
      Left            =   6120
      Picture         =   "FMainScreen.frx":26CC
      Top             =   2220
      Width           =   3150
   End
   Begin VB.Image imgIsland 
      Height          =   2925
      Left            =   120
      Picture         =   "FMainScreen.frx":3856
      Top             =   8520
      Width           =   3750
   End
   Begin VB.Image imgSurf 
      Height          =   2400
      Left            =   840
      Picture         =   "FMainScreen.frx":6A19
      Top             =   1800
      Width           =   3765
   End
   Begin VB.Image imgHike 
      Height          =   3000
      Left            =   720
      Picture         =   "FMainScreen.frx":7B9C
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
      Left            =   4433
      TabIndex        =   1
      Top             =   4740
      Width           =   6495
   End
   Begin VB.Image imgSun 
      Height          =   3510
      Left            =   11160
      Picture         =   "FMainScreen.frx":8D4F
      Top             =   1320
      Width           =   3750
   End
   Begin VB.Image imgPoolside 
      Height          =   2745
      Left            =   11640
      Picture         =   "FMainScreen.frx":AAFA
      Top             =   6240
      Width           =   3375
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
   Begin VB.Image imgScuba 
      Height          =   2595
      Left            =   9000
      Picture         =   "FMainScreen.frx":D2EF
      Top             =   8520
      Width           =   1290
   End
End
Attribute VB_Name = "FMainScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'internal flags
Private bGotSpacePress As Boolean
Private bGotShowMouseToggle As Boolean

Private Sub Form_Load()
    Me.BackColor = vbBlack  'in case it is not the default
    If bShowMouse Then Me.MousePointer = 0
    prepareForm
End Sub

Public Sub prepareForm()
'NB this Sub is Public so that other forms can use it
    bGotSpacePress = False
    bGotShowMouseToggle = False
    Me.MousePointer = 99 'set mouse pointer to custom
    lblCompanyTagLine.Caption = strCompanyTagLine
    lblProgress.Caption = FTaskInfoScreen.iItemSet & "/" & FStartScreen.lblArrayMax.Caption
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If bGotSpacePress Then Exit Sub
    If KeyAscii = asciiSpaceBar Then
        bGotSpacePress = True
        If (FTaskInfoScreen.iItemSet = UBound(MenuSet)) Then
            Print #h_OutputFile, Format(Now, "hh:mm:ss  dd mmmm yyyy")
            Close #h_OutputFile
            End
        Else
            lblCompanyTagLine.Caption = strGetReady
            Call Wait(mainScreenDelayTime)
            FTaskInfoScreen.prepareForm
            FTaskInfoScreen.Show
            Me.Hide
        End If
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
    
    'if the key below the Escape key is hit then show the mouse
    If KeyAscii = asciiMouseShowToggle Then
        bGotShowMouseToggle = Not bGotShowMouseToggle
        If bGotShowMouseToggle Then
            Me.MousePointer = 1 'set it to default pointer
        ElseIf Not bGotShowMouseToggle Then
            FMainScreen.MousePointer = 99
        Else
            MsgBox ("Having trouble with KeyPress() and asciiMouseShowToggle")
            End
        End If
    End If
    
End Sub

Private Sub lblProgress_Click()
'if this is clicked, it means the user wants to change the trial to start from
    
    txtStartMenu.Text = Val(Replace(lblProgress.Caption, (Right(lblProgress.Caption, 3)), "")) + 1
    
    ' toggle the bits that user can change things with
    lblStartingMenu.Visible = Not lblStartingMenu.Visible
    txtStartMenu.Visible = Not txtStartMenu.Visible
    cmdSetStartMenu.Visible = Not cmdSetStartMenu.Visible
End Sub

Private Sub cmdSetStartMenu_Click()
''
''  This bit of code is there so that the user can change
''  the next set of menus that are to be presented.
''  NB lblProgress shows how many trials have been completed. NOT the
''  next trial in the sequence.
''
    Dim iNewMenu As Integer
    iNewMenu = Val(txtStartMenu.Text)
    If (iNewMenu < LBound(MenuSet)) Then
        iNewMenu = LBound(MenuSet)
    ElseIf (iNewMenu > FStartScreen.lblArrayMax) Then
        iNewMenu = FStartScreen.lblArrayMax
    End If
    FTaskInfoScreen.iItemSet = iNewMenu - 1
    
    lblStartingMenu.Visible = False
    txtStartMenu.Visible = False
    cmdSetStartMenu.Visible = False
    Call prepareForm
    'to mimic the MouseShow toggle switch
    bGotShowMouseToggle = Not bGotShowMouseToggle
    Form_KeyPress (asciiMouseShowToggle)
End Sub

