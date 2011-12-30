VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FDisplay 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Using the food allergy task"
   ClientHeight    =   9690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl MM 
      Height          =   330
      Index           =   0
      Left            =   1965
      TabIndex        =   5
      Top             =   2115
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MCI.MMControl MM 
      Height          =   330
      Index           =   1
      Left            =   2625
      TabIndex        =   4
      Top             =   1965
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Shape shpRateCursor 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   7260
      Width           =   90
   End
   Begin VB.Image imgRating 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   855
      Top             =   7260
      Width           =   8730
   End
   Begin VB.Image imgFeedback 
      Height          =   2940
      Left            =   2955
      Stretch         =   -1  'True
      Top             =   3675
      Width           =   5250
   End
   Begin VB.Image imgFace 
      Height          =   1605
      Left            =   4545
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1620
   End
   Begin VB.Label lblLowRate 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Prevents Allergy"
      Height          =   765
      Left            =   810
      TabIndex        =   2
      Top             =   7560
      Width           =   2955
   End
   Begin VB.Label lblHighRate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Causes Allergy"
      Height          =   1020
      Left            =   6900
      TabIndex        =   1
      Top             =   7530
      Width           =   2655
   End
   Begin VB.Label lblMid 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "No Effect"
      Height          =   840
      Left            =   3840
      TabIndex        =   0
      Top             =   7530
      Width           =   2910
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Press spacebar to continue"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   975
      TabIndex        =   3
      Top             =   765
      Width           =   3465
   End
   Begin VB.Shape shpBot 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1320
      Left            =   0
      Top             =   15
      Width           =   12435
   End
   Begin VB.Shape shpTop 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   810
      Left            =   120
      Top             =   8925
      Width           =   12435
   End
   Begin VB.Image imgLeft 
      Height          =   3795
      Left            =   675
      Stretch         =   -1  'True
      Top             =   2970
      Width           =   2295
   End
   Begin VB.Image imgRight 
      Height          =   3930
      Left            =   8025
      Stretch         =   -1  'True
      Top             =   2850
      Width           =   3540
   End
   Begin VB.Image imgCentre 
      Height          =   2505
      Left            =   4065
      Stretch         =   -1  'True
      Top             =   3090
      Width           =   2610
   End
End
Attribute VB_Name = "FDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lITIDuration_msec As Long         'blank between trials
Public bFixedStimulusDuration As Boolean 'Do not wait for response
Public lMaxStimDuration_msec As Long     'Time to wait for response
Public lMaxChoiceHold_msec As Long       'Max hold before stimulus ends
Public nSpacebarPerTrial

Public lRatingFeedbackDuration_msec As Long  'Period to show selected rating
Public lPreFeedbackDuration_msec As Long     'Period to show colour borders
Public lFeedbackDuration_msec As Long        'Period to show allergy outcome
Public lRatingDuration_msec As Long          'Period to allow rating cursor movement

Private m_nResponse As Integer               'Allergy predicted?
Private m_lResponseDuration_msec As Long     'Duration of keypress

Private m_bGotSpace As Boolean              'internal status flags
Private m_bDoingAllergy As Boolean          '
Private m_bDoingRating As Boolean           '

Private m_lResponseStartTime As Long            'Start of keypress
Private m_sRating As Single                     'rating (0-1)
Private m_bMouseRating As Boolean               'Use mouse click VAS
Private m_bNumberRating As Boolean              'Use number keys (not cursor)
Private m_nRatingDivisions As Integer           'resolution of scale for keys/cursor

Private Const KEY_ALLERGY = vbKeyZ
Private Const KEY_NOALLERGY = 191 '?/' key
Private Const ALLERGY_WAV = "allergy.wav"
Private Const NOALLERGY_WAV = "noallergy.wav"
Private Const ALLERGYSOUND = 0
Private Const NOALLERGYSOUND = 1
'read only
Public Property Get nResponse() As Integer
    nResponse = m_nResponse
End Property
Public Property Get lResponseDuration_msec() As Long
    lResponseDuration_msec = m_lResponseDuration_msec
End Property
'write only
Public Property Let bUseFaces(ByVal bUse As Boolean)
If bUse Then
    'size the image controls (stimuli)
    imgFace.Visible = True
    imgFace.Top = Screen.Height / 8
    imgFace.Height = Screen.Height / 4
    imgFace.Width = Screen.Height / 4
    imgFace.Left = (Screen.Width - imgFace.Width) / 2
    imgCentre.Top = imgFace.Top + imgFace.Height
    imgRight.Top = imgCentre.Top
    imgLeft.Top = imgCentre.Top
    imgLeft.Height = 3 * Screen.Height / 8
    imgRight.Height = imgLeft.Height
    imgCentre.Height = imgLeft.Height
    imgLeft.Left = 0
    imgRight.Left = Screen.Width / 2
    imgCentre.Left = Screen.Width / 4
    imgLeft.Width = Screen.Width / 2
    imgRight.Width = Screen.Width / 2
    imgCentre.Width = Screen.Width / 2
    imgFeedback.Top = Screen.Height / 8
    imgFeedback.Left = 0
    imgFeedback.Width = Screen.Width
    imgFeedback.Height = 3 * Screen.Height / 4
Else
    'size the image controls (stimuli)
    imgFace.Visible = False
    imgCentre.Top = Screen.Height / 4
    imgRight.Top = imgCentre.Top
    imgLeft.Top = imgCentre.Top
    imgLeft.Height = Screen.Height / 2
    imgRight.Height = imgLeft.Height
    imgCentre.Height = imgLeft.Height
    imgLeft.Left = 0
    imgRight.Left = Screen.Width / 2
    imgCentre.Left = Screen.Width / 4
    imgLeft.Width = Screen.Width / 2
    imgRight.Width = Screen.Width / 2
    imgCentre.Width = Screen.Width / 2
    imgFeedback.Top = Screen.Height / 8
    imgFeedback.Left = 0
    imgFeedback.Width = Screen.Width
    imgFeedback.Height = 3 * Screen.Height / 4
End If
End Property

Private Sub Form_Load()
    'Become full screen & position features
    Move 0, 0, Screen.Width, Screen.Height
    clearStimuli
    'size the borders (shapes)
    shpTop.Top = 0
    shpTop.Left = 0
    shpTop.Width = Screen.Width
    shpTop.Height = Screen.Height / 8
    shpBot.Left = 0
    shpBot.Height = shpTop.Height
    shpBot.Top = Screen.Height - shpTop.Height
    shpBot.Width = Screen.Width
    bUseFaces = False
    'Rating items
    imgRateScale.Move Screen.Width / 4, 3 * Screen.Height / 4 + imgRateScale.Height, Screen.Width / 2
    lblLowRate.Move imgRateScale.Left, imgRateScale.Top + imgRateScale.Height + 1
    lblMidRate.Move imgRateScale.Left + (imgRateScale.Width - lblMidRate.Width) / 2, imgRateScale.Top + imgRateScale.Height + 1
    lblHighRate.Move imgRateScale.Left + imgRateScale.Width - lblHighRate.Width, imgRateScale.Top + imgRateScale.Height + 1
    shpRateCursor.Move imgRateScale.Left + (imgRateScale.Width - shpRateCursor.Width) / 2, imgRateScale.Top
   'FeedbackMessage
    Call lblMessage.Move(0, Screen.Height / 2, Screen.Width)
    'Load sounds
    Dim i As Integer
    For i = 0 To 1
         ' Set properties needed by MCI to open.
        MM(i).Notify = False
        MM(i).Wait = True
        MM(i).Shareable = False
        MM(i).DeviceType = "WaveAudio"
    Next
    MM(ALLERGYSOUND).FileName = modMain.strStimulusDir & ALLERGY_WAV
    MM(ALLERGYSOUND).Command = "Open"
    MM(NOALLERGYSOUND).FileName = modMain.strStimulusDir & NOALLERGY_WAV
    MM(NOALLERGYSOUND).Command = "Open"
    'Default Timing values
    lRatingFeedbackDuration_msec = 1000
    lRatingDuration_msec = 4000
    bFixedStimulusDuration = False
    lMaxStimDuration_msec = 10000
    lITIDuration_msec = 50
    lPreFeedbackDuration_msec = 500
    lFeedbackDuration_msec = 2000
    lMaxChoiceHold_msec = 2000
    Hideall
    Call setRatingMode(True, False, 9) 'default to mouse-click rating
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If m_bDoingAllergy Then
        If m_lResponseStartTime = 0 Then
            m_lResponseStartTime = timeGetTime
            Select Case KeyCode
                Case KEY_ALLERGY: m_nResponse = RESPONSE_ALLERGY
                Case KEY_NOALLERGY:    m_nResponse = RESPONSE_NOALLERGY
                Case Else:   m_lResponseStartTime = 0
            End Select
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If m_bDoingAllergy Then
        If (m_lResponseStartTime > 0) And (m_lResponseDuration_msec = -1) Then
            If (KeyCode = KEY_ALLERGY) And (m_nResponse = RESPONSE_ALLERGY) Then
                m_lResponseDuration_msec = timeGetTime - m_lResponseStartTime
            ElseIf (KeyCode = KEY_NOALLERGY) And (m_nResponse = RESPONSE_NOALLERGY) Then
                m_lResponseDuration_msec = timeGetTime - m_lResponseStartTime
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'waiting for space?
    If KeyAscii = Asc(" ") Then m_bGotSpace = True
    'Rating with keyboard?
    If m_bDoingRating And Not m_bMouseRating Then
        Select Case KeyAscii
        Case Asc("z"), Asc("Z")
            If Not m_bNumberRating Then
                m_sRating = m_sRating - (1 / m_nRatingDivisions)
                If m_sRating <= 0 Then m_sRating = 0
                Call setRateCursor(m_sRating)
            End If
        Case Asc("/"), Asc("?")
            If Not m_bNumberRating Then
                m_sRating = m_sRating + (1 / m_nRatingDivisions)
                If m_sRating >= 1 Then m_sRating = 1
                Call setRateCursor(m_sRating)
            End If
        Case Asc("1") To Asc("9")
            If m_bNumberRating Then
                m_sRating = (KeyAscii - Asc("1")) / (m_nRatingDivisions - 1)
            End If
        End Select
    End If
End Sub

Private Sub imgRateScale_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_bMouseRating And m_bDoingRating Then
        m_sRating = X / imgRateScale.Width
    End If
End Sub

Public Sub doAllergyTrial(strStim1 As String, strStim2 As String, strFace As String, bAllergy As Boolean)
    m_lResponseDuration_msec = -1
    m_lResponseStartTime = 0
    m_nResponse = 0
    SetFocus
    Call Wait(lITIDuration_msec)
    Call showStimulus(strStim1, strStim2, strFace)
    m_bDoingAllergy = True
    If bFixedStimulusDuration Then
        'wait for the specified time
        Call Wait(lMaxStimDuration_msec)
    Else
        Dim lTimeOut As Long, lNow As Long
        lTimeOut = timeGetTime + lMaxStimDuration_msec
        Do
            DoEvents
            lNow = timeGetTime
            If m_lResponseStartTime Then
                If lNow > (m_lResponseStartTime + lMaxChoiceHold_msec) Then
                    m_lResponseDuration_msec = lMaxChoiceHold_msec
                End If
            End If
        Loop Until (lNow > lTimeOut) Or (lResponseDuration_msec > -1)
    End If
    'force key up events (in case still holding)
    Call Form_KeyUp(KEY_ALLERGY, 0)
    Call Form_KeyUp(KEY_NOALLERGY, 0)
    m_bDoingAllergy = False
    Call showOutcome(bAllergy)
End Sub

'Place borders / wait prefeedback / outcome / wait feedback / clear
Private Sub showOutcome(allergy As Boolean)
    If allergy Then
        With MM(ALLERGYSOUND)
            .To = 0
            .Command = "Seek"
            .Command = "Play"
        End With
        shpTop.BackColor = vbRed
        shpBot.BackColor = vbRed
        DoEvents
        Wait lPreFeedbackDuration_msec
        imgFeedback.Picture = LoadPicture("stimuli\+.jpg")
        imgFace.Picture = LoadPicture
    Else
        With MM(NOALLERGYSOUND)
            .To = 0
            .Command = "Seek"
            .Command = "Play"
        End With
        shpTop.BackColor = vbGreen
        shpBot.BackColor = vbGreen
        DoEvents
        Wait lPreFeedbackDuration_msec
        imgFeedback.Picture = LoadPicture("stimuli\-.jpg")
        imgFace.Picture = LoadPicture
    End If
    DoEvents
    Wait lFeedbackDuration_msec
    clearStimuli
End Sub
Private Sub Hideall()
    imgLeft.Picture = LoadPicture
    imgRight.Picture = LoadPicture
    imgCentre.Picture = LoadPicture
    imgFace.Picture = LoadPicture
    Call showRating(False)
End Sub

Private Sub showStimulus(Left As String, Optional Right As String = "", Optional face As String = "")
    If Right <> "" Then
        If randint(1, 2) = 2 Then
            imgLeft.Picture = LoadPicture(Left)
            imgRight.Picture = LoadPicture(Right)
        Else
            imgLeft.Picture = LoadPicture(Right)
            imgRight.Picture = LoadPicture(Left)
        End If
    Else
        imgCentre.Picture = LoadPicture(Left)
    End If
    imgFace.Picture = LoadPicture(face)
End Sub

''RATING TRIALS
Public Sub SetRatingText(Low As String, High As String, Optional Middle As String = "")
    'should be called after setRatingMode
    If m_bNumberRating Then
        lblLowRate.Caption = "1 - " & Low
        lblHighRate.Caption = Format(m_nRatingDivisions) & " - " & High
    Else
        lblLowRate.Caption = Low
        lblHighRate.Caption = High
    End If
    lblMidRate.Caption = Middle
End Sub

Public Sub setRatingMode(bmouse As Boolean, bNumberKeys As Boolean, nDivisions As Integer)
    m_bMouseRating = bmouse
    If Not bmouse Then
        m_bNumberRating = bNumberKeys
        m_nRatingDivisions = nDivisions
        If m_bNumberRating Then
            If m_nRatingDivisions < 3 Then m_nRatingDivisions = 3
            If m_nRatingDivisions > 9 Then m_nRatingDivisions = 9
        Else
            'left / right response for rating.
            'force an odd number of divisions (cursor starts in middle)
            If nDivisions Mod 2 = 0 Then m_nRatingDivisions = nDivisions + 1
        End If
    End If
End Sub

Private Sub showRating(ByVal bShow As Boolean)
    m_bDoingRating = bShow
    imgRateScale.Visible = bShow
    lblHighRate.Visible = bShow
    lblLowRate.Visible = bShow
    lblMidRate.Visible = bShow
    shpRateCursor.Visible = False 'hide
    If m_bMouseRating Then
        'Move mouse to centre bottom of screen
        showMouse (bShow)
        Call SetCursorPos(Screen.Width / (2 * Screen.TwipsPerPixelX), Screen.Height / (2 * Screen.TwipsPerPixelY))
    Else
        If Not m_bNumberRating Then
            'show movable cursor
            Call setRateCursor(0.5)
            shpRateCursor.Visible = bShow
        End If
    End If
End Sub

Private Sub setRateCursor(ByVal sRating As Single)
    shpRateCursor.Left = imgRateScale.Left + sRating * imgRateScale.Width - shpRateCursor.Width / 2
End Sub


Public Function getRating(Left As String, Optional Right As String = "", Optional face As String = "") As Single
    SetFocus
    Call showStimulus(Left, Right, face)
    Call showRating(True)
    If m_bNumberRating Or m_bMouseRating Then
        m_sRating = 0
        Do
            DoEvents
        Loop Until m_sRating <> 0
    Else
        m_sRating = 0.5
        Call Wait(lRatingDuration_msec)
    End If
    getRating = m_sRating
    If lRatingFeedbackDuration_msec Then
        Call setRateCursor(m_sRating)
        shpRateCursor.Visible = True
        Call Wait(lRatingFeedbackDuration_msec)
    End If
    Call showRating(False)
    Call Hideall
End Function

'clear all food & feedback stimuli
Private Sub clearStimuli()
    imgCentre.Picture = LoadPicture
    imgRight.Picture = LoadPicture
    imgLeft.Picture = LoadPicture
    imgFeedback.Picture = LoadPicture
    imgFace.Picture = LoadPicture
    shpTop.BackColor = Me.BackColor
    shpBot.BackColor = Me.BackColor
End Sub

Public Sub waitSpace(ByVal text As String)
    On Error Resume Next
    lblMessage.Caption = text
    lblMessage.Visible = True
    m_bGotSpace = False
    Do
        DoEvents
    Loop Until m_bGotSpace
    lblMessage.Visible = False
End Sub

Private Sub Form_LostFocus()
    On Error Resume Next
    Me.SetFocus
End Sub
