VERSION 5.00
Begin VB.Form FRateFaces 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7073
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   10725
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7073
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   10365
      Width           =   1215
   End
   Begin VB.Image imgFacePic 
      Height          =   3255
      Left            =   6000
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label lblInstruction 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Click 'Start' to Begin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   143
      TabIndex        =   6
      Top             =   4800
      Width           =   15075
   End
   Begin VB.Label lblLike 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Attractive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   11070
      TabIndex        =   5
      Top             =   9645
      Width           =   1455
   End
   Begin VB.Label lblDislike 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Unattractive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2850
      TabIndex        =   4
      Top             =   9645
      Width           =   1440
   End
   Begin VB.Shape shpRateCursor 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      FillColor       =   &H0080FFFF&
      Height          =   345
      Left            =   7650
      Top             =   9225
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblRating 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6540
      TabIndex        =   3
      Top             =   8940
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblIndifferent 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Neither"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   9645
      Width           =   1680
   End
   Begin VB.Image imgQuit 
      Height          =   975
      Left            =   14033
      Top             =   10530
      Width           =   1095
   End
   Begin VB.Image imgRateScale 
      Enabled         =   0   'False
      Height          =   360
      Left            =   3450
      Picture         =   "FRateFaces.frx":0000
      Stretch         =   -1  'True
      Top             =   9210
      Width           =   8460
   End
End
Attribute VB_Name = "FRateFaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'to measure times ... time functions should be (Timer - xTime)
Private dChoiceTime As Double
Private dStartTime As Double

'internal counting
Private iItemNumber As Integer
Private sRateBin As Single
Private sBinWidth As Single
Private bGotRating As Boolean           'internal flag
Private m_sRating As Single             'rating (0-1)

'text strings
Private Const PleaseRate = "Please Rate"

' API functions
' these are copy/pasted from the CD of a Visual Basic book
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, _
    lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, _
    ByVal Y As Long) As Long

Private Sub Form_Load()
    Me.KeyPreview = True
    Me.BackColor = vbButtonFace  ' (ie gray) in case default is not black
    sBinWidth = 1 / iNumberOfRatingBins
    Call reCentreRateCursor(imgRateScale, shpRateCursor)
    shpRateCursor.Visible = True
End Sub

Private Sub cmdStart_Click()
    cmdStart.Visible = False
    lblInstruction.Visible = False
    lblRating.Caption = PleaseRate
    lblRating.Visible = True
    imgRateScale.Enabled = True
    Call resetMousePos
    dStartTime = Timer 'so I can measure time from start of trial
    StartChoice
End Sub

Private Sub StartChoice()
''''''''''''''''''''''''''''''''''''''''''''''
'' Used repeatedly to update the Pic with
'' the latest File and reset the Timer
''''''''''''''''''''''''''''''''''''''''''''''
    'reset things for next choice trial
    imgFacePic.Visible = False
    shpRateCursor.Visible = False
    iItemNumber = iItemNumber + 1    'increment to next "FaceItem(iItemNumber)"
    imgFacePic.Picture = LoadPicture(App.Path & "\FacePics\" & FaceItems(iItemNumber) & ".bmp")
    Call centreFacePic
    Call Wait(0.5)
    lblRating.Caption = PleaseRate
    bGotRating = False
    imgFacePic.Visible = True
    dChoiceTime = Timer  'this will be used to calculate the time for a choice
End Sub

Private Sub imgRateScale_MouseDown(Button As Integer, Shift As Integer, Xcoord As Single, Ycoord As Single)
''
'' This subroutine takes the Rating from the
'' imgRateScale object and prints it to File
''
    'to make sure only the first click per trial is recorded
    If bGotRating Then Exit Sub
    bGotRating = True
    
    m_sRating = Xcoord / imgRateScale.Width     'normalising the rating (0-1)
    
    Call setRateBin(m_sRating)
    Call writeToFile(m_sRating)    'write to the output to file
    Call setRateCursor(m_sRating, imgRateScale, shpRateCursor)  'display the rating with the cursor
    lblRating.Caption = ""
'    lblRating.Caption = m_sRating  'this line useful for debugging
    Call Wait(waitDelay)
    Call resetMousePos
    
    'Check to see if we've done the last choice trial
    If (iItemNumber = UBound(FaceItems)) Then Final Else StartChoice
End Sub

Private Sub setRateBin(sRatingValue As Single)
''
'' There are two ways I could do this and I decided to use the For
'' loop instead of the Do While loop.
'' The variable sRateBin should be overwritten with a new value
'' before leaving this subroutine, either way
''
    Dim n As Integer
    For n = 1 To iNumberOfRatingBins
        If (sRatingValue < (n * sBinWidth)) Then
            sRateBin = n * sBinWidth
            Exit Sub
        End If
    Next n
    
    'if the above doesn't overwrite sRateBin, then it should be because
    'sRatingValue = 1
    If sRatingValue = 1 Then
        sRateBin = 111   'to make it really obvious in the output file
    Else
        MsgBox "Something wrong with rating?"
    End If
 
    
''''''''''''''''''''''''''''''''''''''''''''''
'' This bit could replace the For-loop above
''''''''''''''''''''''''''''''''''''''''''''''
'    sRateBin = 0
'    Do While sRateBin = 0
'        n = n + 1
'        If (sRatingValue >= ((n - 1) * sBinWidth)) _
'                            And (sRatingValue < (n * sBinWidth)) _
'                            Then sRateBin = n * sBinWidth
'    Loop
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub writeToFile(ByVal sRating As Single)
    'Write the data to a csv string, then store string in file
    Dim strOutputBuffer As String, i As Integer
    strOutputBuffer = FaceItems(iItemNumber) & "," _
                    & sRating & "," _
                    & sRateBin & "," _
                    & (Timer - dChoiceTime) & "," _
                    & (Timer - dStartTime) & "," _
                    & FStartScreen.outputFile
    Print #h_OutputFile, strOutputBuffer
End Sub

Private Sub setRateCursor(ByVal sRating As Single, objRateScale As Object, _
                            objRateCursor As Object)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' This is to set the location of the rate cursor (objRateCursor)
'' to the sRating value on the rating scale (objRateScale)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    objRateCursor.Visible = False   'just in case it's not already been done
    Call reCentreRateCursor(objRateScale, objRateCursor) 'reset the cursor
    objRateCursor.Left = objRateScale.Left + sRating _
                        * objRateScale.Width - objRateCursor.Width / 2
    objRateCursor.Visible = True
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

Public Sub centreFacePic()
''
'' This is to centre the Face pic in the Form before displaying it
''
    If (imgFacePic.Height > Me.ScaleHeight) Or (imgFacePic.Width > Me.ScaleWidth) Then
        MsgBox ("FacePic image too big to fit on screen - resize your images.  Program will terminate")
        End
    End If
    imgFacePic.Top = (((5 / 6) * Me.ScaleHeight) - imgFacePic.Height) / 2
    imgFacePic.Left = (Me.ScaleWidth - imgFacePic.Width) / 2
End Sub

Private Sub resetMousePos()
''
''  This code here is copy/pasted from a Visual Basic book
''
''  I just decided to centre the Mouse on one of the hidden
''  command buttons rahter than defining a location somewhere
''  or sending the mouse off-screen
''
    ' Get the coordinates (in pixels) of the center of the Command1 button.
    ' The coordinates are relative to the button's client area.
    Dim lpPoint As POINTAPI
    lpPoint.X = ScaleX(cmdClose.Width / 2, vbTwips, vbPixels)
    lpPoint.Y = ScaleY(cmdClose.Height / 2, vbTwips, vbPixels)
    ' Convert to screen coordinates.
    ClientToScreen cmdClose.hWnd, lpPoint
    ' Move the mouse cursor to that point.
    SetCursorPos lpPoint.X, lpPoint.Y
End Sub

Private Sub Wait(delay_sec As Single)
''
'' Just a loop that I call whenever I want a short delay
''
    Dim sEndWait As Single
    sEndWait = Timer + delay_sec
    Do
        DoEvents
    Loop Until Timer > sEndWait
End Sub

Private Sub Final()
''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Final output to Form at end of FaceItem array
''''''''''''''''''''''''''''''''''''''''''''''''''''
    Close #h_OutputFile
    shpRateCursor.Visible = False
    lblRating.Enabled = False
    imgFacePic.Visible = False
    cmdClose.Visible = True
    lblInstruction.Visible = True
    lblInstruction = "TheEnd"
End Sub

Private Sub cmdClose_Click()
    ' A VERY exciting and groundbreaking little subroutine!!!!
    End
    ' ... there.  Wasn't that just FABULOUS!!  I rock.
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If user hits the "Escape" key then give choice to exit
    ' 27="Escape" key
    If KeyAscii = 27 Then checkBeforeEnding Else Exit Sub
End Sub

'' I think this bit makes sure the form is always in Focus. Useful?
Private Sub Form_LostFocus()
    On Error Resume Next 'dunno what this bit means
    Me.SetFocus
End Sub


