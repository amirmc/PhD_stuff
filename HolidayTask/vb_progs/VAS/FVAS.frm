VERSION 5.00
Begin VB.Form FVAS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Analogue Scale"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTextBox 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8460
      TabIndex        =   5
      Text            =   "txtTextBox"
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdSetRating 
      Caption         =   "Set"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5633
      TabIndex        =   0
      Top             =   4013
      Width           =   735
   End
   Begin VB.TextBox txtRating 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   5033
      TabIndex        =   1
      Text            =   "txtRating"
      Top             =   3533
      Width           =   1935
   End
   Begin VB.Label lblAsciiCode 
      Caption         =   "lblAsciiCode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   6
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Shape shpRateCursor 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      FillColor       =   &H0080FFFF&
      Height          =   345
      Left            =   5903
      Top             =   1748
      Visible         =   0   'False
      Width           =   75
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
      Left            =   9323
      TabIndex        =   4
      Top             =   2168
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
      Left            =   1103
      TabIndex        =   3
      Top             =   2168
      Width           =   1440
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
      Left            =   5093
      TabIndex        =   2
      Top             =   2168
      Width           =   1680
   End
   Begin VB.Image imgRateScale 
      Enabled         =   0   'False
      Height          =   360
      Left            =   1703
      Picture         =   "FVAS.frx":0000
      Stretch         =   -1  'True
      Top             =   1733
      Width           =   8460
   End
End
Attribute VB_Name = "FVAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sRating As Single                     'rating (0-1)
Private m_bGotRating As Boolean
'Private Const PleaseRate = "Please Rate"

Private Sub Form_Load()
    Me.BackColor = vbBlack
    imgRateScale.Enabled = True
    StartRateTrial
End Sub

Private Sub StartRateTrial()
'    Call reCentreRateCursor
    shpRateCursor.Visible = True
'    m_bGotRating = False
End Sub

Private Sub imgRateScale_MouseDown(Button As Integer, Shift As Integer, Xcoord As Single, Ycoord As Single)
'    If m_bGotRating Then Exit Sub
'    m_bGotRating = True
    m_sRating = Xcoord / imgRateScale.Width
    Call setRateCursor(m_sRating)
    txtRating.text = m_sRating
'    Call Wait(0.75)
'    shpRateCursor.Visible = False
'    txtRating.text = PleaseRate
    StartRateTrial
End Sub

Private Sub txtRating_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSetRating_Click
End Sub

Private Sub cmdSetRating_Click()
    On Error Resume Next
    If (txtRating.text < 0) And (txtRating.text > 1) Then
        MsgBox ("Must be number between 0 and 1")
        Exit Sub
    End If
    setRateCursor txtRating.text
End Sub

Private Sub setRateCursor(ByVal sRating As Single)
    shpRateCursor.Visible = False
    Call reCentreRateCursor
    shpRateCursor.Left = imgRateScale.Left + sRating * imgRateScale.Width - shpRateCursor.Width / 2
    shpRateCursor.Visible = True
End Sub

Private Sub reCentreRateCursor()
    shpRateCursor.Visible = False
    shpRateCursor.Move imgRateScale.Left + (imgRateScale.Width - shpRateCursor.Width) / 2, imgRateScale.Top + (imgRateScale.Height - shpRateCursor.Height) / 2
End Sub

Private Sub Wait(delay_sec As Single)
    Dim sEndWait As Single
    sEndWait = Timer + delay_sec
    Do
        DoEvents
    Loop Until Timer > sEndWait
End Sub

Private Sub txtTextBox_KeyPress(KeyAscii As Integer)
    lblAsciiCode.Caption = KeyAscii & " " & Chr$(KeyAscii)
End Sub
'' I think this bit makes sure the form is always in Focus. Useful
Private Sub Form_LostFocus()
    On Error Resume Next 'dunno what this bit means
    Me.SetFocus
End Sub

