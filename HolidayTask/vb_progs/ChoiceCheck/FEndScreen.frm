VERSION 5.00
Begin VB.Form FEndScreen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EndScreen"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3165
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4185
      TabIndex        =   11
      Text            =   "xx"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2745
      TabIndex        =   10
      Text            =   "xx"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4185
      TabIndex        =   9
      Text            =   "xx"
      Top             =   2080
      Width           =   495
   End
   Begin VB.TextBox txtD3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2745
      TabIndex        =   8
      Text            =   "xx"
      Top             =   2080
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4185
      TabIndex        =   7
      Text            =   "xx"
      Top             =   1520
      Width           =   495
   End
   Begin VB.TextBox txtD2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2745
      TabIndex        =   6
      Text            =   "xx"
      Top             =   1520
      Width           =   495
   End
   Begin VB.TextBox txtNonAffective 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4185
      TabIndex        =   5
      Text            =   "xx"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtDecision 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2745
      TabIndex        =   4
      Text            =   "xx"
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   3105
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label lblNonAffective 
      Alignment       =   2  'Center
      Caption         =   "Non Affective"
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
      Left            =   3825
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblDecision 
      Alignment       =   2  'Center
      Caption         =   "Decision"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2385
      TabIndex        =   2
      Top             =   300
      Width           =   1215
   End
   Begin VB.Label lblFinished 
      Alignment       =   2  'Center
      Caption         =   "Finished"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3045
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "FEndScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sRateDifference As Single

Private Sub Form_Load()
    cmdClose.Enabled = False
    lblFinished.Visible = False
End Sub

Private Sub cmdStart_Click()
    'read in values from form and then do stuff
    
End Sub

Private Sub cmdClose_Click()
    End
End Sub

Private Sub startWorking()
    Call rearrangeData
End Sub

Private Sub rearrangeData()
    
End Sub
Private Sub ratingDifferences()
    m_sRateDifference = MenuScreen(iPage)
End Sub

