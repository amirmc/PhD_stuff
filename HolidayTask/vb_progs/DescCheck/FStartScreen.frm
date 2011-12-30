VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FStartScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Holiday Preference Study"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "Input File"
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
      Left            =   5108
      TabIndex        =   2
      Top             =   3975
      Width           =   1335
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5048
      TabIndex        =   1
      Top             =   4815
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save As"
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
      Left            =   5108
      TabIndex        =   0
      Top             =   2535
      Width           =   1335
   End
   Begin VB.Label lblHolidayTaskTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "The Holiday Store"
      BeginProperty Font 
         Name            =   "BernhardMod BT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   405
      Width           =   11055
   End
   Begin VB.Label lblArrayMax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "xxx"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   11160
      TabIndex        =   5
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label lblInput 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Choose Input Data"
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
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3375
      Width           =   11040
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Create File to Save data"
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
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1935
      Width           =   11040
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FStartScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public outputFile As String
Public strHeaderRowBuffer As String

Private Sub Form_Load()
    Me.BackColor = vbBlack      'in case default is not
    cmdSave.Visible = False
End Sub

'Private Sub CmdSave_Click()
'    CommonDialog1.FileName = "" 'otherwise, it will still be what it was!
'    CommonDialog1.Filter = "CSV Files (*.csv)|*.csv"
'    CommonDialog1.ShowSave
'    outputFile = CommonDialog1.FileName ', ".csv", "Scans.csv")
'
'''''''''''''''''''''''''
''' To format the Form ''
'''''''''''''''''''''''''
'    cmdInput.Enabled = True
'    lblInput.Enabled = True
'    lblSave.Caption = outputFile
'End Sub

Private Sub cmdInput_Click()
    Dim strInputFile As String, HeaderItem As String
    Dim h_InputFile As Integer, arrayMax As Integer
    Dim i As Integer, pageNumber As Integer, item As Integer
    arrayMax = 1 'VB default is 0 but I need it to start counting from 1
    
    CommonDialog1.FileName = "" 'otherwise, it will still be what it was!
    CommonDialog1.Filter = "CSV Files (*.csv)|*.csv"
    CommonDialog1.ShowOpen
    strInputFile = CommonDialog1.FileName
    h_InputFile = FreeFile
    
    'instead of asking for the user for a file to save to
    outputFile = Replace(CommonDialog1.FileName, ".csv", "Edited.csv")
    lblSave.Caption = outputFile
    
    Open strInputFile For Input Access Read As #h_InputFile
    'Read in Header Row
    Input #h_InputFile, HeaderItem 'read in first item
    strHeaderRowBuffer = HeaderItem
    For item = 1 To iNumberOfInputColumns - 1 'this is set in PETScanningModule
        Input #h_InputFile, HeaderItem
        strHeaderRowBuffer = strHeaderRowBuffer & "," & HeaderItem
    Next
    'Read in Holidays, descriptions etc
    Do While Not EOF(h_InputFile)
        ReDim Preserve EditedMenuSet(1 To arrayMax) 'resize array to input new item
        With EditedMenuSet(arrayMax)
            For i = LBound(.Title) To UBound(.Title)
                Input #h_InputFile, .Title(i), _
                                    .Desc(i), _
                                    .Condition(i), _
                                    .Incentive(i), _
                                    .Response(i), _
                                    .Trial(i), _
                                    .Page(i), _
                                    .ItemOrder(i)
            Next i
        End With
        arrayMax = arrayMax + 1
    Loop
    Close #h_InputFile
    lblArrayMax.Caption = arrayMax - 1

''''''''''''''''''''''''
'' To format the Form ''
''''''''''''''''''''''''
    cmdStart.Enabled = True
    lblInput.Caption = strInputFile
End Sub

Private Sub cmdStart_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''
'' Load all the PRACTICE forms and hide this one ''
'''''''''''''''''''''''''''''''''''''''''''''''''''
    Load FDisplayHolidayMenus
    FDisplayHolidayMenus.prepareForm
    FDisplayHolidayMenus.Show
    Me.Hide
End Sub
