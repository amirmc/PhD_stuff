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
   Begin VB.CheckBox chkShowMouse 
      BackColor       =   &H00000000&
      Caption         =   " Show Mouse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   8640
      TabIndex        =   7
      Top             =   4800
      Width           =   1935
   End
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
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   240
      TabIndex        =   6
      Top             =   165
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
      Enabled         =   0   'False
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
Private outputFile As String
Private strHeaderRowBuffer As String
Private strPracticeHeaderRowBuffer As String

Private Sub Form_Load()
    Me.BackColor = vbBlack      'in case default is not
End Sub

Private Sub CmdSave_Click()
    CommonDialog1.FileName = "" 'otherwise, it will still be what it was!
    CommonDialog1.Filter = "CSV Files (*.csv)|*.csv"
    CommonDialog1.ShowSave
    outputFile = Replace(CommonDialog1.FileName, ".csv", "Scans.csv")
        
''''''''''''''''''''''''
'' To format the Form ''
''''''''''''''''''''''''
    cmdInput.Enabled = True
    lblInput.Enabled = True
    lblSave.Caption = outputFile
End Sub

Private Sub cmdInput_Click()
    Dim Titles As String, HeaderItem As String
    Dim h_InputFile As Integer, arrayMax As Integer
    Dim i As Integer, pageNumber As Integer, item As Integer
    arrayMax = 1 'VB default is 0 but I need it to start counting from 1
    
    CommonDialog1.FileName = "" 'otherwise, it will still be what it was!
    CommonDialog1.Filter = "CSV Files (*.csv)|*.csv"
    CommonDialog1.ShowOpen
    Titles = CommonDialog1.FileName
    h_InputFile = FreeFile
    
    ''
    ''  The Following Loop reads in the data for the MAIN Experiment
    ''
    Open Titles For Input Access Read As #h_InputFile
    'Read in Header Row
    Input #h_InputFile, HeaderItem 'read in first item
    strHeaderRowBuffer = HeaderItem
    For item = 1 To iNumberOfInputColumns - 1 'this is set in PETScanningModule
        Input #h_InputFile, HeaderItem
        strHeaderRowBuffer = strHeaderRowBuffer & "," & HeaderItem
    Next
    'Read in Holidays, descriptions etc
    Do While Not EOF(h_InputFile)
        ReDim Preserve MenuSet(1 To arrayMax) 'resize array to input new item
        With MenuSet(arrayMax)
            For pageNumber = LBound(.MenuPage) To UBound(.MenuPage)
                For i = LBound(.MenuPage(pageNumber).Title) To UBound(.MenuPage(pageNumber).Title)
                    Input #h_InputFile, .MenuPage(pageNumber).Title(i), _
                                        .MenuPage(pageNumber).Desc(i), _
                                        .MenuPage(pageNumber).Condition(i), _
                                        .MenuPage(pageNumber).Incentive(i), _
                                        .MenuPage(pageNumber).Response(i), _
                                        .MenuPage(pageNumber).Trial(i), _
                                        .MenuPage(pageNumber).Page(i), _
                                        .MenuPage(pageNumber).ItemOrder(i)
                Next i
            Next pageNumber
        End With
        arrayMax = arrayMax + 1
    Loop
    lblArrayMax.Caption = arrayMax - 1
    Close #h_InputFile
    
    ''
    ''  The Following Loop reads in the data for the PRACTICE Experiment
    ''  It reuses the variables from the loop above
    ''
    'strPracticeMenus is defined in PETScanningModule
    Dim arrayCount As Integer, h_PracticeFile As Integer
    arrayCount = 1 'VB default is 0 but I need it to start counting from 1
    h_PracticeFile = FreeFile
    Open strPracticeMenus For Input Access Read As #h_PracticeFile
    'Have to read in Header Row to start with
    For item = 1 To iNumberOfInputColumns 'this is set in PETScanningModule
        Input #h_PracticeFile, HeaderItem
        strPracticeHeaderRowBuffer = HeaderItem
    Next
    'Read in Holidays, descriptions etc
    Do While Not EOF(h_PracticeFile)
        ReDim Preserve PracticeMenuSet(1 To arrayCount) 'resize array to input new item
        With PracticeMenuSet(arrayCount)
            For i = LBound(.Title) To UBound(.Title)
                Input #h_PracticeFile, .Title(i), _
                                       .Desc(i), _
                                       .Condition(i), _
                                       .Incentive(i), _
                                       .Response(i), _
                                       .Trial(i), _
                                       .Page(i), _
                                       .ItemOrder(i)
            Next i
        End With
        arrayCount = arrayCount + 1
    Loop
    Close #h_PracticeFile

''''''''''''''''''''''''
'' To format the Form ''
''''''''''''''''''''''''
    cmdStart.Enabled = True
    lblInput.Caption = Titles
End Sub

Private Sub cmdStart_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''
'' Open outputFile, assign it to #h_OutputFile ''
'' and then write two header lines             ''
'''''''''''''''''''''''''''''''''''''''''''''''''
    
    If chkShowMouse.Value = 0 Then bShowMouse = False
    If chkShowMouse.Value = 1 Then bShowMouse = True
    
    h_OutputFile = FreeFile
    Open outputFile For Output Access Write Lock Read Write As #h_OutputFile 'from JP's code
    
    'Write the Header row for the outputFile
    strHeaderRowBuffer = strHeaderRowBuffer & "," & "Chosen?" & "," _
                                                  & "Latency" & "," _
                                                  & "ND Errors?"
    
    Print #h_OutputFile, Format(Now, "hh:mm:ss  dd mmmm yyyy")
    Print #h_OutputFile, strHeaderRowBuffer
    
'''''''''''''''''''''''''''''''''''''''''''''''''''
'' Load all the PRACTICE forms and hide this one ''
'''''''''''''''''''''''''''''''''''''''''''''''''''
    Load FPracticeScreen
    Load FPracticeMainScreen
    Load FPracticeTouch
    FPracticeScreen.Show
    Me.Hide
End Sub
