VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FStartScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Post Scan Debrief"
   ClientHeight    =   3810
   ClientLeft      =   3120
   ClientTop       =   3120
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRateDiv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   8640
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "xx"
      Top             =   600
      Width           =   220
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save As"
      Height          =   375
      Left            =   3930
      TabIndex        =   0
      Top             =   795
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3930
      TabIndex        =   2
      Top             =   3075
      Width           =   1215
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "Input File"
      Height          =   375
      Left            =   3930
      TabIndex        =   1
      Top             =   2115
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblRateBarDivisions 
      Caption         =   "Rate Bar Divisions"
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      Caption         =   "Create File to Save data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   315
      Width           =   8760
   End
   Begin VB.Label lblInput 
      Alignment       =   2  'Center
      Caption         =   "Choose Input Data"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1635
      Width           =   8760
   End
   Begin VB.Label lblArrayMax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "xxx"
      Enabled         =   0   'False
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   8625
      TabIndex        =   3
      Top             =   3555
      Width           =   375
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

Private Sub Form_Load()
    iNumberOfRatingBins = 10      'this is the default value
    txtRateDiv.Text = iNumberOfRatingBins
End Sub

Private Sub CmdSave_Click()
    CommonDialog1.FileName = "" 'otherwise, it will still be what it was!
    CommonDialog1.Filter = "CSV Files (*.csv)|*.csv"
    CommonDialog1.ShowSave
    outputFile = Replace(CommonDialog1.FileName, ".csv", "PostScan.csv")
        
''''''''''''''''''''''''
'' To format the Form ''
''''''''''''''''''''''''
    cmdInput.Enabled = True
    lblInput.Enabled = True
    lblSave.Caption = outputFile
End Sub

Private Sub cmdInput_Click()
    Dim Titles As String
    Dim FileNumber As Integer
    Dim arrayMax As Integer, i As Integer, iItem As Integer
    
    Dim HeaderItem As String
    arrayMax = 1 'VB default is 0 but I need it to start counting from 1
    
    CommonDialog1.FileName = "" 'otherwise, it will still be what it was!
    CommonDialog1.Filter = "CSV Files (*.csv)|*.csv"
    CommonDialog1.ShowOpen
    Titles = CommonDialog1.FileName
    FileNumber = FreeFile
    
    Open Titles For Input Access Read As #FileNumber
    'Read in Header Row
    Input #FileNumber, HeaderItem
    strHeaderRowBuffer = HeaderItem
    For iItem = 1 To iNumberOfInputColumns - 1 'this is set in PostScanModule
        Input #FileNumber, HeaderItem
        strHeaderRowBuffer = strHeaderRowBuffer & "," & HeaderItem
    Next
    'Read in Holidays, descriptions etc
    Do While Not EOF(FileNumber)
        ReDim Preserve MenuScreen(1 To arrayMax) 'resize array to input new item
        With MenuScreen(arrayMax)
            For i = LBound(.Title) To UBound(.Title)
                Input #FileNumber, .Title(i), _
                                   .Desc(i), _
                                   .Condition(i), _
                                   .Incentive(i), _
                                   .Response(i), _
                                   .Trial(i), _
                                   .Page(i), _
                                   .ItemOrder(i)
            Next i  'i should step from 0 to 2
        End With
        arrayMax = arrayMax + 1
    Loop
    
    'this should display the number of 'menu screens' loaded into the array
    lblArrayMax.Caption = arrayMax - 1
    Close #FileNumber

''''''''''''''''''''''''
'' To format the Form ''
''''''''''''''''''''''''
    cmdStart.Enabled = True
    lblInput.Caption = Titles
End Sub

Private Sub cmdStart_Click()
    'The maximum value for the rateBin Divisions is 10
    If (txtRateDiv.Text > 10) Or (txtRateDiv.Text < 1) Then
        txtRateDiv.Text = 10
        MsgBox "Reset to Default Value of 10"
    End If
    iNumberOfRatingBins = txtRateDiv.Text

'''''''''''''''''''''''''''''''''''''''''''''''''
'' Open outputFile, assign it to #h_OutputFile ''
'' and then write two header lines             ''
'''''''''''''''''''''''''''''''''''''''''''''''''
    h_OutputFile = FreeFile
    Open outputFile For Output Access Write Lock Read Write As #h_OutputFile 'from JP's code
    
    'Write the Header row for the outputFile
    strHeaderRowBuffer = strHeaderRowBuffer & "," & "Holiday Rating" & "," _
                                                  & "Holiday RateBin" & "," _
                                                  & "Difficulty Rating" & "," _
                                                  & "Difficulty RateBin" & "," _
                                                  & "PackageRating" & "," _
                                                  & "PackageRateBin" & "," _
                                                  & "Non Affective Answer"
    Print #h_OutputFile, Format(Now, "hh:mm:ss  dd mmmm yyyy")
    Print #h_OutputFile, strHeaderRowBuffer

'''''''''''''''''''''''''''''''''''
'' swap from one Form to another ''
'''''''''''''''''''''''''''''''''''
    Load FPostScanQuestionnaire
    FPostScanQuestionnaire.Show
    Me.Hide
End Sub
