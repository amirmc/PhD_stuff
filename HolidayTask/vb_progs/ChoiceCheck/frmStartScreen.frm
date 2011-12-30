VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStartScreen 
   Caption         =   "Start Screen"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInputFile 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmStartScreen"
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
    For iItem = 1 To iNumberOfInputColumns 'this is set in PostScanModule
        Input #FileNumber, HeaderItem
        strHeaderRowBuffer = strHeaderRowBuffer & HeaderItem & ","
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
    strHeaderRowBuffer = strHeaderRowBuffer & "Holiday Rating" & "," _
                                            & "Holiday RateBin" & "," _
                                            & "Difficulty Rating" & "," _
                                            & "Difficulty RateBin" & "," _
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


