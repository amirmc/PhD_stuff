Attribute VB_Name = "PostScan"
Option Explicit

'delay times used in the program. written here so they can be easily changed
Public Const delayBetweenRatingPage = 0.75  'seconds

'this is the number of columns in the input *.csv file
Public Const iNumberOfInputColumns = 8

'this is the number of items that are
'displayed on screen at a time
Public Const nDisplay = 3

'No. of Divisions wanted on the Rate Scale.
' default value is set in FStartScreen
Public iNumberOfRatingBins As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' This Type bit is my own.  Basically trying to pull all ''
'' info from only one text file and put into in one array ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Type detailArray
    Title(0 To nDisplay - 1) As String
    Desc(0 To nDisplay - 1) As String
    Condition(0 To nDisplay - 1) As String
    Incentive(0 To nDisplay - 1) As String
    Response(0 To nDisplay - 1) As String
    Trial(0 To nDisplay - 1) As Integer
    Page(0 To nDisplay - 1) As Integer
    ItemOrder(0 To nDisplay - 1) As Integer
End Type

Type checkBeforeContinue
   HolidayCheck(0 To nDisplay - 1) As Boolean
   nonAffectiveCheck As Boolean
   DifficultyCheck As Boolean
   PackageHolidayCheck As Boolean
End Type


' initialise the array for the Menu Items and open-ended questions
' the data will start counting from 1 upto the max defined by the file
    Public MenuScreen() As detailArray
    
'this is the name of the file to be opened to save subjects data into
    Public h_OutputFile As Integer
