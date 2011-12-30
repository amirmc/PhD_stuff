Attribute VB_Name = "Task"
'the delay tinme between each trial
Public Const waitDelay = 0.75 'seconds

'No. of columns (minus first one) in the input file
Public Const nDesc = 39

'No. of Divisions wanted on the Rate Scale.
' default value is set in FStartScreen
Public iNumberOfRatingBins As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' This Type bit is my own.  Basically trying to pull all ''
'' info from only one text file and put into in one array ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Type detailArray
    MenuTitle As String
    MenuDesc(1 To nDesc) As String
End Type
Type POINTAPI
    X As Long
    Y As Long
End Type

'initialise the array for the Menu Items and open-ended questions
    Public MenuItems() As detailArray
    Public questionArray() As String
    
'this is the name of the file to be opened to save subjects data into
    Public h_OutputFile As Integer

Public Sub checkBeforeEnding()
    Dim vButtonChoice As VbMsgBoxResult
    vButtonChoice = MsgBox("Are your really SURE you meant to exit the program?", vbYesNo, "Escape Program")
    If vButtonChoice = vbYes Then
        End
    ElseIf vButtonChoice = vbNo Then
        Exit Sub
    Else
        MsgBox ("Something screwy in checkBeforeEnding for exiting program")
    End If
End Sub
