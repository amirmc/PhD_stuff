Attribute VB_Name = "Module1"
'the delay time between each trial
Public Const waitDelay = 0.75 'seconds

'No. of Divisions wanted on the Rate Scale.
' default value is set in FStartScreen
Public iNumberOfRatingBins As Integer

Type POINTAPI
    X As Long
    Y As Long
End Type

'initialise the array for the Menu Items and open-ended questions
    Public FaceItems() As String
    
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

