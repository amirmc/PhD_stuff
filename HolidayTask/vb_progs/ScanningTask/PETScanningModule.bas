Attribute VB_Name = "PETScanningModule"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''  Bits used by different Forms                          ''
''  (to make it easier to change them from one location)  ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'delay times used in the program. written here so they can be easily changed
Public Const mainScreenDelayTime = 10      'seconds
Public Const highlightTime = 0.5            'seconds
Public Const betweenPageDelayTime = 0.5     'seconds

Public bShowMouse As Boolean

''
''  Text Strings used in both the Practice Task and the Main Expt
''
    '' On FStartScreen
Public Const strPracticeMenus = "PracticeMenus.csv"
    ''  On the forms F*MainScreen
Public Const strCompanyTagLine = "Helping you find your ideal break"
Public Const strGetReady = "Get ready for next session"
Public Const strEndOfSessions = "End of Sessions"
    ''  On the forms F*TaskInfoScreen
    'for 'Condition' Column in input file
Public Const strHighDecision = "HD"
Public Const strHighNoDecision = "HND"
Public Const strHighNonAffective = "HNA"
Public Const strLowDecision = "LD"
Public Const strLowNoDecision = "LND"
Public Const strLowNonAffective = "LNA"
    'full task instructions
Public Const strDinstr = "In this task, you will be required to choose your preferred holiday.                                                                                                                    While reading each item, take time to think about what each holiday would be like.                                                                                  When finished, please SELECT your preferred holiday option."
Public Const strNDinstr = "In this task, you will NOT be required to choose your preferred holiday.                                                                                                                You will be presented with Package Holidays comprised of three activities.  While reading about each activity, take time to think about what the Package Holiday would be like.                                                                                        When finished, please TOUCH the LAST item."
Public Const strNAinstr = "In this task, you will NOT be required to choose your preferred holiday NOR think about what the holidays would be like.                                                                                                         Instead, you are asked to read and consider each item and think about which one of the holiday options was the most popular in a random survey of 100 people.                                                                                                When finished, please SELECT the option that you think was the most popular choice."
    '' On the forms F*Scans
    'for the link lines
Public Const strNDLink = "and"
Public Const strDLink = "or"
    'mini task instructions (ie just a reminder)
Public Const strDtrial = "Read and consider each item and then select your preferred holiday option"
Public Const strNDtrial = "Read and consider the Package Holiday and then touch the last item"
Public Const strNAtrial = "Read and consider each item and then select the option you think was the most popular"
    '' For all the highlighting
Public Const colTitleHighlight = &HFFFF00      'light blue
Public Const colDescHighlight = &HFFFF00       'light blue
Public Const colTitleDefault = &H80FFFF        'yellowish (dark)
Public Const colDescDefault = &HC0FFFF         'yellowish (light)
    '' For the KeyPress events in the different forms
Public Const asciiBackToFirstForm = 27 'Ascii code for the Escape Key
Public Const asciiSpaceBar = 32 'Ascii code for the Space Bar Key
Public Const asciiMouseShowToggle = 96 'Ascii code for the key right bellow Escape key (the "`" key)


'this is the number of items that are
'displayed on screen at a time
Public Const nDisplay = 3

'this is the number of columns in the input *.csv file
Public Const iNumberOfInputColumns = 8


'''''''''''''''''''''''''''
'' Other definitions etc ''
'''''''''''''''''''''''''''

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
Type pageDetailArray
    MenuPage(1 To 3) As detailArray
End Type

'initialise the array for the Menu Items
    Public MenuSet() As pageDetailArray
    Public PracticeMenuSet() As detailArray

'this is the name of the file to be opened to save subjects data into
    Public h_OutputFile As Integer


'''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Subroutines used by different parts of the code ''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub writeToFile(ByVal iPage As Integer, ByVal iHolidayChoice As Integer, _
                        dLatency As Double, strErrors As String)

    ' this bit is prob a bit dodgy but VB defaults boolean to FALSE
    ' so it should work... it doesn't look good though
    Dim Chosen(0 To nDisplay - 1) As Boolean
    Chosen(iHolidayChoice) = True   'this should be ok?

    'Write the data to a csv string, then store string in file
    Dim i As Integer, strOutputBuffer As String
    With MenuSet(FTaskInfoScreen.iItemSet)
        For i = LBound(.MenuPage(iPage).Title) To UBound(.MenuPage(iPage).Title)
            strOutputBuffer = .MenuPage(iPage).Title(i) & "," _
                            & .MenuPage(iPage).Desc(i) & "," _
                            & .MenuPage(iPage).Condition(i) & "," _
                            & .MenuPage(iPage).Incentive(i) & "," _
                            & .MenuPage(iPage).Response(i) & "," _
                            & .MenuPage(iPage).Trial(i) & "," _
                            & .MenuPage(iPage).Page(i) & "," _
                            & .MenuPage(iPage).ItemOrder(i) & "," _
                            & Chosen(i) & "," _
                            & dLatency & "," _
                            & strErrors
            Print #h_OutputFile, strOutputBuffer
        Next
    End With
End Sub

Public Sub highlightChoice(lblTitle As Label, lblDesc As Label)
    '
    '   need to make the selected choice become
    '   'highlighted' in some way and the rest
    '   of the possible options remain the same
    '
    lblTitle.ForeColor = colTitleHighlight
    lblDesc.ForeColor = colDescHighlight
    Call Wait(highlightTime)
    lblTitle.ForeColor = colTitleDefault
    lblDesc.ForeColor = colDescDefault
End Sub

Public Sub resetColours(lblTitle As Label, lblDesc As Label)
    Dim v As Variant
    For Each v In lblTitle
        v.ForeColor = colTitleDefault
    Next
    For Each v In lblDesc
        v.ForeColor = colDescDefault
    Next
End Sub


Public Sub Wait(delay_sec As Single)
    Dim sEndWait As Single
    sEndWait = Timer + delay_sec
    Do
        DoEvents
    Loop Until Timer > sEndWait
End Sub
