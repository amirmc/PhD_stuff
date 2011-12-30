Attribute VB_Name = "PETScanningModule"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''  Bits used by different Forms                          ''
''  (to make it easier to change them from one location)  ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'this is the name of the file to be opened to save subjects data into
    Public h_OutputFile As Integer

'delay times used in the program. written here so they can be easily changed
Public Const betweenPageDelayTime = 0.5     'seconds

'this is the number of items that are
'displayed on screen at a time
Public Const nDisplay = 3

'this is the number of columns in the input *.csv file
Public Const iNumberOfInputColumns = 8

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
Public Const strNAinstr = "In this task, you will NOT be required to choose your preferred holiday NOR think about what the holidays would be like.                                                                                                         Instead, you are asked to read and consider each item and think about which one of the holiday options was the most popular in a random sample of 100 people.                                                                                                      When finished, please SELECT the option that you think was the most popular choice."
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

'initialise the array for the Menu Items
    Public EditedMenuSet() As detailArray


'''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Subroutines used by different parts of the code ''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub writeToFile()
    h_OutputFile = FreeFile
    Open FStartScreen.outputFile For Output Access Write Lock Read Write As #h_OutputFile 'from JP's code
    
    'Write the Header row for the outputFile
    Print #h_OutputFile, FStartScreen.strHeaderRowBuffer
    
    'Write the data to a csv string, then store string in file
    Dim strOutputBuffer As String
    Dim iPageItem As Integer, iPage As Integer
    For iPage = LBound(EditedMenuSet) To UBound(EditedMenuSet)
        With EditedMenuSet(iPage)
            For iPageItem = LBound(.Desc) To UBound(.Desc)
                strOutputBuffer = .Title(iPageItem) & "," _
                                & .Desc(iPageItem) & "," _
                                & .Condition(iPageItem) & "," _
                                & .Incentive(iPageItem) & "," _
                                & .Response(iPageItem) & "," _
                                & .Trial(iPageItem) & "," _
                                & .Page(iPageItem) & "," _
                                & .ItemOrder(iPageItem)
                Print #h_OutputFile, strOutputBuffer
                strOutputBuffer = ""
            Next iPageItem
        End With
    Next iPage
    Close #h_OutputFile
End Sub

Public Sub Wait(delay_sec As Single)
    Dim sEndWait As Single
    sEndWait = Timer + delay_sec
    Do
        DoEvents
    Loop Until Timer > sEndWait
End Sub
