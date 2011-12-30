Attribute VB_Name = "Module2"
Private k As Integer
Private i As Integer
Private j As Integer
Private Same As Boolean
Private BlockNumber As Integer
Private wait As Integer
Private Const minimumFeedbackDuration = 2000
Private Const TotalTimeInSeconds = 1080    '= 18mins
Private loopit As String
Private endyet As String

'text strings
Public Const strWaitForScanner = "Waiting for Scanner"
Public Const strPaused = "Paused"

Public timeStimShown As Double 'needed in frmStimuli
Private timeRespReceived As Double
Private timeFeedbackShown As Double
Private feedbackDuration As Double
Private timeFeedbackRemoved As Double
Private jitterTime As Double
Private blockStartTime As Double

Private whichOptions As String
Private feedbackOptions As String
Private strFeedbackShown As String

Private whichStimulus(40) As Variant
Private whichFeedback As String
'Const SamePictureCriterion = 6
Private SamePictureCriterion As Integer
'range of criteria
Private Const criterionLBound = 5
Private Const criterionUBound = 6
Private prevCriteria(1 To 3) As Integer

'input/output for GSR Trials
Public Const iClearSignal = 0
Public Const iOutputForPositive = 1
Public Const iOutputForNegative = 2
Public Const iThreeStimuliSignal = 4
Public Const iTwoStimuliSignal = 8

Public allowResponse As Boolean
Public bPreReversal As Boolean
Private bGotRevResp As Boolean
Private bIncorrectResponse As Boolean
Private bEndOfRun As Boolean
Public bKeyBlocked(37 To 39) As Boolean
Private iPortOut As Integer
Public iStimulusSignal As Integer
Public bGSR_Expt As Boolean
Public bFMRI_Expt As Boolean

Private stimSet(1 To 2) As Variant
Public currStimOrder(1 To 2) As Variant
Private prevRewardedStim As String
Private chosenRevStim As String
Private chosenStimulus As String

'used to decide what to do at reversal stage
Private bGotTotalCriteria As Boolean
Private bErrOnRev As Boolean
'Private bSameStimChoice As Boolean
Private prevSelectedStim As String

Public Sub prepCounters()
    k = 1
End Sub

Public Sub EXP_RUN()
    Randomize
    bPreReversal = True
    bIncorrectResponse = False
    bEndOfRun = False
    BlockNumber = frmStimuli.iStartBlock
    SamePictureCriterion = 6
    j = 1
    k = k + 3
    
    'create multidimensional arrays
    Dim strStim(1 To 3) As String
    Dim intValence(1 To 3) As Integer
    Dim currStim(1 To 3) As Variant
    Dim currValence(1 To 3) As Integer
    stimSet(1) = strStim()
    stimSet(2) = intValence()
    currStimOrder(1) = currStim()
    currStimOrder(2) = currValence()
    
    'NB stimuli are loaded but NOT displayed at this point
    Call frmStimuli.loadStimuli((BlockNumber * 6) - 5, 1)
    If bPreReversal Then Call storeRewardedStim((BlockNumber * 6) - 5, 1)

'    'useful for debugging
'    frmStimuli.lblDebug.Caption = frmStimuli.lblDebug.Caption & ", " & _
                        SamePictureCriterion
'    frmStimuli.lblDebug.Caption = frmStimuli.lblDebug.Caption & ", " & _
                        bPreReversal
    
'    For i = 1 To 3
'    frmStimuli.lblDebug.Caption = frmStimuli.lblDebug.Caption & ", " & _
'                        stimSet(1)(i)
'    frmStimuli.lblDebug.Caption = frmStimuli.lblDebug.Caption & ", " & _
'                        stimSet(2)(i)
'    Next i
'    frmStimuli.lblDebug.Caption = frmStimuli.lblDebug.Caption & " --"
    'useful for debugging
    frmStimuli.lblDebug.Caption = ""
    frmStimuli.lblDebug.Caption = SamePictureCriterion & ", " & BlockNumber
    For i = LBound(currStimOrder(1)) To UBound(currStimOrder(1))
    frmStimuli.lblDebug.Caption = frmStimuli.lblDebug.Caption & ", " & _
                        currStimOrder(1)(i)
    frmStimuli.lblDebug.Caption = frmStimuli.lblDebug.Caption & ", " & _
                        currStimOrder(2)(i)
    Next i
    
'    Beep
    'timeStimShown = GetTimer / 1000000
    blockStartTime = GetTimer / 1000
    allowResponse = True
    frmStimuli.SetFocus
End Sub

Public Sub buttonKeyResponse(keyCode As Integer)
    'if we've already had a response then do not do anything
    If allowResponse = False Then Exit Sub Else allowResponse = False
    If (keyCode <> 37) _
            And (keyCode <> 38) And (keyCode <> 40) _
                    And (keyCode <> 39) Then Exit Sub
    
    'record time of Response
    timeRespReceived = GetTimer
    
    'check if subject has hit a 'blocked' key. If so, note it
    If bKeyBlocked(37) = True And keyCode = 37 Then bIncorrectResponse = True
    If bKeyBlocked(38) = True And (keyCode = 38 Or keyCode = 40) Then bIncorrectResponse = True
    If bKeyBlocked(39) = True And keyCode = 39 Then bIncorrectResponse = True

'    If Not bGotRevResp And Not bPreReversal Then Call assignCorrectStim(keyCode)

    If Not bGotTotalCriteria And Not bPreReversal And Not bGotRevResp Then _
                                                    Call checkCurrResp(keyCode)
    If bGotTotalCriteria And Not bGotRevResp And Not bPreReversal Then _
                                                Call assignCorrectStim(keyCode)


    'send keyCode to appropriate place
    If keyCode = 37 Then
        EXP_RESPONSE ("left")
    ElseIf (keyCode = 38) Or (keyCode = 40) Then
        EXP_RESPONSE ("centre")
    ElseIf keyCode = 39 Then
        EXP_RESPONSE ("right")
    End If
End Sub

Private Sub checkCurrResp(keyCode As Integer)
    'if current response is same as prev rewarded
    'stim then allow reversal to go ahead next time
    '
    'if not then revert back to original block of
    'stimuli contincengcies and wait for criteria to be reached again
    If bGotTotalCriteria Then Exit Sub
    If bPreReversal Then Exit Sub
    If bGotRevResp Then Exit Sub
    'if subject hit prevRewardedStim and
    'if subject hits same as stimulus as before
    'ie if they hit the unrewarded stim to criterion
    Select Case keyCode
        Case 37
            If currStimOrder(1)(1) = prevRewardedStim Then _
                            bGotTotalCriteria = True Else bGotTotalCriteria = False
            If currStimOrder(1)(1) = prevSelectedStim Then _
                            bGotTotalCriteria = True Else bGotTotalCriteria = False ' bSameStimChoice = True Else bSameStimChoice = False
        Case 38, 40
            If currStimOrder(1)(2) = prevRewardedStim Then _
                            bGotTotalCriteria = True Else bGotTotalCriteria = False
            If currStimOrder(1)(2) = prevSelectedStim Then _
                            bGotTotalCriteria = True Else bGotTotalCriteria = False 'bSameStimChoice = True Else bSameStimChoice = False
        Case 39
            If currStimOrder(1)(3) = prevRewardedStim Then _
                            bGotTotalCriteria = True Else bGotTotalCriteria = False
            If currStimOrder(1)(3) = prevSelectedStim Then _
                            bGotTotalCriteria = True Else bGotTotalCriteria = False 'bSameStimChoice = True Else bSameStimChoice = False
    End Select
    
    If Not bGotTotalCriteria Then
        bErrOnRev = True
    '    k = k - 1   'ie do not print to file as a new block
    End If
End Sub

Private Sub assignCorrectStim(keyCode As Integer)
    If bPreReversal Then Exit Sub
    If bGotRevResp Then Exit Sub
    If Not bGotTotalCriteria Then Exit Sub
    If bIncorrectResponse Then Exit Sub
    bGotRevResp = True
    Call zeroAllFeedbacks 'just in case it's not been done
    Select Case keyCode
        Case 37
            If currStimOrder(1)(1) = prevRewardedStim Then
                bGotRevResp = False
            '    Exit Sub
            Else
                chosenRevStim = currStimOrder(1)(1)
                currStimOrder(2)(1) = 1
            End If
        Case 38, 40
            If currStimOrder(1)(2) = prevRewardedStim Then
                bGotRevResp = False
            '    Exit Sub
            Else
                chosenRevStim = currStimOrder(1)(2)
                currStimOrder(2)(2) = 1
            End If
        Case 39
            If currStimOrder(1)(3) = prevRewardedStim Then
                bGotRevResp = False
            '    Exit Sub
            Else
                chosenRevStim = currStimOrder(1)(3)
                currStimOrder(2)(3) = 1
            End If
    End Select
'    If bIncorrectResponse = True Then bGotRevResp = False
End Sub

Public Sub EXP_RESPONSE(whichKey As String)
    Dim randNum As Double

    k = k + 1
    
    allowResponse = False ' should already be set
'    Beep
'    frmStimuli.Label1.Caption = WhichKey
'    timeRespReceived = GetTimer

    For i = LBound(whichStimulus) To (UBound(whichStimulus) - 2) 'from 0 to 38
        whichStimulus(UBound(whichStimulus) - i) = whichStimulus(UBound(whichStimulus) - i - 1)
    Next i
    
    'record which stimulus was selected
    Call stimulusChoice(whichKey)
    
    'record which type of feedback to display
    Call feedbackSelection(whichKey)
    'this is the one to skip if the first trial after
    'a reversal is always supposed to be correct
    
    'load the feedback, depending on which type it is
    Call loadFeedback
        
    'the loadFeedback writes appropriate info to file before this
    'line exits the subroutine if an incorrect response was made
    If bIncorrectResponse = True Then
        Call OUTPUT_DATA(whichKey)
        bIncorrectResponse = False
        allowResponse = True
        frmStimuli.SetFocus
        Exit Sub
    End If
    
    Call frmStimuli.centreFeedback
    Call showStimuli(False)
    If bGSR_Expt Then Call pllOut(iClearSignal) 'to clear signals to printer port
    '
    ' you can insert time delay between stimulus
    ' and feedback in here if desired
    '
    frmStimuli.imgFeedback.Visible = True
    timeFeedbackShown = GetTimer / 1000000
    If bGSR_Expt Then Call pllOut(iPortOut) 'to output appropriate signal to the parallel printer port
    
    Randomize
    ' time delay to present feedback for
     randNum = (Int(Rnd * 500) + 1) 'an integer number of milliseconds
    For wait = 1 To 1
        For WaitFor = 1 To 1000
        DoEvents
        Next
        feedbackDuration = GetTimer / 1000000 - timeFeedbackShown
        If feedbackDuration * 1000000 < (minimumFeedbackDuration + randNum) Then
            wait = wait - 1
            'putcodehere
        End If
    Next
    
    chosenStimulus = whichStimulus(1)
    Call OUTPUT_DATA(whichKey)
        
    'test for critereon of 'n' consecutive stimulus selections
    Same = True
    For i = 1 To SamePictureCriterion
        If whichStimulus(1) <> whichStimulus(i) Then Same = False
    Next i
    
'    If Same = True Then Call prepNextBlock
    If Same = True Then
        Call setBoolRevSwitch
        prevSelectedStim = whichStimulus(1)
        Call prepNextBlock
    End If
    If bErrOnRev Then
        bErrOnRev = False
        BlockNumber = BlockNumber - 2
        bPreReversal = True
        bGotTotalCriteria = False
        Call prepNextBlock
    End If
    
    loopit = frmStimuli.WkbObj.Worksheets(1).Cells(BlockNumber * 6 - 5, j + 1).Value
    If Left(loopit, 3) <> "ret" Then
        j = j + 1
    Else
        j = 1
    End If
    
    frmStimuli.imgFeedback.Visible = False
    timeFeedbackRemoved = GetTimer / 1000000
    If bGSR_Expt Then Call pllOut(iClearSignal) 'to cancel signal to parallel port
    
    ' delay (jitter) between feedback
    ' and subsequent stimulus
    Randomize
    'jitter in range of 1 to 1.5 seconds
    randNum = (Int(Rnd * 500) + 1000) 'an integer number of milliseconds
    For wait = 1 To 1
        For WaitFor = 1 To 1000
            DoEvents
        Next
        jitterTime = GetTimer / 1000000 - timeFeedbackRemoved
        If jitterTime < randNum / 1000000 Then
            wait = wait - 1
            'putcodehere
        End If
    Next
    
    
    bIncorrectResponse = False
    refreshStimulusForm
End Sub

Private Sub refreshStimulusForm()
    Call frmStimuli.loadStimuli((BlockNumber * 6) - 5, j)

'    'useful for debugging
    frmStimuli.lblDebug.Caption = ""
    frmStimuli.lblDebug.Caption = SamePictureCriterion & ", " & BlockNumber
    For i = LBound(currStimOrder(1)) To UBound(currStimOrder(1))
    frmStimuli.lblDebug.Caption = frmStimuli.lblDebug.Caption & ", " & _
                        currStimOrder(1)(i)
    frmStimuli.lblDebug.Caption = frmStimuli.lblDebug.Caption & ", " & _
                        currStimOrder(2)(i)
    Next i
    
    bIncorrectResponse = False 'reset here
    allowResponse = True 'reset here
    
'    If GetTimer / 1000 > TotalTimeInSeconds Then
    If GetTimer / 1000 > (blockStartTime + TotalTimeInSeconds) Then
        Call EXP_END
    Else
        frmStimuli.imgFeedback.Visible = False
        If bGSR_Expt Then Call pllOut(iClearSignal) 'to cancel signal to parallel port
        Call showStimuli(True)
        timeStimShown = GetTimer / 1000
        If bGSR_Expt Then Call pllOut(iStimulusSignal) 'to denote that stimuli are displayed on screen
        If bFMRI_Expt Then
            dblCalcPulseTime_StimOn = objSS.GetLastPulseTime(False)
            intCalcPulseNum_StimOn = objSS.GetLastPulseNum(False)
            dblLastPulseTime_StimOn = objSS.GetLastPulseTime(True)
            intLastPulseNum_StimOn = objSS.GetLastPulseNum(True)
        End If
        
        ''''''''''''''''''''''''''''''''''''
        'added to make it work with scanner
        If bFMRI_Expt And Not bEndOfRun Then Call SS_waitForButtonBox ' waits for input from button box
        ''''''''''''''''''''''''''''''''''''
                
        frmStimuli.SetFocus
    End If
End Sub

Private Sub setBoolRevSwitch()
    'first there's a reversal, then a stim change, then another reversal
    'and so on. So this boolean value must toggle between the two every
    'other block.
    'when it's true, program uses input from the excel sheet.
    'when it's false, program must use the subject's responses.
    bPreReversal = Not bPreReversal
    If Not bPreReversal Then bGotRevResp = False
    bGotTotalCriteria = False
    bErrOnRev = False
'    If bPreReversal Then Call prepNextBlock
End Sub
Private Sub prepNextBlock()
    
'this bit of code moved to setBoolRevSwitch()
'    'first there's a reversal, then a stim change, then another reversal
'    'and so on. So this boolean value must toggle between the two every
'    'other block.
'    'when it's true, program uses input from the excel sheet.
'    'when it's false, program must use the subject's responses.
'    bPreReversal = Not bPreReversal
'    If Not bPreReversal Then bGotRevResp = False
    
    BlockNumber = BlockNumber + 1
    k = k + 1
    j = 1
    
    endyet = frmStimuli.WkbObj.Worksheets(1).Cells((BlockNumber * 6) - 5, 1).Value
    If Left(endyet, 3) = "END" Then BlockNumber = 1
    For i = LBound(whichStimulus) To UBound(whichStimulus)
        whichStimulus(i) = ""
    Next
    If bPreReversal Then Call storeRewardedStim((BlockNumber * 6) - 5, 1)
    
    Call setCriterion 'randomise criterion between specified ranges
    'update values in prevCriteria array
    For i = LBound(prevCriteria) To (UBound(prevCriteria) - 1)
        prevCriteria(UBound(prevCriteria) + 1 - i) = prevCriteria(UBound(prevCriteria) - i)
    Next i
    prevCriteria(LBound(prevCriteria)) = SamePictureCriterion
    If Not bPreReversal Then SamePictureCriterion = SamePictureCriterion + 1

'    'useful for debugging
'    frmStimuli.lblDebug.Caption = frmStimuli.lblDebug.Caption & ", " & _
'                        SamePictureCriterion
'    frmStimuli.lblDebug.Caption = frmStimuli.lblDebug.Caption & ", " & _
'                        bPreReversal
'    For i = 1 To 3
'    frmStimuli.lblDebug.Caption = frmStimuli.lblDebug.Caption & ", " & _
'                        stimSet(1)(i)
'    frmStimuli.lblDebug.Caption = frmStimuli.lblDebug.Caption & ", " & _
'                        stimSet(2)(i)
'    Next i
'    frmStimuli.lblDebug.Caption = frmStimuli.lblDebug.Caption & " --"
End Sub

Public Sub EXP_END()
    'end the current run and display the block number reached
    'can then enter this next time program is started to resume
    'where you left off
    Call showStimuli(False)
    bEndOfRun = True
    With frmStimuli
        .imgFeedback.Visible = False
        .cmdClose.Caption = "Close"
        .cmdClose.Enabled = True
        .txtStartBlock.Text = BlockNumber 'display how far it got
        .txtPosStim.Text = .intPositiveFeedbackCount
        .txtNegStim.Text = .intNegativeFeedbackCount
        .lblInfo.Caption = strPaused
        
        .lblInfo.Visible = True
        .txtStartBlock.Visible = True
        .txtPosStim.Visible = True
        .txtNegStim.Visible = True
        .cmdClose.Visible = True
        .cmdRun.Visible = True
        If .bRun1 Then
            .bRun1 = False
            .bRun2 = True
            .cmdRun.Caption = "Run 2"
        ElseIf .bRun2 Then
            .bRun2 = False
            .cmdRun.Visible = False
        End If
        .MousePointer = 0 'default mouse pointer
    End With
End Sub

Private Sub stimulusChoice(whichKey As String)
    whichOptions = currStimOrder(1)(1) & "_" & _
                   currStimOrder(1)(2) & "_" & _
                   currStimOrder(1)(3)
    Select Case whichKey
        Case "left"
            whichStimulus(1) = currStimOrder(1)(1)
        Case "centre"
            whichStimulus(1) = currStimOrder(1)(2)
        Case "right"
            whichStimulus(1) = currStimOrder(1)(3)
        Case Else
            MsgBox ("err in select case in 'stimulusChoice'")
    End Select
End Sub

Private Sub feedbackSelection(whichKey As String)
    feedbackOptions = currStimOrder(2)(1) & "_" & _
                      currStimOrder(2)(2) & "_" & _
                      currStimOrder(2)(3)
    Select Case whichKey
        Case "left"
            whichFeedback = currStimOrder(2)(1)
        Case "centre"
            whichFeedback = currStimOrder(2)(2)
        Case "right"
            whichFeedback = currStimOrder(2)(3)
        Case Else
            MsgBox ("err in select case in 'feedbackSelection'")
    End Select
End Sub

Private Sub loadFeedback()
    If bIncorrectResponse = True Then
        'selection of inactive stimulus, do nothing
        strFeedbackShown = "NONE"
        Exit Sub
    End If
    If (whichFeedback = 1) Then
        'correct choice, positive feedback
        frmStimuli.intPositiveFeedbackCount = frmStimuli.intPositiveFeedbackCount + 1
        If frmStimuli.wkbFeedbackFile.Worksheets(1).Cells(frmStimuli.intPositiveFeedbackCount, 1).Value = "return" Then frmStimuli.intPositiveFeedbackCount = 1
        frmStimuli.imgFeedback.Picture = LoadPicture(App.Path & "\feedback\" & frmStimuli.wkbFeedbackFile.Worksheets(1).Cells(frmStimuli.intPositiveFeedbackCount, 1).Value) ' & ".bmp")
        strFeedbackShown = frmStimuli.wkbFeedbackFile.Worksheets(1).Cells(frmStimuli.intPositiveFeedbackCount, 1).Value
        iPortOut = iOutputForPositive 'tag for positive affect
    ElseIf (whichFeedback = -1) Then
        'probablistic error so show negative feedback
        frmStimuli.intNegativeFeedbackCount = frmStimuli.intNegativeFeedbackCount + 1
        If frmStimuli.wkbFeedbackFile.Worksheets(1).Cells(frmStimuli.intNegativeFeedbackCount, 2).Value = "return" Then frmStimuli.intNegativeFeedbackCount = 1
        frmStimuli.imgFeedback.Picture = LoadPicture(App.Path & "\feedback\" & frmStimuli.wkbFeedbackFile.Worksheets(1).Cells(frmStimuli.intNegativeFeedbackCount, 2).Value) ' & ".bmp")
        strFeedbackShown = frmStimuli.wkbFeedbackFile.Worksheets(1).Cells(frmStimuli.intNegativeFeedbackCount, 2).Value & "_prbErr"
        iPortOut = iOutputForNegative 'tag for negative affect
    ElseIf (whichFeedback = 0) Then
        'unrewarded stimulus, show negative/neutral feedback
        frmStimuli.intNegativeFeedbackCount = frmStimuli.intNegativeFeedbackCount + 1
        If frmStimuli.wkbFeedbackFile.Worksheets(1).Cells(frmStimuli.intNegativeFeedbackCount, 2).Value = "return" Then frmStimuli.intNegativeFeedbackCount = 1
        frmStimuli.imgFeedback.Picture = LoadPicture(App.Path & "\feedback\" & frmStimuli.wkbFeedbackFile.Worksheets(1).Cells(frmStimuli.intNegativeFeedbackCount, 2).Value) ' & ".bmp")
        strFeedbackShown = frmStimuli.wkbFeedbackFile.Worksheets(1).Cells(frmStimuli.intNegativeFeedbackCount, 2).Value
        iPortOut = iOutputForNegative 'tag for negative affect
    Else
        MsgBox ("Err in EXP_Response, loadFeedback")
    End If
End Sub

Private Sub setCriterion()
    Randomize
    Dim bSameCriteria As Boolean
    Dim iCount As Integer
    
    SamePictureCriterion = Int(((criterionUBound - criterionLBound) + 1) * Rnd) + criterionLBound
'    SamePictureCriterion = Int((2 * Rnd) + 5) 'ie 5 or 6
    
    bSameCriteria = True
    For iCount = LBound(prevCriteria) To UBound(prevCriteria)
        If SamePictureCriterion <> prevCriteria(iCount) Then bSameCriteria = False
    Next iCount
    If bSameCriteria = True Then setCriterion
    'to make sure you don't get 'UBound(prevCriteria)+1' trials in a row
    'where the criterion is the same
End Sub

Private Sub storeRewardedStim(iFirstRow As Integer, iColumn As Integer)
    'store the strings and integers for this new block of stimuli
    With frmStimuli.WkbObj.Worksheets(1)
        stimSet(1)(1) = .Cells(iFirstRow, iColumn).Value
        stimSet(1)(2) = .Cells(iFirstRow + 1, iColumn).Value
        stimSet(1)(3) = .Cells(iFirstRow + 2, iColumn).Value
        stimSet(2)(1) = .Cells(iFirstRow + 3, iColumn).Value
        stimSet(2)(2) = .Cells(iFirstRow + 4, iColumn).Value
        stimSet(2)(3) = .Cells(iFirstRow + 5, iColumn).Value
    End With
    Call calcRewardedStim
End Sub
Private Sub calcRewardedStim()
    Dim sumFeedback As Integer
    sumFeedback = (stimSet(2)(1) + stimSet(2)(2) + stimSet(2)(3))
    If sumFeedback > 1 Or sumFeedback < -1 Then
        MsgBox ("check inputfile for duplicate feedback near line " & BlockNumber * 6)
    End If
    If stimSet(2)(1) = 1 Then prevRewardedStim = stimSet(1)(1)
    If stimSet(2)(2) = 1 Then prevRewardedStim = stimSet(1)(2)
    If stimSet(2)(3) = 1 Then prevRewardedStim = stimSet(1)(3)
End Sub

Public Sub resetLoadedContingencies()
    If bPreReversal Then Exit Sub
    If Not bGotRevResp Then
        'set all feedbacks to zero and must wait for input
        Call zeroAllFeedbacks
    ElseIf bGotRevResp Then
        'simply reset the values taking into account probablistic err
        Select Case chosenRevStim
            Case currStimOrder(1)(1)
                If currStimOrder(2)(1) = -1 Then
                    Call zeroAllFeedbacks
                    currStimOrder(2)(1) = -1
                Else
                    Call zeroAllFeedbacks
                    currStimOrder(2)(1) = 1
                End If
            Case currStimOrder(1)(2)
                If currStimOrder(2)(2) = -1 Then
                    Call zeroAllFeedbacks
                    currStimOrder(2)(2) = -1
                Else
                    Call zeroAllFeedbacks
                    currStimOrder(2)(2) = 1
                End If
            Case currStimOrder(1)(3)
                If currStimOrder(2)(3) = -1 Then
                    Call zeroAllFeedbacks
                    currStimOrder(2)(3) = -1
                Else
                    Call zeroAllFeedbacks
                    currStimOrder(2)(3) = 1
                End If
            Case Else
                MsgBox ("something wrong in resetLoadedContingencies")
        End Select
'        If currStimOrder(1)(1) = chosenRevStim Then
'            If currStimOrder(2)(1) = -1 Then
'                Call zeroAllFeedbacks
'                currStimOrder(2)(1) = -1
'            Else
'                Call zeroAllFeedbacks
'                currStimOrder(2)(1) = 1
'            End If
'        elseif
    End If
End Sub
Private Sub zeroAllFeedbacks()
    currStimOrder(2)(1) = 0
    currStimOrder(2)(2) = 0
    currStimOrder(2)(3) = 0
End Sub

Public Function OUTPUT_DATA(Optional whichKey As String)
    With frmStimuli.Active_Excel.ActiveWorkbook.Worksheets(1)
        .Cells(k, 1).Value = whichOptions
        .Cells(k, 2).Value = feedbackOptions
        .Cells(k, 3).Value = strFeedbackShown
        .Cells(k, 4).Value = chosenStimulus
        .Cells(k, 5).Value = whichFeedback
        .Cells(k, 6).Value = whichKey
        .Cells(k, 7).Value = timeStimShown
        .Cells(k, 8).Value = timeRespReceived / 1000
        .Cells(k, 9).Value = timeFeedbackShown * 1000
        .Cells(k, 10).Value = feedbackDuration * 1000
        .Cells(k, 11).Value = jitterTime * 1000
        .Cells(k, 12).Value = bIncorrectResponse
        .Cells(k, 13).Value = SamePictureCriterion
        If bFMRI_Expt Then      'extra outputs if in MR scanner
            .Cells(k, 14).Value = dblCalcPulseTime_StimOn
            .Cells(k, 15).Value = intCalcPulseNum_StimOn
            .Cells(k, 16).Value = dblLastPulseTime_StimOn
            .Cells(k, 17).Value = intLastPulseNum_StimOn
            .Cells(k, 18).Value = "" 'column break to make numbers easier to read
            .Cells(k, 19).Value = timeButtonBoxResponse - dblCalcPulseTime_Resp
            .Cells(k, 20).Value = timeButtonBoxResponse
            .Cells(k, 21).Value = dblCalcPulseTime_Resp
            .Cells(k, 22).Value = intCalcPulseNum_Resp
            .Cells(k, 23).Value = dblLastPulseTime_Resp
            .Cells(k, 24).Value = intLastPulseNum_Resp
        End If
    End With
    frmStimuli.Active_Workbook.Save
End Function

Public Sub showStimuli(bShow As Boolean)
    Dim v As Variant
    For Each v In frmStimuli.imgStimulus
        v.Visible = bShow
    Next
End Sub

Public Sub pllOut(iSignal As Variant)
    If bGSR_Expt = True Then Out &H378, iSignal
End Sub

Public Sub printRunStart()
    frmStimuli.Active_Excel.ActiveWorkbook.Worksheets(1).Cells(k + 2, 1).Value = "Run 2 Starts Here"
End Sub
