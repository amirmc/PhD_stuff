Attribute VB_Name = "ScanSyncCode"
Option Explicit

'object for all scannersync commands
Public objSS As New ScannerSync

''pulse measure output
''
' calculated values
Public intCalcPulseNum_StimOn As Integer
Public dblCalcPulseTime_StimOn As Double
Public intCalcPulseNum_Resp As Integer
Public dblCalcPulseTime_Resp As Double
' last measured values
Public intLastPulseNum_StimOn As Integer
Public dblLastPulseTime_StimOn As Double
Public intLastPulseNum_Resp As Integer
Public dblLastPulseTime_Resp As Double

'time taking
Public timeButtonBoxResponse As Double
'Public restimer As Double

'check pulse sync
Public dblChkForHowLong As Double
Public WaitForResponse As Integer

Private WholeExpLoop As Integer
Private response As Byte
Private r1 As String
Private r2 As String
Private r3 As String
Private r4 As String

Public Sub SS_waitForScanner()
    ' should not be in this sub if
    ' it's not an fMRI expt
    If bFMRI_Expt = False Then Exit Sub
    
    'wait till scanner has fired 18 pulses
    'because we discard these

''
'' Copied from Ola's ID/ED
''
    PumpUpTheThreadPriority
    objSS.StartExperiment (1000) 'ie TR is about 1second
    RestoreThreadPriority
    'Call StartTimer ' if in dummy mode
    'COLLECT DUMMY RUNS HERE
    
    Do
        DoEvents
        frmStimuli.lblPulseCount.Caption = 18 - objSS.GetLastPulseNum(True)
        objSS.SynchroniseExperiment True, 0 ' Actually wait for a pulse
    Loop Until objSS.GetLastPulseNum(True) > 18 'count 18 dummies (3 plus the start expt and checkpulsesynchrony pulses)
End Sub

Public Sub SS_waitForButtonBox()
    ' if it's NOT an fMRI expt then
    ' do not enter this subroutine
    If bFMRI_Expt = False Then Exit Sub
''
'' Copied from Ola's ID/ED
''
    For WholeExpLoop = 1 To 1
        WholeExpLoop = WholeExpLoop - 1
        For WaitForResponse = 1 To 1
            WaitForResponse = WaitForResponse - 1
            DoEvents
            response = objSS.GetResponse()
            'onlyonce = onlyonce + 1
            'If onlyonce = 1 Then
                'If (response And 2) = 0 Or (response And 4) = 0 Or (response And 8) = 0 Or (response And 16) = 0 Then
                    'r1 = "Button error"
                    'Exit Do
                'End If
            'End If
               
            If (response And 2) = 0 Or (response And 4) = 0 Or (response And 8) = 0 Then 'Or (response And 16) = 0 Then
                r1 = (response And 2)
                r2 = (response And 4)
                r3 = (response And 8)
                'r4 = (response And 16)
                     
                WaitForResponse = 4
            End If
        
            'WaitForTime GetTimer + 40
            dblChkForHowLong = 40
            objSS.CheckPulseSynchronyForTime (dblChkForHowLong)
        Next
        dblCalcPulseTime_Resp = objSS.GetLastPulseTime(False)
        intCalcPulseNum_Resp = objSS.GetLastPulseNum(False)
        timeButtonBoxResponse = GetTimer
        dblLastPulseTime_Resp = objSS.GetLastPulseTime(True)
        intLastPulseNum_Resp = objSS.GetLastPulseNum(True)
        If allowResponse = True Then buttonKeyResponse (SS_calculateKeyCode)
        DoEvents
        WaitForResponse = 1
    Next
End Sub

Private Function SS_calculateKeyCode() As Integer
    ' should not be calculating
    ' this if it's NOT an fMRI expt
    If bFMRI_Expt = False Then Exit Function
''
'' This code taken and adapted from Ola's ID/ED task
''
    If r1 = 0 Then
        SS_calculateKeyCode = 37
    ElseIf r2 = 0 Then
        SS_calculateKeyCode = 38
    ElseIf r3 = 0 Then
        SS_calculateKeyCode = 39
    Else
        'SS_calculateKeyCode = 0 'a default number?
    End If
End Function
