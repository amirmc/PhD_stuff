Attribute VB_Name = "Module1"

' Performance presentation v0.1 definitions

Public Declare Function StartTimer Lib "perfpres.dll" Alias "_StartTimer@0" () As Long
Public Declare Function GetTimer Lib "perfpres.dll" Alias "_GetTimer@0" () As Double
Public Declare Sub PumpUpTheThreadPriority Lib "perfpres.dll" Alias "_PumpUpTheThreadPriority@0" ()
Public Declare Sub RestoreThreadPriority Lib "perfpres.dll" Alias "_RestoreThreadPriority@0" ()
Public Declare Function WaitForTime Lib "perfpres.dll" Alias "_WaitForTime@8" (ByVal mytime As Double) As Double

'to deal with input and output to printer port
Public Declare Function inp Lib "inpout32.dll" _
Alias "Inp32" (ByVal PortAddress As Integer) As Integer
Public Declare Sub Out Lib "inpout32.dll" _
Alias "Out32" (ByVal PortAddress As Integer, ByVal Value As Integer)
