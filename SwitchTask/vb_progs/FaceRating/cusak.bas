Attribute VB_Name = "Module2"
' Performance presentation v0.1 definitions

Public Declare Function StartTimer Lib "perfpres.dll" Alias "_StartTimer@0" () As Long
Public Declare Function GetTimer Lib "perfpres.dll" Alias "_GetTimer@0" () As Double
Public Declare Sub PumpUpTheThreadPriority Lib "perfpres.dll" Alias "_PumpUpTheThreadPriority@0" ()
Public Declare Sub RestoreThreadPriority Lib "perfpres.dll" Alias "_RestoreThreadPriority@0" ()
Public Declare Function WaitForTime Lib "perfpres.dll" Alias "_WaitForTime@8" (ByVal mytime As Double) As Double

