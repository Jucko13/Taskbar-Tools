VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Declare Function QueryPerformanceFrequencyAny Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Any) As Long
Private Declare Function QueryPerformanceCounterAny Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Any) As Long


Private frequency As Currency
Private startTime As Currency
Private endTime As Currency
Private result As Double


Sub tStart()
    If QueryPerformanceFrequencyAny(frequency) = 0 Then
        MsgBox "This computer doesn't support high-res timers", vbCritical
        Exit Sub
    End If

    QueryPerformanceCounterAny startTime
End Sub

Function tStop() As Double
    QueryPerformanceCounterAny endTime

    ' note that both dividend and divisor are scaled
    ' by 10,000, so you don't need to scale the result
    tStop = (endTime - startTime) / frequency
End Function
