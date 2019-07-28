VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPUUsageNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type LARGE_INTEGER
    dwLow As Long
    dwHigh As Long
End Type
Private Type SYSTEM_BASIC_INFORMATION
    dwUnknown1 As Long
    uKeMaximumIncrement As Long
    uPageSize As Long
    uMmNumberOfPhysicalPages As Long
    uMmLowestPhysicalPage As Long
    uMmHighestPhysicalPage As Long
    uAllocationGranularity As Long
    pLowestUserAddress As Long
    pMmHighestUserAddress As Long
    uKeActiveProcessors As Long
    bKeNumberProcessors As Byte
    bUnknown2 As Byte
    wUnknown3 As Integer
End Type
Private Type SYSTEM_PERFORMANCE_INFORMATION
    liIdleTime As LARGE_INTEGER
    dwSpare(0 To 151) As Long
    '    dwSpare(0 To  75) As Long ' for OS ver < 6.0
    '    dwSpare(0 To 151) As Long ' for OS ver >= 6.0

End Type
'Private lArrLen As Long
Private Type SYSTEM_TIME_INFORMATION
    liKeBootTime As LARGE_INTEGER
    liKeSystemTime As LARGE_INTEGER
    liExpTimeZoneBias As LARGE_INTEGER
    uCurrentTimeZoneId As Long
    dwReserved As Long
End Type
Private Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal dwInfoType As Long, ByVal lpStructure As Long, ByVal dwSize As Long, ByVal dwReserved As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Const SYSTEM_BASICINFORMATION = 0&
Private Const SYSTEM_PERFORMANCEINFORMATION = 2&
Private Const SYSTEM_TIMEINFORMATION = 3&
Private Const NO_ERROR = 0

Private liOldIdleTime As LARGE_INTEGER
Private liOldSystemTime As LARGE_INTEGER

Public Sub Initialize()
    Dim SysTimeInfo As SYSTEM_TIME_INFORMATION
    Dim SysPerfInfo As SYSTEM_PERFORMANCE_INFORMATION
    Dim Ret As Long
    'get new system time
    Ret = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(SysTimeInfo), LenB(SysTimeInfo), 0&)
    If Ret <> NO_ERROR Then
        Debug.Print "Error while initializing the system's time!", vbCritical
        Exit Sub
    End If
    'get new CPU's idle time
    'lArrLen = (152 - CLng(IIf(WinNTVersion() >= 60, 152, 76))) * LenB(Ret)
    Ret = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(SysPerfInfo), LenB(SysPerfInfo), 0&)    ' - lArrLen
    'Debug.Print Hex(Ret)
    If Ret <> NO_ERROR Then
        Debug.Print "Error while initializing the CPU's idle time!", vbCritical
        End
        Exit Sub
    End If
    'store new CPU's idle and system time
    liOldIdleTime = SysPerfInfo.liIdleTime
    liOldSystemTime = SysTimeInfo.liKeSystemTime
End Sub
Public Function Query() As Long
    Dim SysBaseInfo As SYSTEM_BASIC_INFORMATION
    Dim SysPerfInfo As SYSTEM_PERFORMANCE_INFORMATION
    Dim SysTimeInfo As SYSTEM_TIME_INFORMATION
    Dim dbIdleTime As Currency
    Dim dbSystemTime As Currency
    Dim Ret As Long
    Query = -1
    'get number of processors in the system
    Ret = NtQuerySystemInformation(SYSTEM_BASICINFORMATION, VarPtr(SysBaseInfo), LenB(SysBaseInfo), 0&)
    If Ret <> NO_ERROR Then
        Debug.Print "Error while retrieving the number of processors!", vbCritical
        Exit Function
    End If
    'get new system time
    Ret = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(SysTimeInfo), LenB(SysTimeInfo), 0&)
    If Ret <> NO_ERROR Then
        Debug.Print "Error while retrieving the system's time!", vbCritical
        Exit Function
    End If
    'get new CPU's idle time
    Ret = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(SysPerfInfo), LenB(SysPerfInfo), ByVal 0&)    ' - lArrLen
    If Ret <> NO_ERROR Then
        Debug.Print "Error while retrieving the CPU's idle time!", vbCritical
        Exit Function
    End If
    'CurrentValue = NewValue - OldValue
    dbIdleTime = LI2Currency(SysPerfInfo.liIdleTime) - LI2Currency(liOldIdleTime)
    dbSystemTime = LI2Currency(SysTimeInfo.liKeSystemTime) - LI2Currency(liOldSystemTime)
    'CurrentCpuIdle = IdleTime / SystemTime
    If dbSystemTime <> 0 Then dbIdleTime = dbIdleTime / dbSystemTime
    'CurrentCpuUsage% = 100 - (CurrentCpuIdle * 100) / NumberOfProcessors
    dbIdleTime = 100 - dbIdleTime * 100 / SysBaseInfo.bKeNumberProcessors    ' + 0.5
    Query = Int(dbIdleTime)
    'store new CPU's idle and system time
    liOldIdleTime = SysPerfInfo.liIdleTime
    liOldSystemTime = SysTimeInfo.liKeSystemTime
End Function
Private Function LI2Currency(liInput As LARGE_INTEGER) As Currency
    CopyMemory LI2Currency, liInput, LenB(liInput)
End Function
Public Sub Terminate()
'nothing to do
End Sub

