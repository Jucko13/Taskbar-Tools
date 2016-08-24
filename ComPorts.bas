Attribute VB_Name = "ComPorts"
Private Const INVALID_HANDLE_VALUE As Long = -1
'--- error codes
Private Const ERROR_ACCESS_DENIED As Long = 5&
Private Const ERROR_GEN_FAILURE As Long = 31&
Private Const ERROR_SHARING_VIOLATION As Long = 32&
Private Const ERROR_SEM_TIMEOUT As Long = 121&

Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As Long, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Function PrintError(sFunc As String)
    Debug.Print sFunc; ": "; Error
End Function

Public Function IsNT() As Boolean
    IsNT = True
End Function

Public Function EnumSerialPorts() As Long()
    Const FUNC_NAME As String = "EnumSerialPorts"
    Dim sBuffer As String
    Dim lIdx As Long
    Dim hFile As Long
    Dim vRet() As Long
    Dim lCount As Long

    On Error GoTo EH
    ReDim vRet(0 To 30) As Long
    '    If IsNT Then
    sBuffer = String$(30000, " ")
    Call QueryDosDevice(0, sBuffer, Len(sBuffer))
    sBuffer = Chr$(0) & sBuffer
    For lIdx = 1 To 30
        If InStr(1, sBuffer, Chr$(0) & "COM" & lIdx & Chr$(0), vbTextCompare) > 0 Then
            vRet(lCount) = lIdx
            lCount = lCount + 1
        End If
    Next
    '    Else
    '        For lIdx = 1 To 255
    '            hFile = CreateFile("COM" & lIdx, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
    '            If hFile = INVALID_HANDLE_VALUE Then
    '                Select Case Err.LastDllError
    '                Case ERROR_ACCESS_DENIED, ERROR_GEN_FAILURE, ERROR_SHARING_VIOLATION, ERROR_SEM_TIMEOUT
    '                    hFile = 0
    '                End Select
    '            Else
    '                Call CloseHandle(hFile)
    '                hFile = 0
    '            End If
    '            If hFile = 0 Then
    '                vRet(lCount) = "COM" & lIdx
    '                lCount = lCount + 1
    '            End If
    '        Next
    '    End If
    If lCount = 0 Then
        ReDim vRet(0) As Long
        EnumSerialPorts = vRet
    Else
        ReDim Preserve vRet(0 To lCount - 1) As Long
        EnumSerialPorts = vRet
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function
