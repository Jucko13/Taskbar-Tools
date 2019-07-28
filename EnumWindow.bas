Attribute VB_Name = "EnumWindow"
Option Explicit


Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public TopCount As Integer  ' Number of Top level Windows
Public ChildCount As Integer  ' Number of Child Windows
Public ThreadCount As Integer  ' Number of Thread Windows

Global MSTaskListWClass As Long
Global MSTaskSwWClass As Long

Function EnumWinProc(ByVal lhWnd As Long, ByVal lParam As Long) As Long
    Dim RetVal As Long, ProcessID As Long, ThreadID As Long
    Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
    Dim WinClass As String, WinTitle As String

    RetVal = GetClassName(lhWnd, WinClassBuf, 255)
    WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
    RetVal = GetWindowText(lhWnd, WinTitleBuf, 255)
    WinTitle = StripNulls(WinTitleBuf)

    'Debug.Print WinTitle & " : " & WinTitle

    TopCount = TopCount + 1
    ' see the Windows Class and Title for each top level Window
    'Debug.Print "Top level Class = "; WinClass; ", Title = "; WinTitle
    ' Usually either enumerate Child or Thread Windows, not both.
    ' In this example, EnumThreadWindows may produce a very long list!
    RetVal = EnumChildWindows(lhWnd, AddressOf EnumChildProc, lParam)
    ThreadID = GetWindowThreadProcessId(lhWnd, ProcessID)
    RetVal = EnumThreadWindows(ThreadID, AddressOf EnumThreadProc, lParam)
    EnumWinProc = True
End Function

Function EnumChildProc(ByVal lhWnd As Long, ByVal lParam As Long) As Long
    Dim RetVal As Long
    Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
    Dim WinClass As String, WinTitle As String
    Dim WinRect As RECT
    Dim WinWidth As Long, WinHeight As Long

    RetVal = GetClassName(lhWnd, WinClassBuf, 255)
    WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
    RetVal = GetWindowText(lhWnd, WinTitleBuf, 255)
    WinTitle = StripNulls(WinTitleBuf)
    ChildCount = ChildCount + 1
    ' see the Windows Class and Title for each Child Window enumerated
    'Debug.Print "   Child Class = "; WinClass; ", Title = "; WinTitle; " ,HWND = "; lhWnd

    EnumChildProc = True

    If WinClass = "MSTaskListWClass" Then  'Or WinClass = "ReBarWindow32"
        MSTaskListWClass = lhWnd
        EnumChildProc = False
    End If




    ' You can find any type of Window by searching for its WinClass
    '    If WinClass = "ThunderTextBox" Then    ' TextBox Window
    '        RetVal = GetWindowRect(lhWnd, WinRect)  ' get current size
    '        WinWidth = WinRect.Right - WinRect.Left    ' keep current width
    '        WinHeight = (WinRect.Bottom - WinRect.Top) * 2    ' double height
    '        RetVal = MoveWindow(lhWnd, 0, 0, WinWidth, WinHeight, True)
    '        EnumChildProc = False
    '    Else
    '        EnumChildProc = True
    '    End If
End Function

Function EnumThreadProc(ByVal lhWnd As Long, ByVal lParam As Long) As Long
    Dim RetVal As Long
    Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
    Dim WinClass As String, WinTitle As String

    RetVal = GetClassName(lhWnd, WinClassBuf, 255)
    WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
    RetVal = GetWindowText(lhWnd, WinTitleBuf, 255)
    WinTitle = StripNulls(WinTitleBuf)
    ThreadCount = ThreadCount + 1
    ' see the Windows Class and Title for top level Window
    'Debug.Print "Thread Window Class = "; WinClass; ", Title = "; WinTitle
    EnumThreadProc = True
End Function

Public Function StripNulls(OriginalStr As String) As String

' This removes the extra Nulls so String comparisons will work
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

