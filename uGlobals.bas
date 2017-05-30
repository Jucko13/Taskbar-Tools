Attribute VB_Name = "uGlobals"
Option Explicit

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long
Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long

Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Type POINTAPI
    x As Long
    y As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


'----------- uControls --------------
Global pts() As POINTAPI

Public Type Sel_Style
    FontName As String
    ForeColor As Long
    Underline As Boolean
    Italic As Boolean
    Bold As Boolean
End Type

'--------- debug drawing ------------

Global uDontDrawDots As Boolean
'----------- uControls --------------


Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
       As Long

    If Topmost = True Then    'Make the window topmost
        SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
                                        0, flags)
    Else
        SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
                                        0, 0, flags)
        SetTopMostWindow = False
    End If
End Function


