Attribute VB_Name = "uGlobals"
Option Explicit

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long
Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long

Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Type POINTAPI
    X As Long
    Y As Long
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
'----------- uControls --------------


'-------------- debug ---------------
Global uDontDrawDots As Boolean
Global uEnableMouseHooks As Boolean
'-------------- debug ---------------


'----------- uMouseWheel ------------
Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function GetWindowSubclass Lib "comctl32" Alias "#411" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, pdwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_NCDESTROY As Long = &H82 ' RemoveWindowsHook must be called prior to destruction.

Public Function HookSet(ByVal hWnd As Long, ByVal Thing As uMouseWheel, Optional dwRefData As Long) As Boolean ' http://msdn.microsoft.com/en-us/library/bb762102(VS.85).aspx
    If uEnableMouseHooks Then
        HookSet = CBool(SetWindowSubclass(hWnd, AddressOf SubclassProc, ObjPtr(Thing), dwRefData))
    Else
        HookSet = False
    End If
End Function

Public Function HookGetData(ByVal hWnd As Long, ByVal Thing As uMouseWheel) As Long ' http://msdn.microsoft.com/en-us/library/bb776430(VS.85).aspx
    Dim dwRefData As Long
    If GetWindowSubclass(hWnd, AddressOf SubclassProc, ObjPtr(Thing), dwRefData) Then
       HookGetData = dwRefData
    End If
End Function

Public Function HookClear(ByVal hWnd As Long, ByVal Thing As uMouseWheel) As Boolean ' http://msdn.microsoft.com/en-us/library/bb762094(VS.85).aspx
    HookClear = CBool(RemoveWindowSubclass(hWnd, AddressOf SubclassProc, ObjPtr(Thing)))
End Function

Public Function HookDefault(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long ' http://msdn.microsoft.com/en-us/library/bb776403(VS.85).aspx
    HookDefault = DefSubclassProc(hWnd, uiMsg, wParam, lParam)
End Function

Public Function SubclassProc(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As uMouseWheel, ByVal dwRefData As Long) As Long ' http://msdn.microsoft.com/en-us/library/bb776774(VS.85).aspx
    SubclassProc = uIdSubclass.Message(hWnd, uiMsg, wParam, lParam, dwRefData)
    If uiMsg = WM_NCDESTROY Then ' This should *never* be necessary, but just in case client fails to...
        Call HookClear(hWnd, uIdSubclass)
    End If
End Function
'----------- uMouseWheel ------------




Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) _
       As Long

    If Topmost = True Then    'Make the window topmost
        SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, _
                                        0, flags)
    Else
        SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, _
                                        0, 0, flags)
        SetTopMostWindow = False
    End If
End Function


