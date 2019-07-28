Attribute VB_Name = "HotKey"
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As _
                                              Long, ByVal nCmdShow As Long) As Long

Global Const WM_KILLFOCUS = &H8
Global Const WM_SETFOCUS = &H7

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



Global ParentOrigional As Long
Global ParentTaskBar As Long
'Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

