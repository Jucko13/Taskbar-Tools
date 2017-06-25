VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "uMouseWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Event onMouseWheel(direction As Long)
Private m_hWnd As Long

Private Const WM_NCDESTROY As Long = &H82
Private Const WM_XBUTTONDOWN As Long = &H20B         '(_WIN32_WINNT >= 0x0500)
Private Const WM_XBUTTONUP As Long = &H20C           '(_WIN32_WINNT >= 0x0500)
Private Const WM_XBUTTONDBLCLK As Long = &H20D       '(_WIN32_WINNT >= 0x0500)
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_MOUSEHWHEEL As Long = &H20E         '(_WIN32_WINNT >= 0x0600)
Private Const SM_MOUSEWHEELPRESENT As Long = 75
Private Const SPI_GETWHEELSCROLLLINES As Long = 104
Private Const SPI_SETWHEELSCROLLLINES As Long = 105
Private Const SPI_GETWHEELSCROLLCHARS As Long = 108  '(_WIN32_WINNT >= 0x0600)
Private Const SPI_SETWHEELSCROLLCHARS As Long = 109  '(_WIN32_WINNT >= 0x0600)



Private Sub Class_Terminate()
    Call Unhook
End Sub


Public Function WordHi(ByVal DWord As Long) As Integer
   Call CopyMemory(WordHi, ByVal VarPtr(DWord) + 2, 2)
End Function

Public Function WordLo(ByVal DWord As Long) As Integer
   Call CopyMemory(WordLo, DWord, 2)
End Function

Private Sub Unhook()
   ' Clear existing hook.
   If m_hWnd Then
      Call HookClear(m_hWnd, Me)
      m_hWnd = 0
   End If
End Sub


Public Property Let hWnd(ByVal NewValue As Long)
   ' Clear any existing hook, first.
   Call Unhook
   ' Set hook into new window, if possible.
   If IsWindow(NewValue) Then
      If HookSet(NewValue, Me) Then
         m_hWnd = NewValue
      End If
   End If
End Property

Public Property Get hWnd() As Long
   hWnd = m_hWnd  'report
End Property

Public Property Get ScrollLines() As Long
   Call SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0&, ScrollLines, 0&)
End Property


Public Function Message(ByVal hWnd As Long, _
                        ByVal uiMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long, _
                        ByVal dwRefData As Long) As Long
    Dim Delta As Long
    
    Select Case uiMsg
        Case WM_MOUSEWHEEL, WM_MOUSEHWHEEL
             Delta = WordHi(wParam)
             RaiseEvent onMouseWheel(ScrollLines * IIf(Delta > 0, 1, -1))
             
        Case WM_NCDESTROY
            Call Unhook  ' !!!
            
    End Select
    
    
    Message = HookDefault(hWnd, uiMsg, wParam, lParam)
    
End Function

