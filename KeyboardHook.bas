Attribute VB_Name = "KeyboardHook"
Option Explicit


Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
                                          ByVal idHook As Long, _
                                          ByVal lpfn As Long, _
                                          ByVal hmod As Long, _
                                          ByVal dwThreadId As Long _
                                          ) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
                                             ByVal hHook As Long _
                                             ) As Long

Private Declare Function CallNextHookEx Lib "user32" ( _
                                        ByVal hHook As Long, _
                                        ByVal ncode As Long, _
                                        ByVal wParam As Long, _
                                        lParam As Any _
                                        ) As Long


Private Const WH_KEYBOARD_LL = 13
Private Const WH_MOUSE_LL = 14
Private Const HC_ACTION = 0
Private Const HC_NOREMOVE = 3

Private Type KBDLLHOOKSTRUCT
    VKCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type


Const VK_H = 72
Const VK_E = 69
Const VK_L = 76
Const VK_O = 79
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


Private Declare Function OemKeyScan Lib "user32" (ByVal wOemChar As Integer) As _
                                    Long
Private Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal _
                                                                    lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar _
                                                                    As Byte) As Integer
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" _
                                       (ByVal wCode As Long, ByVal wMapType As Long) As Long

Private Const KEYEVENTF_KEYDOWN As Long = &H0

Type VKType
    VKCode As Integer
    scanCode As Integer
    Control As Boolean
    Shift As Boolean
    Alt As Boolean
End Type



Private hHook As Long
Private IsHooked As Boolean


Private bAlt As Boolean, bShift As Boolean, bControl As Boolean, bWindows As Boolean
Attribute bShift.VB_VarUserMemId = 1073741826
Attribute bControl.VB_VarUserMemId = 1073741826
Attribute bWindows.VB_VarUserMemId = 1073741826




Sub VbSendKeys(ByVal sKeystrokes As String)
    Dim iKeyStrokesLen As Integer
    Dim lRepetitions As Long
    Dim bShiftKey As Boolean
    Dim bControlKey As Boolean
    Dim bAltKey As Boolean
    Dim lResult As Long
    Dim sKey As String
    Dim iAsciiKey As Integer
    Dim iVirtualKey As Integer
    Dim i As Long
    Dim j As Long

    Static bInitialized As Boolean
    Static AsciiKeys(0 To 255) As VKType
    Static VirtualKeys(0 To 255) As VKType

    On Error GoTo 0

    If Not bInitialized Then
        Dim iKey As Integer
        Dim OEMChar As String
        Dim keyScan As Integer

        ' Initialize AsciiKeys()
        For iKey = LBound(AsciiKeys) To UBound(AsciiKeys)
            keyScan = VkKeyScan(iKey)
            AsciiKeys(iKey).VKCode = keyScan And &HFF   ' low-byte of key scan
            ' code
            AsciiKeys(iKey).Shift = (keyScan And &H100)
            AsciiKeys(iKey).Control = (keyScan And &H200)
            AsciiKeys(iKey).Alt = (keyScan And &H400)
            ' Get the ScanCode
            OEMChar = "  "    ' 2 Char
            CharToOem Chr(iKey), OEMChar
            AsciiKeys(iKey).scanCode = OemKeyScan(Asc(OEMChar)) And &HFF
        Next iKey

        ' Initialize VirtualKeys()
        For iKey = LBound(VirtualKeys) To UBound(VirtualKeys)
            VirtualKeys(iKey).VKCode = iKey
            VirtualKeys(iKey).scanCode = MapVirtualKey(iKey, 0)
            ' no use in initializing remaining elements
        Next iKey
        bInitialized = True     ' don't run this code twice
    End If    ' End of initialization routine

    ' Parse the string in the same way that SendKeys() would
    Do While Len(sKeystrokes)
        lRepetitions = 1    ' Default number of repetitions for each character
        bShiftKey = False
        bControlKey = False
        bAltKey = False

        ' Pull off Control, Alt or Shift specifiers
        sKey = Left$(sKeystrokes, 1)
        sKeystrokes = Mid$(sKeystrokes, 2)

        Do While InStr(" ^%+", sKey) > 1    ' The space in " ^%+" is necessary
            If sKey = "+" Then
                bShiftKey = True
            ElseIf sKey = "^" Then
                bControlKey = True
            ElseIf sKey = "%" Then
                bAltKey = True
            End If
            sKey = Left$(sKeystrokes, 1)
            sKeystrokes = Mid$(sKeystrokes, 2)
        Loop

        ' Look for "{}"
        If sKey = "{" Then
            ' Look for the  "}"
            i = InStr(sKeystrokes, "}")
            If i > 0 Then
                sKey = Left$(sKeystrokes, i - 1)    ' extract the content between
                ' the {}
                sKeystrokes = Mid$(sKeystrokes, i + 1)    ' Remove the }
            End If

            ' Look for repetitions
            i = Len(sKey)
            Do While Mid$(sKey, i, 1) >= "0" And Mid$(sKey, i, _
                                                      1) <= "9" And i >= 3
                i = i - 1
            Loop

            If i < Len(sKey) Then    ' If any digits were found...
                If i >= 2 Then    ' If there is something preceding it...
                    If Mid$(sKey, i, 1) = " " Then  ' If a space precedes the
                        ' digits...
                        On Error Resume Next    ' On overflow, ignore the value
                        lRepetitions = CLng(Mid$(sKey, i + 1))
                        On Error GoTo 0
                        sKey = Left$(sKey, i - 1)
                    End If
                End If
            End If
        End If

        ' Look for special words
        Select Case UCase$(sKey)
            Case "LBUTTON"    ' New
                iVirtualKey = vbKeyLButton
            Case "RBUTTON"    ' New
                iVirtualKey = vbKeyRButton
            Case "BREAK", "CANCEL"
                iVirtualKey = vbKeyCancel
            Case "MBUTTON"    ' New
                iVirtualKey = vbKeyMButton
            Case "BACKSPACE", "BS", "BKSP"
                iVirtualKey = vbKeyBack
            Case "TAB"
                iVirtualKey = vbKeyTab
            Case "CLEAR"    ' New
                iVirtualKey = vbKeyClear
            Case "ENTER", "~"
                iVirtualKey = vbKeyReturn
            Case "SHIFT"    ' New
                iVirtualKey = vbKeyShift
            Case "CONTROL"    ' New
                iVirtualKey = vbKeyControl
            Case "MENU", "ALT"    ' New
                iVirtualKey = vbKeyMenu
            Case "PAUSE"    ' New
                iVirtualKey = vbKeyPause
            Case "CAPSLOCK"
                iVirtualKey = vbKeyCapital
            Case "ESCAPE", "ESC"
                iVirtualKey = vbKeyEscape
            Case "SPACE"    ' New
                iVirtualKey = vbKeySpace
            Case "PGUP"
                iVirtualKey = vbKeyPageUp
            Case "PGDN"
                iVirtualKey = vbKeyPageDown
            Case "END"
                iVirtualKey = vbKeyEnd
            Case "HOME"    ' New
                iVirtualKey = vbKeyHome
            Case "LEFT"
                iVirtualKey = vbKeyLeft
            Case "UP"
                iVirtualKey = vbKeyUp
            Case "RIGHT"
                iVirtualKey = vbKeyRight
            Case "DOWN"
                iVirtualKey = vbKeyDown
            Case "SELECT"    ' New
                iVirtualKey = vbKeySelect
            Case "PRTSC"
                iVirtualKey = vbKeyPrint
            Case "EXECUTE"    ' New
                iVirtualKey = vbKeyExecute
            Case "SNAPSHOT"    ' New
                iVirtualKey = vbKeySnapshot
            Case "INSERT", "INS"
                iVirtualKey = vbKeyInsert
            Case "DELETE", "DEL"
                iVirtualKey = vbKeyDelete
            Case "HELP"
                iVirtualKey = vbKeyHelp
            Case "NUMLOCK"
                iVirtualKey = vbKeyNumlock
            Case "SCROLLLOCK"
                iVirtualKey = vbKeyScrollLock
            Case "NUMPAD0"    ' New
                iVirtualKey = vbKeyNumpad0
            Case "NUMPAD1"    ' New
                iVirtualKey = vbKeyNumpad1
            Case "NUMPAD2"    ' New
                iVirtualKey = vbKeyNumpad2
            Case "NUMPAD3"    ' New
                iVirtualKey = vbKeyNumpad3
            Case "NUMPAD4"    ' New
                iVirtualKey = vbKeyNumpad4
            Case "NUMPAD5"    ' New
                iVirtualKey = vbKeyNumpad5
            Case "NUMPAD6"    ' New
                iVirtualKey = vbKeyNumpad6
            Case "NUMPAD7"    ' New
                iVirtualKey = vbKeyNumpad7
            Case "NUMPAD8"    ' New
                iVirtualKey = vbKeyNumpad8
            Case "NUMPAD9"    ' New
                iVirtualKey = vbKeyNumpad9
            Case "MULTIPLY"    ' New
                iVirtualKey = vbKeyMultiply
            Case "ADD"    ' New
                iVirtualKey = vbKeyAdd
            Case "SEPARATOR"    ' New
                iVirtualKey = vbKeySeparator
            Case "SUBTRACT"    ' New
                iVirtualKey = vbKeySubtract
            Case "DECIMAL"    ' New
                iVirtualKey = vbKeyDecimal
            Case "DIVIDE"    ' New
                iVirtualKey = vbKeyDivide
            Case "F1"
                iVirtualKey = vbKeyF1
            Case "F2"
                iVirtualKey = vbKeyF2
            Case "F3"
                iVirtualKey = vbKeyF3
            Case "F4"
                iVirtualKey = vbKeyF4
            Case "F5"
                iVirtualKey = vbKeyF5
            Case "F6"
                iVirtualKey = vbKeyF6
            Case "F7"
                iVirtualKey = vbKeyF7
            Case "F8"
                iVirtualKey = vbKeyF8
            Case "F9"
                iVirtualKey = vbKeyF9
            Case "F10"
                iVirtualKey = vbKeyF10
            Case "F11"
                iVirtualKey = vbKeyF11
            Case "F12"
                iVirtualKey = vbKeyF12
            Case "F13"
                iVirtualKey = vbKeyF13
            Case "F14"
                iVirtualKey = vbKeyF14
            Case "F15"
                iVirtualKey = vbKeyF15
            Case "F16"
                iVirtualKey = vbKeyF16
            Case Else
                ' Not a virtual key
                iVirtualKey = -1
        End Select

        ' Turn on CONTROL, ALT and SHIFT keys as needed
        If bShiftKey Then
            keybd_event VirtualKeys(vbKeyShift).VKCode, _
                        VirtualKeys(vbKeyShift).scanCode, KEYEVENTF_KEYDOWN, 0
        End If

        If bControlKey Then
            keybd_event VirtualKeys(vbKeyControl).VKCode, _
                        VirtualKeys(vbKeyControl).scanCode, KEYEVENTF_KEYDOWN, 0
        End If

        If bAltKey Then
            keybd_event VirtualKeys(vbKeyMenu).VKCode, _
                        VirtualKeys(vbKeyMenu).scanCode, KEYEVENTF_KEYDOWN, 0
        End If

        ' Send the keystrokes
        For i = 1 To lRepetitions
            If iVirtualKey > -1 Then
                ' Virtual key
                keybd_event VirtualKeys(iVirtualKey).VKCode, _
                            VirtualKeys(iVirtualKey).scanCode, KEYEVENTF_KEYDOWN, 0
                keybd_event VirtualKeys(iVirtualKey).VKCode, _
                            VirtualKeys(iVirtualKey).scanCode, KEYEVENTF_KEYUP, 0
            Else
                ' ASCII Keys
                For j = 1 To Len(sKey)
                    iAsciiKey = Asc(Mid$(sKey, j, 1))
                    ' Turn on CONTROL, ALT and SHIFT keys as needed
                    If Not bShiftKey Then
                        If AsciiKeys(iAsciiKey).Shift Then
                            keybd_event VirtualKeys(vbKeyShift).VKCode, _
                                        VirtualKeys(vbKeyShift).scanCode, _
                                        KEYEVENTF_KEYDOWN, 0
                        End If
                    End If

                    If Not bControlKey Then
                        If AsciiKeys(iAsciiKey).Control Then
                            keybd_event VirtualKeys(vbKeyControl).VKCode, _
                                        VirtualKeys(vbKeyControl).scanCode, _
                                        KEYEVENTF_KEYDOWN, 0
                        End If
                    End If

                    If Not bAltKey Then
                        If AsciiKeys(iAsciiKey).Alt Then
                            keybd_event VirtualKeys(vbKeyMenu).VKCode, _
                                        VirtualKeys(vbKeyMenu).scanCode, _
                                        KEYEVENTF_KEYDOWN, 0
                        End If
                    End If

                    ' Press the key
                    keybd_event AsciiKeys(iAsciiKey).VKCode, _
                                AsciiKeys(iAsciiKey).scanCode, KEYEVENTF_KEYDOWN, 0
                    keybd_event AsciiKeys(iAsciiKey).VKCode, _
                                AsciiKeys(iAsciiKey).scanCode, KEYEVENTF_KEYUP, 0

                    ' Turn on CONTROL, ALT and SHIFT keys as needed
                    If Not bShiftKey Then
                        If AsciiKeys(iAsciiKey).Shift Then
                            keybd_event VirtualKeys(vbKeyShift).VKCode, _
                                        VirtualKeys(vbKeyShift).scanCode, _
                                        KEYEVENTF_KEYUP, 0
                        End If
                    End If

                    If Not bControlKey Then
                        If AsciiKeys(iAsciiKey).Control Then
                            keybd_event VirtualKeys(vbKeyControl).VKCode, _
                                        VirtualKeys(vbKeyControl).scanCode, _
                                        KEYEVENTF_KEYUP, 0
                        End If
                    End If

                    If Not bAltKey Then
                        If AsciiKeys(iAsciiKey).Alt Then
                            keybd_event VirtualKeys(vbKeyMenu).VKCode, _
                                        VirtualKeys(vbKeyMenu).scanCode, _
                                        KEYEVENTF_KEYUP, 0
                        End If
                    End If
                Next j    ' Each ASCII key
            End If  ' ASCII keys
        Next i    ' Repetitions

        ' Turn off CONTROL, ALT and SHIFT keys as needed
        If bShiftKey Then
            keybd_event VirtualKeys(vbKeyShift).VKCode, _
                        VirtualKeys(vbKeyShift).scanCode, KEYEVENTF_KEYUP, 0
        End If

        If bControlKey Then
            keybd_event VirtualKeys(vbKeyControl).VKCode, _
                        VirtualKeys(vbKeyControl).scanCode, KEYEVENTF_KEYUP, 0
        End If

        If bAltKey Then
            keybd_event VirtualKeys(vbKeyMenu).VKCode, _
                        VirtualKeys(vbKeyMenu).scanCode, KEYEVENTF_KEYUP, 0
        End If

    Loop    ' sKeyStrokes
End Sub



Public Sub SetKeyboardHook()
    If Not IsHooked Then
        hHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
        IsHooked = True
    End If
End Sub

Public Sub RemoveKeyboardHook()
    If IsHooked Then
        UnhookWindowsHookEx hHook
        IsHooked = False
    End If
End Sub


Public Function LowLevelKeyboardProc(ByVal uCode As Long, ByVal wParam As Long, lParam As KBDLLHOOKSTRUCT) As Long
    If uCode >= 0 And uCode = HC_ACTION Then
        If lParam.VKCode = 91 Then
            bWindows = (wParam = 256)
            'Debug.Print bWindows
        End If

        'Debug.Print "wind: " & bWindows & " key: " & lParam.VKCode & " chr():" & Chr(lParam.VKCode) & " Down?: " & (wParam = 256)

        If bWindows Then
            Select Case lParam.VKCode
                    '                Case 69
                    '                    LowLevelKeyboardProc = 1 'CallNextHookEx(hHook, uCode, wParam, lParam)
                    '
                    '                    keybd_event VK_Z, 0, 0, 0 ' press H
                    '                    keybd_event VK_Z, 0, KEYEVENTF_KEYUP, 0 ' release H
                    '
                    '                    'frmProgramTools.PenButtonPress_Single
                    '                    Exit Function
                    '
                Case 131    'single pen press
                    LowLevelKeyboardProc = 1

                    keybd_event vbKeyZ, 0, 0, 0
                    keybd_event vbKeyZ, 0, KEYEVENTF_KEYUP, 0
                    keybd_event 91, 0, KEYEVENTF_KEYUP, 0

                    frmProgramTools.PenButtonPress_Single
                    Exit Function

                Case 130    'double pen press
                    LowLevelKeyboardProc = 1

                    keybd_event vbKeyZ, 0, 0, 0
                    keybd_event vbKeyZ, 0, KEYEVENTF_KEYUP, 0
                    keybd_event 91, 0, KEYEVENTF_KEYUP, 0

                    frmProgramTools.PenButtonPress_Double
                    Exit Function

            End Select
        End If
    End If

    LowLevelKeyboardProc = CallNextHookEx(hHook, uCode, wParam, lParam)
End Function

