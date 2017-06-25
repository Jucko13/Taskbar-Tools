VERSION 5.00
Begin VB.UserControl uTextBox 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaskColor       =   &H00FFFFFF&
   MousePointer    =   3  'I-Beam
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "uTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'TODO:
'when user presses up or down, move the cursor up and down
'add mouse text selection
'add events as:
'   text replacing
'   mousedown, up, move, click, dblclick
'   selection change
'   keyup, down, press
'
'add option to set the selection markuptext so you can add text in the middle of the textbox with some styles
'maybe add a mode for rtf?

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long



Private m_StrText As String
Private m_StrMarkupText As String
Private m_byteText() As Byte
Private m_byteMarkupText() As Byte

Private m_OleBackgroundColor As OLE_COLOR
Private m_OleForeColor As OLE_COLOR
Private m_OleBorderColor As OLE_COLOR
Private m_MouMousePointer As MousePointerConstants
Private m_StdStandardFont As New StdFont
Private m_StdFont As StdFont
Private m_bStarting As Boolean
Private m_bBorder As Boolean
Private m_lBorderThickness As Long


Private Enum Sel_Edit
    sFontName = 0
    sForeColor = 1
    sUnderline = 2
    sItalic = 3
    sBold = 4
    sMarking = 5
    sFontSize = 6
    sStrikeThrough = 7
    sLine = 8
    sNone = 254
End Enum

Private Type Current_Style
    lStyle As Sel_Edit
    prev_Value As Variant
End Type


Private Type WH
    W As Long 'char width
    H As Long 'char height
    d As Long 'divider height
    X As Long 'x position
    Y As Long 'y position
    r As Long 'belongs to what row?
    p As Long 'part of word
End Type

Private Type WHSL
    W As Long 'width
    H As Long 'height
    s As Long 'startChar
    l As Long 'length
End Type

Private Type NSS
    NumChars As Long
    StartY As Long
    startChar As Long
    Height As Long
End Type

Private Type MarkupStyles
    'lFontName As String
    lLine As Long
    lMarking As Long
    lFontSize As Long
    lForeColor As Long
    
    lStrikeThrough As Byte
    lUnderline As Byte
    lItalic As Byte
    lBold As Byte
    
End Type


Public Enum ScrollBarStyle
    lNone = 0
    lVertical = 1
    lHorizontal = 2
    lBoth = 3
End Enum


Private MarkupS() As MarkupStyles
Private CharMap() As WH    'width and hight of the characters
Private WordMap() As WHSL
Private WordCount As Long
Private RowMap() As NSS

Private m_bWordWrap As Boolean

Private m_lMouseX As Long
Private m_lMouseY As Long
Private m_lMouseDown As Long
Private m_lMouseDownX As Long
Private m_lMouseDownY As Long
Private m_lMouseDownPos As Long
Private m_lMouseDownPrevious As Long

Public m_CursorPos As Long
Private m_SelStart As Long
Private m_SelStartRow As Long
Private m_SelEnd As Long
Private m_SelEndRow As Long
Private m_SelUpDownTheSame As Boolean

Private m_bRefreshing As Boolean
Private m_bRefreshedWhileBusy As Boolean

Private m_bConsoleColors As Boolean
Private m_sConsoleColorBuffer As String

Private m_bLineNumbers As Boolean
Private m_bMarkupCalculated As Boolean
Private m_bMarkupCalculating As Boolean
Private m_bRowMapCalculated As Boolean

Private m_bWordsCalculated As Boolean
Private m_bWordsCalculating As Boolean

Private m_bHideCursor As Boolean
Private m_bLocked As Boolean

Private m_bMultiLine As Boolean
Private m_bRowLines As Boolean
Private m_bAutoResize As Boolean

Private m_OleRowLineColor As OLE_COLOR
Private m_OleLineNumberBackground As OLE_COLOR
Private m_OleLineNumberForeColor As OLE_COLOR
Private m_bRowNumberOnEveryLine As Boolean

Private m_bHasFocus As Boolean


Private WithEvents m_uMouseWheel As uMouseWheel
Attribute m_uMouseWheel.VB_VarHelpID = -1
Private m_sScrollBars As ScrollBarStyle

Private m_lScrollLeft As Long
Private m_lScrollLeftMax As Long
Private m_lScrollTop As Long
Private m_lScrollTopMax As Long
Private m_lScrollTopBarHeight As Long
Private m_lScrollTopHeight As Long
Private m_lScrollTopBarY As Long
Private m_lScrollTopDragStartY As Long
Private m_lScrollTopDragStartValue As Long
Private m_bScrollingTopBar As Boolean

'Private m_timer As clsTimer

Public Event Changed()
Public Event SelectionChanged()
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
Public Event KeyUp(ByRef KeyCode As Integer, ByRef Shift As Integer)
Public Event Click(ByVal charIndex As Long, ByVal charRow As Long)
Public Event OnCursorPositionChanged(ByVal charIndex As Long, ByVal charRow As Long, ByVal charCol As Long, ByVal charVal As Byte)

Private m_lUsercontrolHeight As Long
Private m_lUsercontrolWidth As Long
Private m_lUsercontrolLeft As Long
Private m_lUsercontrolTop As Long

Private UW As Long      'usercontrol width without scrollbars
Private UWS As Long     'usercontrol width
Private UH As Long      'usercontrol height without scrollbars
Private UHS As Long     'usercontrol height
Private TSP As Long     'text spacing
Private SYT As Long     'ScrollYTop

Private m_lRefreshFromCharAt As Long
Private m_lRefreshFromRowAt As Long

Private performance As PerformanceTimer


Private m_bBlockNextKeyPress As Boolean 'for things like ctrl+space autocomplete

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As WH) As Long
Private Declare Function SetTextAlign Lib "gdi32.dll" (ByVal hdc As Long, ByVal wFlags As Long) As Long

Private Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function setCaretPos Lib "user32" Alias "SetCaretPos" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long

Private m_OleConsoleColors(0 To 7) As OLE_COLOR


Public Property Get RawText() As Byte()
    RawText = m_byteText
End Property

Public Function getWordFromChar(Char As Long) As Long
    getWordFromChar = CharMap(Char).p
End Function

Public Function getWordLength(word As Long) As Long
    getWordLength = WordMap(word).l
End Function

Public Function getWordStart(word As Long) As Long
    getWordStart = WordMap(word).s
End Function

Public Sub ClearMarkup()
    Dim i As Long
    
    For i = 0 To UBound(MarkupS)
        With MarkupS(i)
            .lBold = 255
            .lFontSize = -1
            .lForeColor = -1
            .lItalic = 255
            .lLine = -1
            .lMarking = -1
            .lStrikeThrough = 255
            .lUnderline = 255
        End With
    Next i
    'MarkupS(Char).lItalic = bValue
End Sub


Public Sub setCharItallic(Char As Long, bValue As Byte)
    MarkupS(Char).lItalic = bValue
End Sub

Public Sub setCharBold(Char As Long, bValue As Byte)
    MarkupS(Char).lBold = bValue
    CheckCharSize Char, 1
    
    m_lRefreshFromCharAt = Char
    m_bWordsCalculated = False
    m_bRowMapCalculated = False
    
    If Not m_bStarting Then Redraw
    
End Sub

Public Sub setCharForeColor(Char As Long, OleValue As OLE_COLOR)
    MarkupS(Char).lForeColor = IIf(OleValue >= 0, OleValue, -1)
End Sub

Public Sub setCharBackColor(Char As Long, OleValue As OLE_COLOR)
    MarkupS(Char).lMarking = IIf(OleValue >= 0, OleValue, -1)
End Sub

Public Sub setCharBorderColor(Char As Long, OleValue As OLE_COLOR)
    MarkupS(Char).lLine = IIf(OleValue >= 0, OleValue, -1)
End Sub

Public Function getCharItallic(Char As Long) As Byte
    getCharItallic = IIf(MarkupS(Char).lItalic = 255, m_StdFont.Italic, CBool(MarkupS(Char).lItalic))
End Function

Public Function getCharBold(Char As Long) As Byte
    getCharBold = IIf(MarkupS(Char).lBold = 255, m_StdFont.Bold, CBool(MarkupS(Char).lBold))
End Function

Public Function getCharForeColor(Char As Long) As OLE_COLOR
    getCharForeColor = MarkupS(Char).lForeColor
End Function

Public Function getCharBackColor(Char As Long) As OLE_COLOR
   getCharBackColor = MarkupS(Char).lMarking
End Function


Sub updateCaretPos()
    If Not m_bHasFocus Then Exit Sub
    
    If Not Screen.ActiveControl Is Nothing Then
        If Not UserControl.Extender Is Screen.ActiveControl Then
            DestroyCaret
            Exit Sub
        End If
    End If
    
    If m_bHideCursor Then Exit Sub
    
    CreateCaret UserControl.hWnd, 0, 1, CharMap(m_CursorPos).H

    setCaretPos CharMap(m_CursorPos).X, CharMap(m_CursorPos).Y - CharMap(m_CursorPos).H + CharMap(m_CursorPos).d - SYT
    ShowCaret UserControl.hWnd
    
    RaiseEvent OnCursorPositionChanged(m_CursorPos, CharMap(m_CursorPos).r, m_CursorPos - RowMap(CharMap(m_CursorPos).r).startChar, m_byteText(m_CursorPos))
End Sub


Private Sub GetTextSize(pstrText As String, ByRef charsize As WH)
    GetTextExtentPoint32 UserControl.hdc, pstrText, 1, charsize    'lSize
End Sub


Public Property Get AutoResize() As Boolean
    AutoResize = m_bAutoResize
End Property

Public Property Let AutoResize(ByVal bValue As Boolean)
    m_bAutoResize = bValue
    PropertyChanged "AutoResize"
    If Not m_bStarting Then Redraw
End Property


Public Property Get RowNumberOnEveryLine() As Boolean
    RowNumberOnEveryLine = m_bRowNumberOnEveryLine
End Property

Public Property Let RowNumberOnEveryLine(ByVal bValue As Boolean)
    m_bRowNumberOnEveryLine = bValue
    PropertyChanged "RowNumberOnEveryLine"
    If Not m_bStarting Then Redraw
End Property


Public Property Get HideCursor() As Boolean
    HideCursor = m_bHideCursor
End Property

Public Property Let HideCursor(ByVal bValue As Boolean)
    m_bHideCursor = bValue
    PropertyChanged "HideCursor"
    If Not m_bStarting Then Redraw
    
    updateCaretPos
End Property


Public Property Get ConsoleColors() As Boolean
    ConsoleColors = m_bConsoleColors
End Property

Public Property Let ConsoleColors(ByVal bValue As Boolean)
    m_bConsoleColors = bValue
    PropertyChanged "ConsoleColors"
End Property


Public Property Get MultiLine() As Boolean
    MultiLine = m_bMultiLine
End Property

Public Property Let MultiLine(ByVal bValue As Boolean)
    m_bMultiLine = bValue
    PropertyChanged "MultiLine"
    If Not m_bStarting Then Redraw
End Property

Public Property Get LineNumberForeColor() As OLE_COLOR
    LineNumberForeColor = m_OleLineNumberForeColor
End Property

Public Property Let LineNumberForeColor(ByVal OleValue As OLE_COLOR)
    m_OleLineNumberForeColor = OleValue
    PropertyChanged "LineNumberForeColor"
    If Not m_bStarting Then Redraw
End Property


Public Property Get LineNumberBackground() As OLE_COLOR
    LineNumberBackground = m_OleLineNumberBackground
End Property

Public Property Let LineNumberBackground(ByVal OleValue As OLE_COLOR)
    m_OleLineNumberBackground = OleValue
    PropertyChanged "LineNumberBackground"
    If Not m_bStarting Then Redraw
End Property

Public Property Get RowLineColor() As OLE_COLOR
    RowLineColor = m_OleRowLineColor
End Property

Public Property Let RowLineColor(ByVal OleValue As OLE_COLOR)
    m_OleRowLineColor = OleValue
    PropertyChanged "RowLineColor"
    If Not m_bStarting Then Redraw
End Property

Public Property Get RowLines() As Boolean
    RowLines = m_bRowLines
End Property

Public Property Let RowLines(ByVal bValue As Boolean)
    m_bRowLines = bValue
    PropertyChanged "RowLines"
    If Not m_bStarting Then Redraw
End Property



Public Property Get ScrollBars() As ScrollBarStyle
    ScrollBars = m_sScrollBars
End Property

Public Property Let ScrollBars(ByVal sValue As ScrollBarStyle)
    m_sScrollBars = sValue
    PropertyChanged "ScrollBars"
    CalculateUserControlWidthHeight
    If Not m_bStarting Then Redraw
End Property



Public Property Get WordWrap() As Boolean
    WordWrap = m_bWordWrap
End Property

Public Property Let WordWrap(ByVal bValue As Boolean)
    m_bWordWrap = bValue
    PropertyChanged "WordWrap"
    If Not m_bStarting Then Redraw
End Property



Public Property Get LineNumbers() As Boolean
    LineNumbers = m_bLineNumbers
End Property

Public Property Let LineNumbers(ByVal bValue As Boolean)
    m_bLineNumbers = bValue
    PropertyChanged "LineNumbers"
    CalculateUserControlWidthHeight
    If Not m_bStarting Then Redraw
End Property


Public Property Let SelBold(bValue As Boolean)
'ReDim m_SelBold(0 To m_LonBoldCount)

'With m_SelBold(m_LonBoldCount)
'    .bValue = bValue
'    .lLength = m_SelCurrent.lLength
'    .lStart = m_SelCurrent.lStart
'End With

'm_LonBoldCount = m_LonBoldCount + 1

    If Not m_bStarting Then Redraw
End Property

Public Property Get SelStart() As Long
    SelStart = m_SelStart
End Property

Public Property Let SelStart(LonValue As Long)
    If LonValue < 0 Or LonValue > UBound(CharMap) Then Exit Property
    
    m_SelStart = LonValue
    m_SelEnd = m_SelStart
    m_CursorPos = m_SelStart
     
    If Not m_bStarting Then Redraw
    
    updateCaretPos
End Property

Public Property Let SelLength(LonValue As Long)
    Dim tmpswap As Long
    
    m_SelEnd = m_SelStart + LonValue
    
    If m_SelEnd > UBound(CharMap) Then m_SelEnd = UBound(CharMap)
    If m_SelEnd < 0 Then m_SelEnd = 0
    
    If m_SelEnd < m_SelStart Then
        tmpswap = m_SelEnd
        m_SelEnd = m_SelStart
        m_SelStart = tmpswap
    End If
    
    m_CursorPos = m_SelEnd
    
    If Not m_bStarting Then Redraw
    
    updateCaretPos
End Property

Public Property Get SelLength() As Long
    SelLength = m_SelEnd - m_SelStart
End Property

Public Property Let MarkupText(StrValue As String)
    m_StrMarkupText = StrValue

    If Not m_bStarting Then Redraw
End Property


Public Function ByteArrayToString(ByRef bytArray() As Byte) As String
    Dim sAns As String
    Dim iPos As Long
    
    sAns = Left$(StrConv(bytArray, vbUnicode), UBound(bytArray))
    ByteArrayToString = sAns
    
End Function

Public Property Get TextLength() As Long
    TextLength = UBound(m_byteText)
End Property



Public Property Get Text() As String
    Text = ByteArrayToString(m_byteText)
    
End Property


Public Property Let Text(ByVal StrValue As String)
    Clear
    AddCharAtCursor StrValue
    
    PropertyChanged "Text"
    If Not m_bStarting Then Redraw
    
    updateCaretPos
End Property


Public Property Get Font() As StdFont
    Set Font = m_StdFont
End Property

Public Property Set Font(ByVal StdValue As StdFont)
    Set m_StdFont = StdValue
    UserControl.Font = m_StdFont
    PropertyChanged "Font"
    m_bMarkupCalculated = False
    If Not m_bStarting Then Redraw
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = m_MouMousePointer
End Property

Public Property Let MousePointer(ByVal MouValue As MousePointerConstants)
    m_MouMousePointer = MouValue
    PropertyChanged "MousePointer"
End Property


Public Property Get BorderThickness() As Long
    BorderThickness = m_lBorderThickness
End Property

Public Property Let BorderThickness(ByVal lValue As Long)
    m_lBorderThickness = lValue
    PropertyChanged "BorderThickness"
    If Not m_bStarting Then Redraw
End Property


Public Property Get Border() As Boolean
    Border = m_bBorder
End Property

Public Property Let Border(ByVal bValue As Boolean)
    m_bBorder = bValue
    PropertyChanged "Border"
    If Not m_bStarting Then Redraw
End Property



Public Property Get Locked() As Boolean
    Locked = m_bLocked
End Property

Public Property Let Locked(ByVal bValue As Boolean)
    m_bLocked = bValue
    PropertyChanged "Locked"
    'If Not m_bStarting Then Redraw
End Property




Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_OleBorderColor
End Property

Public Property Let BorderColor(ByVal OleValue As OLE_COLOR)
    m_OleBorderColor = OleValue
    PropertyChanged "BorderColor"
    If Not m_bStarting Then Redraw
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_OleForeColor
End Property

Public Property Let ForeColor(ByVal OleValue As OLE_COLOR)
    m_OleForeColor = OleValue
    PropertyChanged "ForeColor"
    If Not m_bStarting Then Redraw
End Property


Public Property Get BackgroundColor() As OLE_COLOR
    BackgroundColor = m_OleBackgroundColor
End Property

Public Property Let BackgroundColor(ByVal OleValue As OLE_COLOR)
    m_OleBackgroundColor = OleValue
    PropertyChanged "BackgroundColor"
    If Not m_bStarting Then Redraw
End Property

Sub RedrawPause()
    m_bStarting = True
End Sub

Sub RedrawResume(Optional bDoNotRedraw As Boolean = False)
    m_bStarting = False
    If Not bDoNotRedraw Then Redraw
    
    updateCaretPos
End Sub

Function hWnd() As Long
    hWnd = UserControl.hWnd
End Function

Private Sub m_uMouseWheel_onMouseWheel(direction As Long)
    Debug.Print direction
    
    m_lScrollTop = m_lScrollTop - direction
    If m_lScrollTop < 0 Then m_lScrollTop = 0
    If m_lScrollTop > UBound(RowMap) Then m_lScrollTop = UBound(RowMap)
    If Not m_bStarting Then Redraw
End Sub

Private Sub UserControl_DblClick()
    Dim word As Long
    
    word = CharMap(m_CursorPos).p
    
    If word = -1 And m_CursorPos > 0 Then
        word = CharMap(m_CursorPos - 1).p
    End If
        
    If word <> -1 Then
        m_SelStart = WordMap(word).s
        m_SelEnd = WordMap(word).s + WordMap(word).l
        If m_SelEnd > UBound(CharMap) Then m_SelEnd = UBound(CharMap)
        m_CursorPos = m_SelEnd
        If Not m_bStarting Then Redraw
        updateCaretPos
    End If
End Sub

Private Sub UserControl_GotFocus()
    m_bHasFocus = True
    updateCaretPos
End Sub

Function FileToString(strFilename As String) As String
  Dim iFile As Long
  
  iFile = FreeFile
  Open strFilename For Input As #iFile
    FileToString = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
End Function

Private Sub UserControl_Initialize()
    m_bStarting = True
    
    m_OleConsoleColors(0) = vbBlack
    m_OleConsoleColors(1) = vbRed
    m_OleConsoleColors(2) = vbGreen
    m_OleConsoleColors(3) = vbYellow
    m_OleConsoleColors(4) = vbBlue
    m_OleConsoleColors(5) = vbMagenta
    m_OleConsoleColors(6) = vbCyan
    m_OleConsoleColors(7) = vbWhite
    
    'Set m_timer = New clsTimer

    'Dim lrand As Long
    Dim newChar As String

    Dim i As Long
    'Dim MS As String 'mid string

    Dim constString As String
    Const randomMarkup As Boolean = False
    
    
    'constString = FileToString("changelog.txt")

    'For i = 0 To 5
    '    constString = constString & "This textbox is made by Ricardo de Roode HereIsAVeryLongWord." & vbCrLf    '& vbCrLf
    'Next i


    If randomMarkup Then
        For i = 1 To Len(constString)
            newChar = ""
            '{\c FFFF00 hoi {\c FF00FF hallo dit is magenta gekleurde text} hoi}
            If Mid$(constString, i, 1) <> " " And Mid$(constString, i, 1) <> vbCr And Mid$(constString, i, 1) <> vbLf Then
                newChar = "{\c " & Fmat(Hex(CLng(Rnd * 255)), 2) & Fmat(Hex(CLng(Rnd * 255)), 2) & Fmat(Hex(CLng(Rnd * 255)), 2) & " "
                newChar = newChar & "{\fb " & Fmat(Hex(CLng(Rnd * 255)), 2) & Fmat(Hex(CLng(Rnd * 255)), 2) & Fmat(Hex(CLng(Rnd * 255)), 2) & " "
                newChar = newChar & "{\m " & Fmat(Hex(CLng(Rnd * 255)), 2) & Fmat(Hex(CLng(Rnd * 255)), 2) & Fmat(Hex(CLng(Rnd * 255)), 2) & " "
                'newChar = newChar & "{\m FF00FF "
                newChar = newChar & "{\fs " & Fix(Rnd * 16 + 8) & " "
                'newChar = newChar & "{\i "

                '            Select Case Round(Rnd * 3)
                '                Case 0
                '                    newChar = newChar & "{\i "
                '                Case 1
                '                    newChar = newChar & "{\b "
                '                Case 2
                '                    newChar = newChar & "{\u "
                '                Case 3
                '                    newChar = newChar & "{\s "
                '            End Select

                Select Case Mid$(constString, i, 1)
                    Case "}", "{", "\"
                        newChar = newChar & "\" & Mid$(constString, i, 1)

                    Case Else
                        newChar = newChar & Mid$(constString, i, 1)
                End Select

                'SnewChar = newChar & "}"
                newChar = newChar & "}"
                'newChar = newChar & "}"
                newChar = newChar & "}"
                newChar = newChar & "}"
                newChar = newChar & "}"

                m_StrMarkupText = m_StrMarkupText & newChar
            Else
                m_StrMarkupText = m_StrMarkupText & Mid$(constString, i, 1)

            End If




        Next i
    Else
        m_StrMarkupText = constString

    End If


    m_OleRowLineColor = &HEEEEEE
    m_bRowLines = False
    m_bConsoleColors = True
    
    m_bLineNumbers = False
    m_OleLineNumberBackground = 0
    m_OleLineNumberForeColor = vbWhite
    m_bRowNumberOnEveryLine = False
    m_lMouseDownPrevious = 99
    m_lBorderThickness = 1
    TSP = 6
    m_bLocked = False
    ReDim RowMap(0 To 0)
    
    'Debug.Print "initialize"
    'CalculateUserControlWidthHeight
    
    Set m_uMouseWheel = New uMouseWheel
    m_uMouseWheel.hWnd = UserControl.hWnd
    
    Set performance = New PerformanceTimer
End Sub

Sub DrawScrollBars()
    Dim d1 As Double
    Dim d2 As Double
    Dim d3 As Double
    
    Dim scrollArea As Long
    Dim scrollPosition As Long
    
    'UH = UserControl.ScaleHeight
    'UW = UserControl.ScaleWidth
    
    d1 = UWS / 15
    d2 = d1 * 1.73205
    d3 = d1 * 3
    
    If m_sScrollBars = lVertical Or m_sScrollBars = lBoth Then
        
        pts(0).X = UW - UWS
        pts(0).Y = 0

        pts(1).X = UW - 1
        pts(1).Y = 0
        
        pts(2).X = UW - 1
        pts(2).Y = UH - 1

        pts(3).X = UW - UWS
        pts(3).Y = UH - 1
    
        UserControl.FillColor = m_OleBackgroundColor
        Polygon UserControl.hdc, pts(0), 4
            
        UserControl.Line (UW - UWS, UH - UHS)-(UW, UH - UHS), m_OleBorderColor     'bottom
        UserControl.Line (UW - UWS, UHS - 1)-(UW, UHS - 1), m_OleBorderColor    'top
        
        'triangle bottom
        UserControl.Line (Fix(UW - UWS / 2 - d3), Fix(UH - UHS / 2 - d2))-(Fix(UW - UWS / 2 + d3), Fix(UH - UHS / 2 - d2)) '_
        UserControl.Line (Fix(UW - UWS / 2 + d3), Fix(UH - UHS / 2 - d2))-(Fix(UW - UWS / 2 - 1), Fix(UH - UHS / 2 + d2)) ' /
        UserControl.Line (Fix(UW - UWS / 2 - d3), Fix(UH - UHS / 2 - d2))-(Fix(UW - UWS / 2 + 1), Fix(UH - UHS / 2 + d2)) '\
        

        'triangle top
        UserControl.Line (Fix(UW - UWS / 2 - d3), Fix(UHS / 2 + d2))-(Fix(UW - UWS / 2 + 1), Fix(UHS / 2 - d2)) '/
        UserControl.Line (Fix(UW - UWS / 2 + d3), Fix(UHS / 2 + d2))-(Fix(UW - UWS / 2 - 1), Fix(UHS / 2 - d2)) ' \
        UserControl.Line (Fix(UW - UWS / 2 - d3), Fix(UHS / 2 + d2))-(Fix(UW - UWS / 2 + d3), Fix(UHS / 2 + d2)) '_
        

        m_lScrollTopHeight = (UH - (UHS * 2)) + 1
        m_lScrollTopBarHeight = m_lScrollTopHeight
        
        If m_lScrollTopMax > 0 Then
            'If m_lScrollTopMax >= 30 Then
                m_lScrollTopBarHeight = m_lScrollTopBarHeight / 20
            'Else
            '    m_lScrollTopBarHeight = m_lScrollTopBarHeight / m_lScrollTopMax
            'End If
            
            If m_lScrollTopBarHeight < 30 Then m_lScrollTopBarHeight = 30
            
            If m_lScrollTop > 0 Then scrollPosition = (m_lScrollTopHeight - m_lScrollTopBarHeight) / (m_lScrollTopMax / m_lScrollTop)
        End If
        
        m_lScrollTopBarY = UHS + scrollPosition - 1
        
        'draggable block
        pts(0).X = UW - UWS
        pts(0).Y = m_lScrollTopBarY

        pts(1).X = UW - 1
        pts(1).Y = m_lScrollTopBarY
        
        pts(2).X = UW - 1
        pts(2).Y = m_lScrollTopBarY + m_lScrollTopBarHeight

        pts(3).X = UW - UWS
        pts(3).Y = m_lScrollTopBarY + m_lScrollTopBarHeight
        
        UserControl.FillColor = m_OleLineNumberBackground
        Polygon UserControl.hdc, pts(0), 4
        
    End If
    

    If m_sScrollBars = lHorizontal Or m_sScrollBars = lBoth Then
        
        UserControl.Line (UWS, UH - UHS)-(UWS, UH), m_OleForeColor
        
        If m_sScrollBars = lBoth Then
            UserControl.Line (0, UH - UHS)-(UW - UWS, UH - UHS), m_OleForeColor
            UserControl.Line (UW - UWS - UWS, UH - UHS)-(UW - UWS - UWS, UH), m_OleForeColor
            
            'triangle right
            UserControl.Line (Fix(UW - UWS - UWS / 2 - d2), Fix(UH - UHS / 2 - d3))-(Fix(UW - UWS - UWS / 2 - d2), Fix(UH - UHS / 2 + d3)) ' |
            UserControl.Line (Fix(UW - UWS - UWS / 2 - d2), Fix(UH - UHS / 2 - d3))-(Fix(UW - UWS - UWS / 2 + d2), Fix(UH - UHS / 2 + 1)) '/
            UserControl.Line (Fix(UW - UWS - UWS / 2 - d2), Fix(UH - UHS / 2 + d3))-(Fix(UW - UWS - UWS / 2 + d2), Fix(UH - UHS / 2 - 1))  '\
               
        Else
            UserControl.Line (UW - UWS, UH - UHS)-(UW - UWS, UH), m_OleForeColor
            UserControl.Line (0, UH - UHS)-(UW, UH - UHS), m_OleForeColor
            
            'triangle right
            UserControl.Line (Fix(UW - UWS / 2 - d2), Fix(UH - UHS / 2 - d3))-(Fix(UW - UWS / 2 - d2), Fix(UH - UHS / 2 + d3))   ' |
            UserControl.Line (Fix(UW - UWS / 2 - d2), Fix(UH - UHS / 2 - d3))-(Fix(UW - UWS / 2 + d2), Fix(UH - UHS / 2 + 1))   '/
            UserControl.Line (Fix(UW - UWS / 2 - d2), Fix(UH - UHS / 2 + d3))-(Fix(UW - UWS / 2 + d2), Fix(UH - UHS / 2 - 1))    '\
        End If
        
        
        If m_lScrollLeftMax > 0 Then 'bar
            pts(0).X = UWS + 2
            pts(0).Y = UH - UHS + 2
    
            pts(1).X = pts(0).X
            pts(1).Y = UH - 3
            
            pts(2).X = (UW - UWS * IIf(m_sScrollBars = lBoth, 3, 2) - 3) - (UW - UWS * IIf(m_sScrollBars = lBoth, 3, 2) - 3) * (1 / (m_lScrollLeftMax + UW) * m_lScrollLeftMax)
            If pts(2).X < 10 Then pts(2).X = 10
            pts(2).X = pts(2).X + pts(0).X
            
            pts(2).Y = pts(1).Y
    
            pts(3).X = pts(2).X
            pts(3).Y = pts(0).Y
            
            Polygon UserControl.hdc, pts(0), 4
        End If
        
        'triangle left
        UserControl.Line (Fix(UWS / 2 + d2), Fix(UH - UHS / 2 - d3))-(Fix(UWS / 2 + d2), Fix(UH - UHS / 2 + d3)) ' |
        UserControl.Line (Fix(UWS / 2 + d2), Fix(UH - UHS / 2 - d3))-(Fix(UWS / 2 - d2), Fix(UH - UHS / 2 + 1)) '/
        UserControl.Line (Fix(UWS / 2 + d2), Fix(UH - UHS / 2 + d3))-(Fix(UWS / 2 - d2), Fix(UH - UHS / 2 - 1))  '\
        
    End If
    
End Sub

Sub growRowMap()
    Dim newSize As Long
    newSize = (UBound(RowMap) + 1) * 2
    ReDim Preserve RowMap(0 To newSize)
End Sub

Sub growWordMap()
    Dim newSize As Long
    newSize = (UBound(WordMap) + 1) * 2
    ReDim Preserve WordMap(0 To newSize)
End Sub

Sub ReCalculateRowMap(Optional fromWhere As Long = 0)
    Dim i As Long
    'Dim WC As Long 'word count
    Dim TL As Long 'text length
    Dim CC As Long
    
    Dim TW As Long    'text width
    Dim LNW As Long    'line number width
    Dim LNR As Long    'line number right
    Dim TextOffsetX As Long
    Dim TextOffsetY As Long
    Dim NRC As Long    'Number Row Count
    
    Dim RH As Long    'row height
    Dim RD As Long    'row d height
    
    Dim RL As Long    'row loop
    Dim TTW As Long    'temp text width
    Dim MTW As Long   'max text width

    Dim NLNR As Boolean    'Next Loop goto NextRow
    Dim POWC As Long    'part of word checked
    
    If fromWhere <= 0 Then
        ReDim RowMap(0)
        fromWhere = 0
        
        'TextOffsetY
        
    Else
        NRC = fromWhere
        TextOffsetY = RowMap(NRC).StartY
        
        RH = RowMap(NRC).Height
        fromWhere = RowMap(NRC).startChar
        RowMap(NRC).NumChars = 0
        'RowMap(NRC).startChar = RowMap(NRC - 1).startChar + RowMap(NRC).NumChars
        
    End If
'
'    If m_lScrollTop - 1 >= 0 Then
'        SYT = CharMap(RowMap(m_lScrollTop - 1).StartChar).y
'        TW = TextWidth(m_lScrollTop & "0")
'    Else
'        SYT = 0
'        TW = TextWidth("00")
'    End If
    
    TW = TextWidth("00000")
    
    
    POWC = -1
    
    LNW = 0
    LNR = 0
    If m_bLineNumbers Then    'draw the container for the line numbers
        LNR = TW + TSP
        LNW = LNR + TSP
        LNW = LNW + TSP
        TextOffsetX = LNW

    Else
        LNW = TSP
        TextOffsetX = TSP
    End If
    
    If m_lScrollLeft < 0 Then m_lScrollLeft = 0
    
    TextOffsetX = TextOffsetX - m_lScrollLeft
    
    If UW - UWS - LNW < 5 Then Exit Sub
    
    ReDim pts(0 To 3)
    'ReDim RowMap(0 To 200)
    
    TTW = LNW
    'RH = 0
    'RD = 0
    
    
    For CC = fromWhere To UBound(m_byteText)

        
        If NLNR = True Or CC = 0 Then
            GoTo MakeNewRule
        End If
        
checkNextChar:



        Select Case m_byteText(CC)
            Case 13
                If m_bMultiLine Then NLNR = True
                CharMap(CC).r = NRC
                
            Case 10
                CharMap(CC).X = TextOffsetX
                CharMap(CC).Y = TextOffsetY
                CharMap(CC).r = NRC
                GoTo NextChar
            Case 32
                'If TL = CC Then GoTo NextChar
                'If m_bWordWrap And TextOffsetX + CharMap(cc).W > UW Then
                '    GoTo NextChar  'TextOffsetX = LNW Or
                'End If
        End Select

        
        
        If CharMap(CC).p <> -1 Then
            If POWC <> CharMap(CC).p Then
                POWC = CharMap(CC).p

                'does the current word fit?
                If m_bWordWrap And TextOffsetX + WordMap(POWC).W > UW - UWS And POWC > 0 Or (NLNR = True And MultiLine = True And m_bWordWrap = False) Then
MakeNewRule:
                    TextOffsetX = LNW - m_lScrollLeft
                    TTW = TextOffsetX
                    RH = 0
                    RD = 0
                    
                    If m_bWordWrap Then
                        If CC = 0 Then
                            POWC = 0
                        End If
                        
                        If POWC <> -1 Then
                            RH = WordMap(POWC).H
                        End If
                        
                        For RL = POWC To WordCount
                            TTW = TTW + WordMap(RL).W
                            If TTW > UW - UWS And RL >= POWC Then Exit For
                            If WordMap(RL).H > RH Then RH = WordMap(RL).H
                        Next RL
                        
                    Else
                        For RL = CC To UBound(m_byteText)
                            TTW = TTW + CharMap(RL).W
                            
                            If m_byteText(RL) = 10 Then Exit For
                            If CharMap(RL).H - CharMap(RL).d > RH Then RH = CharMap(RL).H - CharMap(RL).d
                            If CharMap(RL).d > RD Then RD = CharMap(RL).d
                            If TTW > MTW Then MTW = TTW
                            'If CharMap(RL).H > RH Then RH = CharMap(RL).H
                            
                        Next RL
                    End If
                    
                    If CC = 0 Then
                        If m_bMultiLine Then
                            TextOffsetY = RH 'TSP + RH
                        Else
                            TextOffsetY = (UH - TSP) / 2 + RH / 2 + 1
                        End If
                        
                        'RowMap(0).StartY = TextOffsetY    '+ RH
                        'GoTo checkNextChar
                    Else
                        TextOffsetY = TextOffsetY + RH
                    End If
                    
                    
                    If m_bMultiLine Or NLNR Or CC = 0 Then
                        'RowMap(NRC).Height = RH
                        
                        If CC <> 0 Then NRC = NRC + 1
                        
                        If NRC > UBound(RowMap) Then growRowMap
                        
                        RowMap(NRC).StartY = TextOffsetY
                        RowMap(NRC).startChar = CC
                        RowMap(NRC).NumChars = 0
                    End If
                    
                    'If m_lScrollTop > NRC Then
                    '    TextOffsetY = 0
                    '    'GoTo NextChar
                    'End If
                    
                    If NLNR = True Or CC = 0 Then
                        NLNR = False
                        GoTo checkNextChar
                    End If
                End If

                'if the word started on the previous row check to break again for really long words!
            ElseIf m_bWordWrap And TextOffsetX + CharMap(CC).W > UW - UWS And RowMap(NRC).NumChars > 0 Then
                GoTo MakeNewRule
            End If
        End If
        RowMap(NRC).Height = RH
        RowMap(NRC).NumChars = RowMap(NRC).NumChars + 1
        
        CharMap(CC).X = TextOffsetX
        CharMap(CC).Y = TextOffsetY
        CharMap(CC).r = NRC
        TextOffsetX = TextOffsetX + CharMap(CC).W
        
NextChar:
    Next CC
    
    m_lScrollTopMax = NRC
    If m_lScrollTop > m_lScrollTopMax Then m_lScrollTop = m_lScrollTopMax
    
    ReDim Preserve RowMap(0 To NRC)
    
    m_bRowMapCalculated = True
End Sub


Sub CalculateUserControlWidthHeight()
    Dim TW As Long    'text width
    Dim LNW As Long    'line number width
    Dim LNR As Long    'line number right
    Dim SC As Long 'start char
     
    If m_sScrollBars <> lNone Then
        UWS = 15
        UHS = 15
        UW = UserControl.ScaleWidth    ' - UWS
        UH = UserControl.ScaleHeight    ' - UHS
      
        If m_sScrollBars = lHorizontal Or m_sScrollBars = lBoth Then
            UH = UH - UHS
        End If
        
        If m_sScrollBars = lVertical Or m_sScrollBars = lBoth Then
            'UW = UW ' - UWS ' - TSP
        Else
            'UW = UW - TSP
        End If
    
    Else
        UW = UserControl.ScaleWidth - TSP
        UH = UserControl.ScaleHeight
    End If
    
    If m_lScrollTop - 1 >= 0 Then
        If m_lScrollTop > UBound(RowMap) Then
            m_lScrollTop = UBound(RowMap)
        End If
        
        SC = RowMap(m_lScrollTop - 1).startChar
        If SC > UBound(CharMap) Then
            SC = UBound(CharMap)
        End If
        
        SYT = CharMap(SC).Y
        TW = TextWidth(m_lScrollTop & "0")
    Else
        SYT = 0
        TW = TextWidth("00")
    End If
    
    
    LNW = 0
    LNR = 0
    If m_bLineNumbers Then    'draw the container for the line numbers
        LNR = TW + TSP
        LNW = LNR + TSP
        LNW = LNW + TSP
    Else
        LNW = TSP
    End If
    
    m_lUsercontrolHeight = UH
    m_lUsercontrolWidth = UW
    m_lUsercontrolTop = TSP
    m_lUsercontrolLeft = LNW + TSP
End Sub

Sub ScrollToEnd()
    Dim RH As Long 'row height
    Dim i As Long
    
    If m_bScrollingTopBar Then Exit Sub
    
    For i = UBound(RowMap) To 0 Step -1
        RH = RH + RowMap(i).Height
        If RH > UH Then
            m_lScrollTop = i + 1
            Exit For
        End If
        
    Next i
    
    If Not m_bStarting Then Redraw
End Sub

Sub Redraw()
    If m_bRefreshing Then
        Exit Sub
    End If
    
    m_bRefreshing = True
    
    
    Dim m_timer As PerformanceTimer
   
    Dim i As Long
    Dim CC As Long    'Char Count
    Dim TL As Long    'text length
    
    Dim TW As Long    'text width
    Dim LNW As Long    'line number width
    Dim LNR As Long    'line number right
    Dim TextOffsetX As Long
    Dim TextOffsetY As Long
    Dim NRC As Long    'Number Row Count


    Dim RH As Long    'row height
    Dim RD As Long    'row d height

    Dim RL As Long    'row loop
    Dim TTW As Long    'temp text width
    Dim MTW As Long   'max text width

    Dim NLNR As Boolean    'Next Loop goto NextRow
    Dim CTP As String 'char to print
    
    'currentStyle values
    Dim cForeColor As Long
    Dim cUnderline As Byte
    Dim cItalic As Byte
    Dim cBold As Byte
    Dim cMarking As Long
    Dim cFontSize As Long
    Dim cStrikeThrough As Byte
    Dim cLine As Long
    
    Dim POWC As Long    'part of word checked
    
    UserControl.Cls

    UserControl.FillStyle = vbFSSolid
    UserControl.DrawStyle = 5
    UserControl.DrawMode = 13
    UserControl.BackColor = m_OleBackgroundColor
    
    SetTextAlign UserControl.hdc, 24  ' 24 = TA_BASELINE
    
    If m_bRowMapCalculated = False Then
        CalculateUserControlWidthHeight
    End If


    TW = TextWidth("00000")
    TSP = 6
    POWC = -1

    LNW = 0
    LNR = 0
    If m_bLineNumbers Then    'draw the container for the line numbers
        LNR = TW + TSP
        LNW = LNR + TSP
        LNW = LNW + TSP
        TextOffsetX = LNW

    Else
        LNW = TSP
        TextOffsetX = TSP
    End If
    
    
    
    If TextOffsetX > UW - UWS Then
        m_bRefreshing = False
        Exit Sub
    End If
    
    If m_bMarkupCalculated = False Then ReCalculateMarkup
    If m_bWordsCalculated = False Then ReCalculateWords m_lRefreshFromCharAt
    If m_bRowMapCalculated = False Then ReCalculateRowMap m_lRefreshFromRowAt
    
    'initialize colors to original color
    UserControl.Font = m_StdFont
    UserControl.FontBold = m_StdFont.Bold
    UserControl.FontStrikethru = m_StdFont.Strikethrough
    UserControl.FontUnderline = m_StdFont.Underline
    UserControl.FontItalic = m_StdFont.Italic
    
    cForeColor = MarkupS(0).lForeColor: UserControl.ForeColor = IIf(cForeColor <> -1, cForeColor, m_OleForeColor)
    cFontSize = MarkupS(0).lFontSize: UserControl.FontSize = IIf(cFontSize <> -1, cFontSize, m_StdFont.Size)
    
    cStrikeThrough = 255
    cBold = 255
    cUnderline = 255
    cItalic = 255
    
    cMarking = -1
    
    Set m_timer = New PerformanceTimer
    
    
    m_timer.StartTimer
    
    NRC = IIf(m_lScrollTopMax > UBound(RowMap), UBound(RowMap), m_lScrollTopMax)
    
    If m_lScrollTop - 1 >= 0 Then
        SYT = CharMap(RowMap(m_lScrollTop - 1).startChar).Y
    Else
        SYT = 0
    End If
    
    TL = UBound(m_byteText)
    

    For i = m_lScrollTop To NRC
        'Debug.Print m_byteText(CC);
        'If i > UBound(RowMap) Then GoTo DoneRefreshing
        
        
        TextOffsetY = CharMap(RowMap(i).startChar).Y - SYT 'RowMap(i).StartY
        
        If m_bRowLines Then
            If TextOffsetY < UH Then
                UserControl.DrawStyle = 0
                UserControl.Line (LNW, TextOffsetY)-(UW - UWS - TSP, TextOffsetY), m_OleRowLineColor
                UserControl.DrawStyle = 5
            End If
        End If
                    
        'If m_bRowLines Then
        '    UserControl.DrawStyle = 0
        '    UserControl.Line (LNW, TextOffsetY)-(UW - UWS - TSP, TextOffsetY), vbRed
        'End If
        
        For CC = RowMap(i).startChar To RowMap(i).startChar + RowMap(i).NumChars - 1
            If CC = TL Then GoTo DoneRefreshing 'do not draw the last character
            
            TextOffsetX = CharMap(CC).X
            
            If cBold <> MarkupS(CC).lBold Then
                cBold = MarkupS(CC).lBold
                If cBold = 255 Then
                    UserControl.FontBold = m_StdFont.Bold
                Else
                    UserControl.FontBold = CBool(cBold)
                End If
            End If
    
            If cUnderline <> MarkupS(CC).lUnderline Then
                cUnderline = MarkupS(CC).lUnderline
                If cUnderline = 255 Then
                    UserControl.FontUnderline = m_StdFont.Underline
                Else
                    UserControl.FontUnderline = CBool(cUnderline)
                End If
            End If
    
            If cItalic <> MarkupS(CC).lItalic Then
                cItalic = MarkupS(CC).lItalic
                If cItalic = 255 Then
                    UserControl.FontItalic = m_StdFont.Italic
                Else
                    UserControl.FontItalic = CBool(cItalic)
                End If
            End If
    
            If cFontSize <> MarkupS(CC).lFontSize Then
                cFontSize = MarkupS(CC).lFontSize
                If cFontSize = -1 Then
                    UserControl.FontSize = m_StdFont.Size
                Else
                    UserControl.FontSize = cFontSize
                End If
            End If
    
            If cStrikeThrough <> MarkupS(CC).lStrikeThrough Then
                cStrikeThrough = MarkupS(CC).lStrikeThrough
                If cStrikeThrough = 255 Then
                    UserControl.FontStrikethru = m_StdFont.Strikethrough
                Else
                    UserControl.FontStrikethru = CBool(cStrikeThrough)
                End If
                
            End If
    
            
            
            If TextOffsetY - RowMap(i).Height < UH And TextOffsetX - TSP < UW And TextOffsetX + CharMap(CC).W > 0 And TextOffsetY >= 0 Then  '
                Dim jj As Long
                Dim kk As Long
    
                If cMarking <> MarkupS(CC).lMarking Then
                    cMarking = MarkupS(CC).lMarking
                    If cMarking <> -1 Then
                        UserControl.FillColor = MarkupS(CC).lMarking
                    End If
                End If
    
    
                If cMarking <> -1 Then
                    pts(0).X = TextOffsetX
                    pts(0).Y = TextOffsetY + CharMap(CC).d
    
                    pts(1).X = TextOffsetX + CharMap(CC).W
                    pts(1).Y = pts(0).Y
    
                    pts(2).X = pts(1).X
                    pts(2).Y = pts(0).Y - CharMap(CC).H 'TextOffsetY - CharMap(CC).H + CharMap(CC).d
    
                    pts(3).X = pts(0).X
                    pts(3).Y = pts(2).Y
                    'UserControl.DrawMode = 15
    
                    Polygon UserControl.hdc, pts(0), 4
                    'UserControl.DrawMode = 13
                End If
    
                cLine = MarkupS(CC).lLine
    
    
                If cLine <> -1 Then
                    If cLine <> cForeColor Then
                        cForeColor = cLine
                        UserControl.ForeColor = cLine
                    End If
    
                    
                    CTP = ChrW(m_byteText(CC))
                    For jj = -1 To 1
                        For kk = -1 To 1
                            If Not (jj = 0 And kk = 0) Then
                                ''UserControl.CurrentX = TextOffsetX + jj ' + 1
                                ''UserControl.CurrentY = TextOffsetY + kk '- 1    '- CharMap(CC).H
                                ''UserControl.Print MS;
                                
                                TextOut UserControl.hdc, TextOffsetX + jj, TextOffsetY + kk, CTP, 1
                            End If
                        Next kk
                    Next jj
                    ''UserControl.CurrentX = TextOffsetX + 1
                    ''UserControl.CurrentY = TextOffsetY - 1    ' - CharMap(CC).H
                Else
                    ''UserControl.CurrentX = TextOffsetX
                    ''UserControl.CurrentY = TextOffsetY    ' - CharMap(CC).H
                End If
    
                If cForeColor <> MarkupS(CC).lForeColor Then
                    cForeColor = MarkupS(CC).lForeColor
                    If cForeColor = -1 Then
                        UserControl.ForeColor = m_OleForeColor
                    Else
                        UserControl.ForeColor = cForeColor
                    End If
                End If
                
                
                If m_byteText(CC) <> 10 And m_byteText(CC) <> 13 Then TextOut UserControl.hdc, TextOffsetX, TextOffsetY, ChrW(m_byteText(CC)), 1

                'UserControl.Print Chr(m_byteText(cc));
    
                If CC >= m_SelStart And CC < m_SelEnd And m_byteText(CC) <> 10 Then
    
                    pts(0).X = TextOffsetX
                    pts(0).Y = TextOffsetY + CharMap(CC).d
    
                    pts(1).X = TextOffsetX + CharMap(CC).W
                    pts(1).Y = pts(0).Y
    
                    pts(2).X = pts(1).X
                    pts(2).Y = TextOffsetY - RowMap(i).Height + IIf(m_bMultiLine, CharMap(CC).d, 0)
    
                    pts(3).X = TextOffsetX
                    pts(3).Y = pts(2).Y
                    
                    UserControl.DrawMode = 6 '6
                    Polygon UserControl.hdc, pts(0), 4
                    UserControl.DrawMode = 13
                End If
    
            
            ElseIf TextOffsetY - RowMap(i).Height >= UH Then
                GoTo DoneRefreshing
            End If
            
            'CharMap(cc).x = TextOffsetX
            'CharMap(cc).y = TextOffsetY
            
            'RowMap(NRC).NumChars = RowMap(NRC).NumChars + 1
    
            'TextOffsetX = TextOffsetX + CharMap(cc).W
    
NextChar:
        Next CC
    Next i
DoneRefreshing:
    
    'm_timer.StopTimer
    
    'Debug.Print m_timer.TimeElapsed(pvMilliSecond)
    
    m_lRefreshFromRowAt = -1
    m_lRefreshFromCharAt = -1
    
    
    'Debug.Print Round(m_timer.tStop, 5)
    
    m_lMouseDownPrevious = m_lMouseDown
    'Debug.Print Round(m_timer.tStop, 6)
    
    UserControl.Font = m_StdFont
    UserControl.ForeColor = m_OleForeColor
    UserControl.BackColor = m_OleBackgroundColor
    UserControl.FontSize = m_StdFont.Size
    UserControl.FontStrikethru = m_StdFont.Strikethrough
    UserControl.FontUnderline = m_StdFont.Underline
    UserControl.FontItalic = m_StdFont.Italic
    UserControl.FontBold = m_StdFont.Bold
    UserControl.DrawStyle = 0
    UserControl.DrawMode = 13
    UserControl.ForeColor = m_OleLineNumberForeColor
    UserControl.FillColor = m_OleLineNumberBackground
    
    'ReDim Preserve RowMap(0 To NRC)
    
    If m_bLineNumbers Then
        pts(0).X = 0:         pts(0).Y = 0
        pts(1).X = LNR + TSP: pts(1).Y = 0
        pts(2).X = pts(1).X:  pts(2).Y = UH
        pts(3).X = 0:         pts(3).Y = UH
        Polygon UserControl.hdc, pts(0), 4
        
        For i = m_lScrollTop To NRC
            TW = UserControl.TextWidth(i + 1)
            UserControl.CurrentX = LNR - TW
            If RowMap(i).StartY - RowMap(i).Height - SYT < UH Then
                UserControl.CurrentY = RowMap(i).StartY - SYT    ' - TH
                UserControl.Print CStr(i + 1)
                'UserControl.Line (TSP, RowMap(i).StartY - SYT)-(LNR, RowMap(i).StartY - SYT), m_OleRowLineColor
            Else
                Exit For
            End If
        Next i
    End If
    
    UserControl.FillColor = m_OleBackgroundColor
    
    m_lScrollLeftMax = m_lScrollLeft + (MTW - UW)
    If m_lScrollLeftMax > 0 And m_lScrollLeft > m_lScrollLeftMax Then m_lScrollLeft = m_lScrollLeftMax

    If m_sScrollBars <> lNone Then
        DrawScrollBars
    End If
    
    If m_bAutoResize Then
        'If m_lScrollLeftMax <> 0 Then
            UserControl.Width = ScaleX(MTW + TSP, vbPixels, vbTwips)
        'End If
    End If
    
    
    If m_bBorder Then
        UserControl.DrawWidth = m_lBorderThickness
        UserControl.Line (0, 0)-(0, UserControl.ScaleHeight), m_OleBorderColor
        UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), m_OleBorderColor
        UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), m_OleBorderColor
        UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), m_OleBorderColor

        UserControl.Line (UserControl.ScaleWidth, 0)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), m_OleBorderColor
    End If

    UserControl.Refresh
    
    updateCaretPos
    
    'DoEvents
    m_bRefreshing = False
    'If m_bRefreshedWhileBusy Then
    '    m_bRefreshedWhileBusy = False
    '    Redraw
    'End If
End Sub


Sub ReCalculateWords(Optional fromWhere As Long = 0)
    Dim WC As Long    'word count
    Dim WH As Long    'word height
    Dim WW As Long    'word width
    Dim WL As Long    'word length
    Dim BT As Long    'ByteText
    Dim UB As Long    'ubound bytetext
    Dim TL As Long    'text length
    Dim POW As Long   'part of word
    
    If m_bMarkupCalculating Then Exit Sub
    m_bMarkupCalculating = True

    On Error GoTo endff
    
    'ReDim Preserve WordMap(0 To UBound(m_byteText) + 2)
    'ReDim RowMap(0 To 0)

    UB = UBound(m_byteText)
    
    If fromWhere <= 1 Then
        ReDim WordMap(0)
    Else
        POW = CharMap(fromWhere).p
        If POW <> -1 Then
            fromWhere = WordMap(POW).s
            WC = POW
        End If
    End If
    
    If fromWhere = -1 Then fromWhere = 0

    For TL = fromWhere To UB
        BT = m_byteText(TL)
        
        If TL < UB And (BT = 32 Or BT = 10 Or (BT >= 40 And BT <= 47) Or BT = 58 Or BT = 59) Then      ' a space  Or m_byteText(TL) = 13
            If WL >= 0 Then
                
                WordMap(WC).H = WH
                WordMap(WC).W = WW
                WordMap(WC).l = WL
                'If m_byteText(TL) <> 10 Then
                WC = WC + 1
                
                If WC > UBound(WordMap) Then
                    growWordMap
                End If
                
                'End If
                WH = 0
                WW = 0
                WL = 0

                WordMap(WC).s = TL + 1
                CharMap(TL).p = -1
            End If
        Else
            CharMap(TL).p = WC
            If CharMap(TL).H > WH Then
                WH = CharMap(TL).H
            End If
            WW = WW + CharMap(TL).W
            WL = WL + 1

        End If
    Next TL


    WordMap(WC).H = WH
    WordMap(WC).W = WW
    WordMap(WC).l = WL

    WordCount = WC
endff:

    m_bMarkupCalculating = False
    m_bWordsCalculated = True
End Sub



Function InstrByte(lStart As Long, ByRef lBytes() As Byte, lSearch As Byte) As Long
    Dim i As Long

    For i = lStart To UBound(lBytes)
        If lBytes(i) = lSearch Then
            InstrByte = i
            Exit Function
        End If
    Next i
End Function

Function RGBByte(lStart As Long, ByRef lBytes() As Byte) As Long
    Dim i As Long
    Dim c(0 To 8) As Byte

    For i = 0 To 5
        Select Case lBytes(lStart + i)
            Case 48 To 57
                c(i) = (lBytes(lStart + i) - 48) And 255

            Case 65 To 70
                c(i) = (lBytes(lStart + i) - 55) And 255

            Case 97 To 102
                c(i) = (lBytes(lStart + i) - 85) And 255
        End Select
    Next i

    RGBByte = RGB(c(0) * 16 + c(1), c(2) * 16 + c(3), c(4) * 16 + c(5))
End Function


Public Sub Clear()
    m_SelStart = 0
    m_SelEnd = UBound(m_byteText)
    m_CursorPos = m_SelEnd
    
    AddCharAtCursor , True
    
    If Not m_bStarting Then Redraw
    
    ClearMarkup
End Sub


Private Function parseConsoleColors(ByRef bytes() As Byte, ByRef styleArr() As MarkupStyles, ByRef byteText() As Byte) As String

    'Dim strSplit() As String
    Dim i As Long 'primary index
    Dim j As Long 'to check what the rest of the command contains
    
    Dim lCommand As Long
    Dim posCommand As Long
    Dim isCommand As Boolean
    Dim UB As Long 'upper bound
    Dim CL As Long 'command length
    Dim CN As Long 'command number
    Dim CC As Boolean 'complete command

    UB = UBound(bytes)
    'strSplit = Split(str, Chr(&H1B))
    
    Dim ATL As Long 'actual text length
    
    For i = 0 To UB
        
check_for_next_color:
        
        If bytes(i) = 27 Then
            If i < UB Then
                If bytes(i + 1) = 91 Then
                    If i + 1 < UB Then
                        CN = 0
                        CL = 0
                        CC = False
                        
                        For j = i + 2 To UB 'just run to the end of the command
                            CL = CL + 1
                            Select Case bytes(j)
                                Case 109 'm' end of command
                                    CC = True
                                    Exit For
                                    
                                Case 48 To 57 '0' to '9'
                                    CN = CN * 10 + (bytes(j) - 48)
                                    
                            End Select
                            
                            If CL > 3 Then
                                GoTo process_as_normal_char
                            End If
                        Next j
                        
                        If CC Then
                            'If ATL > 0 Then styleArr(ATL) = styleArr(ATL - 1)
                            
                            Select Case CN
                                Case 30 To 37
                                    styleArr(ATL).lForeColor = m_OleConsoleColors(CN - 30)
                                Case 40 To 47
                                    styleArr(ATL).lMarking = m_OleConsoleColors(CN - 40)
                            End Select
                            i = i + CL + 2
                            If i <= UB Then
                                GoTo check_for_next_color
                            Else
                                GoTo end_of_parsing
                            End If
                            
                        Else 'this happens if the command was not finnished before the ubound of bytes
                            For j = i To UB
                                parseConsoleColors = parseConsoleColors & ChrW(bytes(j))
                            Next j
                            GoTo end_of_parsing
                        End If
                        
                    Else
                        parseConsoleColors = ChrW(bytes(i)) & ChrW(bytes(i + 1))
                        GoTo end_of_parsing
                    End If
                Else
                    GoTo process_as_normal_char
                End If
                
            Else
                parseConsoleColors = ChrW(bytes(i))
                GoTo end_of_parsing
            End If
            
        Else
            
process_as_normal_char:
            
            byteText(ATL) = bytes(i)
            ATL = ATL + 1
            styleArr(ATL) = styleArr(ATL - 1)
        End If
        
        
        'posCommand = InStr(1, strSplit(i), "m")
        'isCommand = (Left$(strSplit(i), 1) = "[")
        
'        If i = UBound(strSplit) Then
'            If Len(strSplit(i)) = 0 And (posCommand = 0 Or isCommand = False) Then
'                parseAndAddText = Chr$(&H1B) & strSplit(i)
'                Exit Function
'            End If
'        End If
'
'        If posCommand > 0 And isCommand Then
'            lCommand = Val(Mid(strSplit(i), 2, posCommand))
'
'            Dim TL As Long
'            TL = txtReceived.TextLength
'            If TL < 0 Then TL = 0
'
'            Select Case lCommand
'                Case 30 To 37
'                    txtReceived.setCharForeColor TL, CLng(ConsoleColors(lCommand - 30))
'
'                Case 40 To 47
'                    txtReceived.setCharBackColor TL, CLng(ConsoleColors(lCommand - 40))
'
'                Case 0
'                    txtReceived.setCharForeColor TL, -1
'            End Select
'
'            Dim strToAdd As String
'
'            strToAdd = Right(strSplit(i), Len(strSplit(i)) - posCommand)
'
'            txtReceived.AddCharAtCursor strToAdd
'        Else
'            If Len(strSplit(i)) > 0 Then
'                txtReceived.AddCharAtCursor strSplit(i), True
'            End If
'        End If
        
        'If Len(strToAdd) <> 5 Then
        '    Debug.Print strToAdd
        'End If
        
        'Debug.Assert Len(strToAdd) = 5
        
    Next i
    
    
    parseConsoleColors = ""
    
end_of_parsing:
    If ATL > 0 Then
        ReDim Preserve byteText(0 To ATL - 1)
        ReDim Preserve styleArr(0 To ATL)
    Else
        ReDim byteText(0)
        Erase styleArr 'way to check if there are no chars to add, only a buffer
    End If
    
    
End Function

Function AddCharAtCursor(Optional ByRef sChar As String = "", Optional noevents As Boolean = False) As Boolean
    Dim lLength As Long
    Dim i As Long

    Dim lInsertLength As Long
    Dim lLengthDifference As Long
    Dim reCalculateFromWhere As Long
    Dim distanceFromEnd As Long
    Dim CursorToEnd As Long
    
    Dim byteText() As Byte
    Dim newByteText() As Byte
    Dim newMarkupStyles() As MarkupStyles
    

    'performance.StartTimer
    
    If m_bConsoleColors Then
        byteText = StrConv(m_sConsoleColorBuffer & sChar, vbFromUnicode)
        
        ReDim newMarkupStyles(0 To UBound(byteText) + 1)
        ReDim newByteText(0 To UBound(byteText) + 1)
        
        If m_SelStart = 0 And m_SelEnd = 0 Then
            newMarkupStyles(0) = MarkupS(UBound(MarkupS))
        ElseIf m_SelStart > 0 Then
            newMarkupStyles(0) = MarkupS(m_SelStart)
        Else
            newMarkupStyles(0) = MarkupS(m_SelEnd)
        End If
        
        'Debug.Print newMarkupStyles(0).lFontSize
        
        m_sConsoleColorBuffer = parseConsoleColors(byteText, newMarkupStyles, newByteText)
        
        lInsertLength = UBound(newByteText) + 1
        
        If (Not (Not newMarkupStyles)) = 0 Then
            lInsertLength = 0
        End If
        If lInsertLength = 0 And m_SelStart = m_SelEnd Then Exit Function
        
    Else
        newByteText = StrConv(sChar, vbFromUnicode)
        
        lInsertLength = Len(sChar)
        If lInsertLength = 0 And m_SelStart = m_SelEnd Then Exit Function
    End If
    
    
    

    If m_SelStart <> m_SelEnd Then
        lLengthDifference = lInsertLength - (m_SelEnd - m_SelStart)
    Else
        lLengthDifference = lInsertLength
    End If
    
    reCalculateFromWhere = IIf(m_SelStart < m_SelEnd, m_SelStart, m_SelEnd)
    distanceFromEnd = UBound(m_byteText) - reCalculateFromWhere
    
    CursorToEnd = UBound(CharMap) - m_SelEnd + 1
    
    If lLengthDifference > 0 Then
        ReDim Preserve CharMap(0 To UBound(CharMap) + lLengthDifference)
        ReDim Preserve m_byteText(0 To UBound(m_byteText) + lLengthDifference)
        ReDim Preserve MarkupS(0 To UBound(MarkupS) + lLengthDifference)
        
        CopyMemory CharMap(m_SelEnd + lLengthDifference), CharMap(m_SelEnd), CursorToEnd * LenB(CharMap(0))
        CopyMemory m_byteText(m_SelEnd + lLengthDifference), m_byteText(m_SelEnd), CursorToEnd
        CopyMemory MarkupS(m_SelEnd + lLengthDifference), MarkupS(m_SelEnd), CursorToEnd * LenB(MarkupS(0))

        'Debug.Print m_byteText(UBound(m_byteText))

    ElseIf lLengthDifference < 0 Then

        CopyMemory CharMap(m_SelEnd + lLengthDifference), CharMap(m_SelEnd), CursorToEnd * LenB(CharMap(0))
        CopyMemory m_byteText(m_SelEnd + lLengthDifference), m_byteText(m_SelEnd), CursorToEnd
        CopyMemory MarkupS(m_SelEnd + lLengthDifference), MarkupS(m_SelEnd), CursorToEnd * LenB(MarkupS(0))

        ReDim Preserve CharMap(0 To UBound(CharMap) + lLengthDifference)
        ReDim Preserve m_byteText(0 To UBound(m_byteText) + lLengthDifference)
        ReDim Preserve MarkupS(0 To UBound(MarkupS) + lLengthDifference)
    End If
    
    If lInsertLength > 0 Then
        CopyMemory m_byteText(m_SelStart), newByteText(0), lInsertLength * LenB(m_byteText(0))
        If m_bConsoleColors Then
            'Debug.Print MarkupS(m_SelStart).lForeColor
            CopyMemory MarkupS(m_SelStart), newMarkupStyles(0), (lInsertLength + 1) * LenB(MarkupS(0))
        Else
            If distanceFromEnd = 0 Then
                For i = 0 To lInsertLength - 1
                    'm_byteText(m_SelStart + i) = newByteText(i) 'Asc(Mid$(sChar, i, 1))
                    MarkupS((m_SelStart + i)) = MarkupS(UBound(MarkupS))
                Next i
            Else
                For i = 0 To lInsertLength - 1
                    'm_byteText(m_SelStart + i) = newByteText(i) 'Asc(Mid$(sChar, i, 1))
                    
                    If m_SelStart + i - 1 >= 0 Then
                        MarkupS((m_SelStart + i)) = MarkupS((m_SelStart + i - 1))
                    ElseIf m_SelStart = 0 And m_SelEnd = 0 Then
                        MarkupS(0) = MarkupS(UBound(MarkupS))
                    Else
                        With MarkupS(m_SelStart + i)
                            .lStrikeThrough = m_StdFont.Strikethrough
                            .lFontSize = -1
                            .lUnderline = 255
                            .lItalic = 255
                            .lBold = 255
                            .lMarking = -1
                            .lForeColor = -1
                            .lLine = -1
                        End With
                    End If
            
                 Next i
            End If
        End If
    End If
    
    
    

    
    CheckCharSize m_SelStart, lInsertLength

    m_SelStart = m_SelStart + lInsertLength
    m_SelEnd = m_SelStart
    m_CursorPos = m_SelStart

    'm_byteText(m_CursorPos) = Asc(sChar)
    'm_bMarkupCalculated = False
    
    If m_lRefreshFromCharAt <> -1 Then
        If m_lRefreshFromCharAt > reCalculateFromWhere Then
            m_lRefreshFromCharAt = reCalculateFromWhere
            m_lRefreshFromRowAt = CharMap(m_lRefreshFromCharAt).r
        End If
        
    Else
        m_lRefreshFromCharAt = reCalculateFromWhere
        If m_lRefreshFromCharAt > 0 Then
            m_lRefreshFromRowAt = CharMap(m_lRefreshFromCharAt - 1).r
        Else
            m_lRefreshFromRowAt = 0
        End If

    End If
    
    m_bWordsCalculated = False
    m_bRowMapCalculated = False
    'CalculateUserControlWidthHeight
    'UserControl_KeyDown vbKeyRight, 0
    
    AddCharAtCursor = True
    If Not noevents Then RaiseEvent Changed
    
    'performance.StopTimer
    
    'Debug.Print performance.TimeElapsed(pvMilliSecond)
End Function


Sub CheckCharSize(lStart As Long, lLength As Long)
    Dim i As Long
    Dim uSize As Long
    
    Dim cForeColor As Long
    Dim cUnderline As Byte
    Dim cItalic As Byte
    Dim cBold As Byte
    Dim cMarking As Long
    Dim cFontSize As Long
    Dim cStrikeThrough As Byte
    Dim cLine As Long
    Dim cDescendHeight As Long
    Dim cTextMetric As TEXTMETRIC
    
    UserControl.Font = m_StdFont
    UserControl.ForeColor = m_OleForeColor
    UserControl.BackColor = m_OleBackgroundColor
    UserControl.FontSize = m_StdFont.Size
    UserControl.FontStrikethru = m_StdFont.Strikethrough
    UserControl.FontUnderline = m_StdFont.Underline
    UserControl.FontItalic = m_StdFont.Italic
    UserControl.FontBold = m_StdFont.Bold

    cForeColor = -1
    cUnderline = 255
    cItalic = 255
    cBold = 255
    cFontSize = -1
    cMarking = -1
    cLine = -1
    cStrikeThrough = 255

    GetTextMetrics UserControl.hdc, cTextMetric
    cDescendHeight = cTextMetric.tmDescent
    
    'uSize = UBound(MarkupS)
    
    For i = lStart To lStart + lLength
        With MarkupS(i)
            If .lFontSize <> cFontSize Then
                cFontSize = .lFontSize
                If .lFontSize = -1 Then
                    UserControl.Font.Size = m_StdFont.Size
                Else
                    UserControl.Font.Size = cFontSize
                End If

                GetTextMetrics UserControl.hdc, cTextMetric
                cDescendHeight = cTextMetric.tmDescent
            End If

            If .lBold <> cBold Then
                cBold = .lBold
                If cBold = 255 Then
                    UserControl.Font.Bold = m_StdFont.Bold
                Else
                    UserControl.Font.Bold = CBool(cBold)
                End If
            End If

            If .lItalic <> cItalic Then
                cItalic = .lItalic
                If cItalic = 255 Then
                    UserControl.Font.Italic = m_StdFont.Italic
                Else
                    UserControl.Font.Italic = cItalic
                End If
            End If

            If .lUnderline <> cUnderline Then
                cUnderline = .lUnderline
                If .lUnderline = 255 Then
                    UserControl.Font.Underline = m_StdFont.Underline
                Else
                    UserControl.Font.Underline = cUnderline
                End If
            End If

            If .lStrikeThrough <> cStrikeThrough Then
                cStrikeThrough = .lStrikeThrough
                If .lStrikeThrough = 255 Then
                    UserControl.Font.Strikethrough = m_StdFont.Strikethrough
                Else
                    UserControl.Font.Strikethrough = cStrikeThrough
                End If
            End If
            

            If Not (m_byteText(i) = 13 Or m_byteText(i) = 10) Then
                GetTextSize Chr(m_byteText(i)), CharMap(i)
            Else
                GetTextSize " ", CharMap(i)
            End If

            CharMap(i).d = cDescendHeight

        End With

    Next i



End Sub

Function getCharAtCursor(X As Long, Y As Long) As Long
Dim i As Long

    Dim TTW     As Long    'total text width
    Dim CTW  As Long    'current text width
    Dim CR      As Long    'char row
    Dim UB      As Long    'ubound of rowmap
    Dim EOR As Long 'end of row
    Dim TS As Long 'total size
    
    UB = UBound(RowMap)
    CR = UB
    
    For i = m_lScrollTop To UB    'number of rows
        If Y < RowMap(i).StartY - SYT Then 'And RowMap(i).NumChars <> 1
            CR = i
            Exit For
        End If
    Next i

    If CR = UB Then
        EOR = UBound(CharMap)
    Else
        EOR = RowMap(CR + 1).startChar - 1
    End If
    
    'If CharMap(RowMap(CR).startChar).x > x Then
    '    getCharAtCursor = RowMap(CR).startChar
    '    Exit Function
    If CharMap(RowMap(CR).startChar + RowMap(CR).NumChars - 1).X < X Then
        getCharAtCursor = EOR
        Exit Function
    End If
    
    'TS = CharMap(RowMap(CR).StartChar).X
    For i = RowMap(CR).startChar To EOR
        If m_byteText(i) <> 10 And m_byteText(i) <> 13 Then
            If X > CharMap(i).X And X <= CharMap(i).X + CharMap(i).W Then
                If X < CharMap(i).X + CharMap(i).W / 2 Then
                    getCharAtCursor = i
                Else
                    getCharAtCursor = i + 1
                    If getCharAtCursor > UBound(m_byteText) Then getCharAtCursor = UBound(m_byteText)
                End If
                Exit Function
            ElseIf CharMap(i).X > X Then
                getCharAtCursor = i
                Exit Function
            End If
        End If

    Next i
End Function


Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    m_bHasFocus = False
    DestroyCaret
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpSwapSel As Long
    Dim mustRedraw As Boolean
    
    m_lMouseDown = m_lMouseDown Or Button
    m_lMouseDownX = X ' - m_lScrollLeft
    m_lMouseDownY = Y ' - m_lScrollTop
    m_lMouseX = X ' - m_lScrollLeft
    m_lMouseY = Y ' - m_lScrollTop
    
    If X > UserControl.ScaleWidth - 15 And m_sScrollBars <> lNone Then
        If m_lMouseDownY >= m_lScrollTopBarY And m_lMouseDownY <= m_lScrollTopBarY + m_lScrollTopBarHeight Then
            m_bScrollingTopBar = True
            m_lScrollTopDragStartY = m_lMouseDownY - m_lScrollTopBarY
            m_lScrollTopDragStartValue = m_lScrollTop
        End If
        DrawScrollBars
    Else
    
        tmpSwapSel = m_CursorPos
        m_CursorPos = getCharAtCursor(CLng(m_lMouseDownX), CLng(m_lMouseDownY))
        m_lMouseDownPos = m_CursorPos
        
        'Debug.Print m_byteText(m_CursorPos); m_CursorPos
        
        getSelectionChanged True
        
        If (Shift And 1) Then
            If m_CursorPos <= tmpSwapSel Then
                m_SelStart = m_CursorPos
                m_SelEnd = tmpSwapSel
            Else
                m_SelStart = tmpSwapSel
                m_SelEnd = m_CursorPos
            End If
        Else
            m_SelStart = m_CursorPos
            m_SelEnd = m_CursorPos
        End If
        
        If getSelectionChanged Then
            RaiseEvent SelectionChanged
            mustRedraw = True
        End If
        
    End If
    
    updateCaretPos
    RaiseEvent Click(m_SelStart, CharMap(m_SelStart).r)
    
    
    If Not m_bStarting And mustRedraw Then Redraw
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpSwapSel As Long
    Dim mustRedraw As Boolean
    
    Dim scrollTopDifference As Long
    Dim scrollTopNewPosition As Long
            
    m_lMouseX = X ' - m_lScrollLeft
    m_lMouseY = Y ' - m_lScrollTop
    
    If X > UW - UWS And ((Y >= m_lScrollTopBarY And Y <= m_lScrollTopBarY + m_lScrollTopBarHeight) Or Y < UWS Or Y > UH - UWS) Then
        UserControl.MousePointer = 0
    Else
        UserControl.MousePointer = 3
    End If
    
    If m_lMouseDown <> lNone Then
        If m_bScrollingTopBar = True Then
            
            scrollTopDifference = (m_lMouseY - m_lScrollTopDragStartY - 15)
            If m_lScrollTopMax = 0 Then
                scrollTopNewPosition = 0
            Else
                scrollTopNewPosition = m_lScrollTopMax / (m_lScrollTopHeight - m_lScrollTopBarHeight) * scrollTopDifference ' / (m_lScrollTopMax / 30)
            End If
            
            If scrollTopDifference <> 0 And m_lScrollTop <> scrollTopDifference Then
                m_lScrollTop = scrollTopNewPosition
                If m_lScrollTop > m_lScrollTopMax Then m_lScrollTop = m_lScrollTopMax
                If m_lScrollTop < 0 Then m_lScrollTop = 0
                mustRedraw = True
                updateCaretPos
            Else
                DrawScrollBars
            End If
            
            
        ElseIf m_lMouseDownX < UW - UWS Then
            getSelectionChanged True
            
            m_CursorPos = getCharAtCursor(CLng(m_lMouseX), CLng(m_lMouseY))
            tmpSwapSel = m_lMouseDownPos 'getCharAtCursor(CLng(m_lMouseDownX), CLng(m_lMouseDownY))
            
            If m_CursorPos <= tmpSwapSel Then
                m_SelStart = m_CursorPos
                m_SelEnd = tmpSwapSel
            Else
                m_SelStart = tmpSwapSel
                m_SelEnd = m_CursorPos
            End If
            updateCaretPos
            
            If getSelectionChanged Then
                mustRedraw = True
                RaiseEvent SelectionChanged
            End If
            
            If CharMap(m_CursorPos).X >= m_lUsercontrolWidth Then
                m_lScrollLeft = m_lScrollLeft + (CharMap(m_CursorPos).X - m_lUsercontrolWidth)
                mustRedraw = True
            ElseIf CharMap(m_CursorPos).X <= m_lUsercontrolLeft Then
                m_lScrollLeft = m_lScrollLeft + (CharMap(m_CursorPos).X - m_lUsercontrolLeft)
                mustRedraw = True
            End If
            
        End If
        
        If Not m_bStarting And mustRedraw Then Redraw
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_lMouseDown = m_lMouseDown And Not Button
    m_bScrollingTopBar = False
    'Debug.Print m_lMouseDown

    If X > UserControl.ScaleWidth - 15 And m_sScrollBars <> lNone Then
        If Y < UWS Then
            m_lScrollTop = m_lScrollTop - 1
            If m_lScrollTop < 0 Then m_lScrollTop = 0
        ElseIf Y > UH - UWS Then
            m_lScrollTop = m_lScrollTop + 1
            If m_lScrollTop > UBound(RowMap) Then m_lScrollTop = UBound(RowMap)
        End If
        DrawScrollBars
    End If
        
        
    If Not m_bStarting Then Redraw
End Sub


Function getNextCharUpDown(U As Boolean, STS As Boolean) As Long 'up, selectionTheSame
    Dim i As Long

    Dim TTW     As Long    'total text width
    Static CTW  As Long    'current text width
    Dim CR      As Long    'current word
    Dim UB      As Long    'ubound of rowmap
    Dim TL As Long 'text length
    
    UB = UBound(RowMap)
    TL = UBound(CharMap)
    CR = UB
    
    For i = 0 To UB    'number of rows
        If m_CursorPos < RowMap(i).startChar Then
            CR = i - 1
            Exit For
        End If
    Next i
    
    
    If Not STS Then
        CTW = 0
        For i = RowMap(CR).startChar To m_CursorPos - 1
            CTW = CTW + CharMap(i).W
        Next i
    End If

    If U Then
        CR = CR - 1
        If CR < 0 Then
            CR = 0
        End If

    Else
        CR = CR + 1
        If CR > UB Then
            CR = UB
        End If
    End If
'
    
    For i = RowMap(CR).startChar To RowMap(CR).NumChars + RowMap(CR).startChar
        If i > TL Then
            getNextCharUpDown = i
            Exit Function
        End If
        
        TTW = TTW + CharMap(i).W
        If (TTW > CTW Or i = RowMap(CR).NumChars + RowMap(CR).startChar) And m_byteText(i) <> 13 And m_byteText(i) <> 10 Then
            getNextCharUpDown = i
            Exit Function
        End If
    Next i

    getNextCharUpDown = RowMap(CR).startChar

End Function


Function getSelectionChanged(Optional init As Boolean = False) As Boolean
    Static tmpSelStart As Long
    Static tmpSelEnd As Long
    Static tmpCursorPos As Long
    
    If init Then
        tmpSelStart = m_SelStart
        tmpSelEnd = m_SelEnd
        tmpCursorPos = m_CursorPos
        getSelectionChanged = False
    Else
        getSelectionChanged = tmpSelStart <> m_SelStart Or tmpSelEnd <> m_SelEnd Or tmpCursorPos <> m_CursorPos
    End If
End Function

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If m_bBlockNextKeyPress Then
        m_bBlockNextKeyPress = False
        Exit Sub
    End If
    
    If (KeyAscii >= 32 And KeyAscii <= 126) Or (KeyAscii >= 128 And KeyAscii <= 255) Then
        AddCharAtCursor Chr(KeyAscii)
        If Not m_bStarting Then Redraw
        
        updateCaretPos
        
    End If
End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Long
    Dim tmpswap As Long
    Dim mustRedraw As Boolean
    Dim tmpCursor As Long
    Dim tmpString As String
    Dim tmpCursorUpDown As Boolean
    Dim tmpStarting As Boolean
    
    getSelectionChanged True
    
    tmpStarting = m_bStarting
    m_bStarting = True
    RaiseEvent KeyDown(KeyCode, Shift)
    m_bStarting = tmpStarting
    
    
    If (KeyCode = 0 And Shift = 0) Or Locked Then
        m_bBlockNextKeyPress = True
        mustRedraw = True
    End If
    
    
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            If m_lMouseDown <> lNone Then Exit Sub
            m_CursorPos = getNextCharUpDown(KeyCode = vbKeyUp, m_SelUpDownTheSame)
            m_SelStart = m_CursorPos
            m_SelEnd = m_CursorPos
            
            If RowMap(m_lScrollTop).startChar > m_SelStart Then
                m_lScrollTop = m_lScrollTop - 1
                If m_lScrollTop < 0 Then m_lScrollTop = 0
                mustRedraw = True
            ElseIf CharMap(m_SelStart).Y - SYT > UH Then
                m_lScrollTop = m_lScrollTop + 1
                If m_lScrollTop > m_lScrollTopMax Then m_lScrollTop = m_lScrollTopMax
                mustRedraw = True
            End If
            
            tmpCursorUpDown = True
            m_SelUpDownTheSame = True
        
        Case vbKeyRight
            If m_lMouseDown <> lNone Then Exit Sub
            
            If (Shift And 2) Then
                tmpCursor = getNextWordFromCursor()
                If (Shift And 1) Then
                    If m_SelStart = m_CursorPos Then
                        If m_SelEnd > m_SelStart Then
                            m_SelStart = tmpCursor
                        Else
                            m_SelEnd = tmpCursor
                        End If
                    Else
                        m_SelEnd = tmpCursor
                    End If
                    mustRedraw = True
                Else
                    If m_SelStart <> m_SelEnd Then mustRedraw = True
                    m_SelStart = tmpCursor
                    m_SelEnd = tmpCursor
                End If
                m_CursorPos = tmpCursor

            ElseIf Shift = 0 Then
                m_CursorPos = getNextChar(m_CursorPos + 1)
                If m_SelStart <> m_SelEnd Then mustRedraw = True
                m_SelStart = m_CursorPos
                m_SelEnd = m_CursorPos


            ElseIf (Shift And 1) Then
                If m_SelStart = m_CursorPos Then
                    If m_SelEnd > m_SelStart Then
                        m_SelStart = getNextChar(m_SelStart + 1)
                        m_CursorPos = m_SelStart
                    Else
                        m_SelEnd = getNextChar(m_SelEnd + 1)
                        m_CursorPos = m_SelEnd
                    End If
                Else
                    m_SelEnd = getNextChar(m_SelEnd + 1)
                    m_CursorPos = m_SelEnd
                End If
                mustRedraw = True
            End If

        Case vbKeyLeft
            If m_lMouseDown <> lNone Then Exit Sub
            
            If (Shift And 2) Then
                tmpCursor = getPreviousWordFromCursor()
                If (Shift And 1) Then

                    If m_SelEnd = m_CursorPos Then
                        If m_SelEnd > m_SelStart Then
                            m_SelEnd = tmpCursor
                        Else
                            m_SelStart = tmpCursor
                        End If
                    Else
                        m_SelStart = tmpCursor
                    End If
                    mustRedraw = True
                Else
                    If m_SelStart <> m_SelEnd Then mustRedraw = True
                    m_SelStart = tmpCursor
                    m_SelEnd = tmpCursor
                End If

                m_CursorPos = tmpCursor

            ElseIf Shift = 0 Then
                m_CursorPos = getPreviousChar(m_CursorPos - 1)
                If m_SelStart <> m_SelEnd Then mustRedraw = True
                m_SelStart = m_CursorPos
                m_SelEnd = m_CursorPos

            ElseIf (Shift And 1) Then
                If m_SelEnd = m_CursorPos Then
                    If m_SelStart < m_SelEnd Then
                        m_SelEnd = getPreviousChar(m_SelEnd - 1)
                        m_CursorPos = m_SelEnd
                    Else
                        m_SelStart = getPreviousChar(m_SelStart - 1)
                        m_CursorPos = m_SelStart
                    End If
                Else
                    m_SelStart = getPreviousChar(m_SelStart - 1)
                    m_CursorPos = m_SelStart
                End If
                mustRedraw = True
            End If

        Case vbKeyA To vbKeyZ
            'If (Shift And 1) Then
            '    'mustRedraw = AddCharAtCursor(UCase(Chr(KeyCode)))
            'ElseIf Shift = 0 Then
            '    'mustRedraw = AddCharAtCursor(LCase(Chr(KeyCode)))
            If (Shift And 2) Then
                Select Case KeyCode
                    Case vbKeyC, vbKeyX

                        tmpString = GetSelectionText()
                        If LenB(tmpString) > 0 Then
                            Clipboard.Clear
                            Clipboard.SetText tmpString
                        Else
                            Exit Sub
                        End If

                        If KeyCode = vbKeyX And Locked = False Then
                            mustRedraw = AddCharAtCursor()
                        End If

                    Case vbKeyV
                        If Locked Then Exit Sub
                        mustRedraw = AddCharAtCursor(Clipboard.GetText)
                    
                    Case vbKeyA
                        m_SelStart = 0
                        m_SelEnd = UBound(CharMap)
                        m_CursorPos = m_SelEnd
                        mustRedraw = True
                End Select

            End If

        'Case vbKey0 To vbKey9
            'If (Shift And 1) Then
            '
            'ElseIf Shift = 0 Then
            '    mustRedraw = AddCharAtCursor(Chr(KeyCode))
            'End If


        'Case vbKeySpace
        '    mustRedraw = AddCharAtCursor(" ")

        Case vbKeyReturn
            If Locked Then Exit Sub
            If m_bMultiLine Then mustRedraw = AddCharAtCursor(vbCrLf)

        Case vbKeyBack
            If Locked Then Exit Sub
            If m_SelStart = m_SelEnd Then
                If m_SelStart > 0 Then
                    m_SelStart = getPreviousChar(m_SelStart - 1)
                Else
                    Exit Sub
                End If
            End If

            mustRedraw = AddCharAtCursor()

        Case vbKeyDelete
            If Locked Then Exit Sub
            If m_SelEnd >= UBound(m_byteText) Then
                Exit Sub
            End If

            If m_SelStart = m_SelEnd Then
                m_SelEnd = getNextChar(m_SelStart + 1)
            End If

            mustRedraw = AddCharAtCursor()

    End Select


    If m_SelEnd < m_SelStart Then    'swap the 2 if the start is bigger than the end
        tmpswap = m_SelEnd
        m_SelEnd = m_SelStart
        m_SelStart = tmpswap
    End If

    If m_SelEnd < 0 Then m_SelEnd = 0
    If m_SelEnd > UBound(CharMap) Then m_SelEnd = UBound(CharMap)

    If m_SelStart < 0 Then m_SelStart = 0
    If m_SelStart > UBound(CharMap) Then m_SelStart = UBound(CharMap)

    If m_CursorPos < 0 Then m_CursorPos = 0
    If m_CursorPos > UBound(CharMap) Then m_CursorPos = UBound(CharMap)
    

    If CharMap(m_CursorPos).X >= m_lUsercontrolWidth Then
        m_lScrollLeft = m_lScrollLeft + (CharMap(m_CursorPos).X - m_lUsercontrolWidth)
        mustRedraw = True
    ElseIf CharMap(m_CursorPos).X <= m_lUsercontrolLeft Then
        m_lScrollLeft = m_lScrollLeft + (CharMap(m_CursorPos).X - m_lUsercontrolLeft)
        mustRedraw = True
    End If

    If getSelectionChanged() And tmpCursorUpDown = False Then
        m_SelUpDownTheSame = False
    End If
    
    If mustRedraw Then Redraw
    
    updateCaretPos
    'DoEvents
End Sub


Function GetSelectionText() As String
    If m_SelStart = m_SelEnd Then Exit Function

    Dim i As Long
    Dim tmpBuffer As String
    Dim tmpBufferLen As Long
    Dim TotalLen As Long

    For i = m_SelStart To m_SelEnd - 1
        tmpBufferLen = tmpBufferLen + 1

        tmpBuffer = tmpBuffer & ChrW(m_byteText(i))

        If tmpBufferLen > TotalLen Then
            TotalLen = tmpBufferLen
            tmpBufferLen = 0
            GetSelectionText = GetSelectionText & tmpBuffer
            tmpBuffer = ""
        End If

    Next i

    GetSelectionText = GetSelectionText & tmpBuffer

    'GetSelectionText = Mid(m_StrText, m_SelStart + 1, m_SelEnd + 1)

End Function


Function getNextWordFromCursor() As Long
    Dim i As Long
    Dim WordPart As Long

    WordPart = CharMap(m_CursorPos).p
    If WordPart = -1 Then
        For i = m_CursorPos To UBound(CharMap)
            WordPart = CharMap(i).p
            If WordPart <> -1 Then
                getNextWordFromCursor = WordMap(WordPart).s
                Exit Function
            End If
        Next i
        getNextWordFromCursor = UBound(CharMap)
    Else
        WordPart = WordPart + 1
        If WordPart > WordCount Then
            getNextWordFromCursor = WordMap(WordCount).s + WordMap(WordCount).l
        Else
            getNextWordFromCursor = WordMap(WordPart).s
            For i = WordMap(WordPart).s To UBound(CharMap)
                If m_byteText(i) <> 10 And m_byteText(i) <> 32 Then    'm_byteText(i) <> 13 And
                    getNextWordFromCursor = i
                    Exit Function
                End If
            Next i
        End If
    End If

End Function

Function getPreviousWordFromCursor() As Long
    Dim i As Long
    Dim WordPart As Long

    WordPart = CharMap(m_CursorPos).p
    If WordPart = -1 Then
        For i = m_CursorPos To 0 Step -1
            WordPart = CharMap(i).p
            If WordPart <> -1 Then
                getPreviousWordFromCursor = WordMap(WordPart).s
                Exit Function
            End If
        Next i
    Else
        If WordMap(WordPart).s = m_CursorPos Then
            WordPart = WordPart - 1
        End If

        If WordPart < 0 Then
            getPreviousWordFromCursor = 0
        Else
            For i = WordMap(WordPart).s To 0 Step -1
                If m_byteText(i) <> 10 And m_byteText(i) <> 32 Then    'm_byteText(i) <> 13 And
                    getPreviousWordFromCursor = i
                    Exit Function
                End If
            Next i
        End If

        'getNextWord = getNextChar(
    End If

End Function



Function getNextChar(lStart As Long) As Long
    Dim i As Long

    getNextChar = lStart

    For i = getNextChar To UBound(CharMap)
        If m_byteText(i) <> 10 Then    'm_byteText(i) <> 13 And
            getNextChar = i
            Exit Function
        End If
    Next i
End Function


Function getPreviousChar(lStart As Long) As Long
    Dim i As Long

    getPreviousChar = lStart

    For i = getPreviousChar To 0 Step -1
        If m_byteText(i) <> 10 Then    'm_byteText(i) <> 13 And
            getPreviousChar = i
            Exit Function
        End If
    Next i
End Function


Private Sub UserControl_Resize()
    m_bRowMapCalculated = False
    
    If Not m_bStarting Then
        Redraw
    End If
    
    
    
    '    Dim j As Long
    '    Dim b() As Byte
    '    Dim i As Long
    '
    '    Dim Start1 As Double
    '    Dim Start2 As Double
    '
    '    'b = StrConv("FF00FF", vbFromUnicode)
    '
    '    m_timer.tStart
    '
    '    For i = 0 To 100
    '        'j = RGBByte(0, b)
    '        ReCalculateWords1
    '    Next i
    '
    '    Start1 = m_timer.tStop
    '
    '
    '    m_timer.tStart
    '
    '    For i = 0 To 100
    '        'j = RGBByte(0, b)
    '        ReCalculateWords
    '    Next i
    '
    '    Start2 = m_timer.tStop
    '
    '    Debug.Print "RecalculateWords1() is  "; Round(Start1, 7); " "; Round(Start2, 7); "seconds faster"

End Sub


Function SizeByte(ByRef lStart As Long, ByRef lBytes() As Byte) As Long
    Dim c(0 To 10) As Long
    Dim lCount As Long
    Dim i As Long

    For i = 0 To 10
        Select Case lBytes(lStart + i)
            Case 32
                Exit For

            Case 48 To 57
                c(lCount) = lBytes(lStart + i) - 48
        End Select

        lCount = lCount + 1
    Next i

    'SizeByte = 0

    For i = 0 To lCount - 1
        SizeByte = SizeByte + c(i) * (10 ^ (lCount - i - 1))
    Next i
    lStart = lCount
End Function

Sub ReplaceWord(newText As String, Optional wordNr As Long = -2)
    If wordNr = -2 Then
        wordNr = getWordFromChar(m_CursorPos)
    End If
    
    If wordNr < 0 Then Exit Sub
    
    m_SelStart = WordMap(wordNr).s
    m_SelEnd = m_SelStart + WordMap(wordNr).l
    
    If m_SelEnd > UBound(CharMap) Then m_SelEnd = UBound(CharMap)
    
    m_CursorPos = m_SelEnd
    
    AddCharAtCursor newText
    
    If Not m_bStarting Then Redraw
End Sub


Sub ReCalculateMarkup()

'Dim TextOffsetX As Long
'Dim TextOffsetY As Long
'Dim TW As Long 'text width
'Dim TH As Long 'text height
    Dim MarkupList() As Current_Style
    'Dim MarkupEmpty As Current_Style

    Dim NTC As Long    'Normal Text Count

    Dim TLength As Long    'text length
    Dim TL As Long    'text length for loop
    'Dim MS As String 'mid string
    'Dim MA As Long 'mid ascii

    'Dim CS As Long 'command style
    Dim SS As Long    'seek string
    Dim FC As Long    'fore color
    'Dim MFC As String 'mid fore color
    Dim MC As Long    'markup count
    'Dim LNW As Long

    'currentStyle values
    Dim cForeColor As Long
    Dim cUnderline As Byte
    Dim cItalic As Byte
    Dim cBold As Byte
    Dim cMarking As Long
    Dim cFontSize As Long
    Dim cStrikeThrough As Byte
    Dim cLine As Long
    Dim cDescendHeight As Long
    Dim cTextMetric As TEXTMETRIC

    'Dim TTL As Long 'temp text length
    'Dim TS As String 'temp string

    Dim IgnoreNextChar As Boolean


    If m_bMarkupCalculating Then Exit Sub

    m_bMarkupCalculating = True

    cForeColor = -1
    cUnderline = m_StdFont.Underline
    cItalic = m_StdFont.Italic
    cBold = m_StdFont.Bold
    cFontSize = -1
    cMarking = -1
    cLine = -1
    cStrikeThrough = m_StdFont.Strikethrough

    GetTextMetrics UserControl.hdc, cTextMetric
    cDescendHeight = cTextMetric.tmDescent


    UserControl.Font = m_StdFont
    UserControl.BackColor = m_OleBackgroundColor
    UserControl.FontSize = m_StdFont.Size
    UserControl.FontStrikethru = m_StdFont.Strikethrough
    UserControl.FontUnderline = m_StdFont.Underline
    UserControl.FontItalic = m_StdFont.Italic
    UserControl.FontBold = m_StdFont.Bold


    TLength = Len(m_StrMarkupText)

    m_byteMarkupText = StrConv(m_StrMarkupText, vbFromUnicode)

    m_StrText = vbNullString

    '    If TLength = 0 Then
    '        ReDim m_byteText(0)
    '        ReDim CharMap(0)
    '        ReDim MarkupS(0)
    '        ReDim RowNumChars(0)
    '        Exit Sub
    '    End If

    ReDim m_byteText(0 To TLength)
    ReDim CharMap(0 To TLength)
    ReDim MarkupS(0 To TLength)
    'ReDim RowNumChars(0 To TLength)

    ReDim Preserve MarkupList(0 To 100)    '100 styles will be enough i think??

    'Exit Sub
    For TL = 0 To TLength


        'CS = 0
        'MS = Mid$(m_StrMarkupText, TL, 1)
        'MA = Asc(MS)

        'MA = m_byteMarkupText(TL)
        'MS = ChrW(MA)

        If IgnoreNextChar Then
            IgnoreNextChar = False
            GoTo DoNotCheck
        End If

        If TL = TLength Then GoTo DoNotCheck

        Select Case m_byteMarkupText(TL)

            Case 92    '"\" 'an new line maybe?
                'Select Case Asc(Mid$(m_StrMarkupText, TL + 1, 1))
                Select Case m_byteMarkupText(TL + 1)
                    Case 92, 123, 125
                        IgnoreNextChar = True
                        GoTo NextChar
                End Select


            Case 123    '"{" 'something importand starts

                'SS = InStr(TL + 1, m_StrMarkupText, " ")
                SS = InstrByte(TL + 1, m_byteMarkupText, 32)
                If SS > 0 Then
                    If m_byteMarkupText(TL + 1) <> 92 Then
                        MarkupList(MC).lStyle = Sel_Edit.sNone
                        MC = MC + 1
                        TL = TL + 3
                        GoTo NextChar
                    End If

                    Select Case m_byteMarkupText(TL + 2)

                        Case 98    '"\b"
                            'ReDim Preserve MarkupList(0 To MC)
                            MarkupList(MC).lStyle = Sel_Edit.sBold
                            MarkupList(MC).prev_Value = cBold
                            cBold = Not cBold
                            MC = MC + 1
                            TL = TL + 3

                        Case 117    '"\u"
                            'ReDim Preserve MarkupList(0 To MC)
                            MarkupList(MC).lStyle = Sel_Edit.sUnderline
                            MarkupList(MC).prev_Value = cUnderline
                            cUnderline = Not cUnderline
                            MC = MC + 1
                            TL = TL + 3

                        Case 105    '"\i"
                            'ReDim Preserve MarkupList(0 To MC)
                            MarkupList(MC).lStyle = Sel_Edit.sItalic
                            MarkupList(MC).prev_Value = cItalic
                            cItalic = Not cItalic
                            MC = MC + 1
                            TL = TL + 3

                        Case 99    '"\c"
                            'FC = InstrByte(TL + 3, m_byteMarkupText, 32)
                            MarkupList(MC).lStyle = Sel_Edit.sForeColor
                            MarkupList(MC).prev_Value = cForeColor
                            cForeColor = RGBByte(TL + 4, m_byteMarkupText)
                            MC = MC + 1
                            TL = TL + 10

                        Case 109    '"\m"
                            'FC = InstrByte(TL + 3, m_byteMarkupText, 32)
                            'FC = RGBByte(FC + 1, m_byteMarkupText)
                            FC = RGBByte(TL + 4, m_byteMarkupText)
                            MarkupList(MC).lStyle = Sel_Edit.sMarking
                            MarkupList(MC).prev_Value = cMarking
                            cMarking = FC
                            MC = MC + 1
                            TL = TL + 10

                        Case 102    '"\f"
                            'CS = Asc(Mid$(m_StrMarkupText, TL + 3, 1)) 'check for size or marking or border

                            Select Case m_byteMarkupText(TL + 3)
                                Case 115  '"\fs"
                                    FC = InstrByte(TL + 4, m_byteMarkupText, 32) + 1
                                    'FC =  'RGBByte(FC + 1, m_byteMarkupText)
                                    'ReDim Preserve MarkupList(0 To MC)
                                    MarkupList(MC).lStyle = sFontSize
                                    MarkupList(MC).prev_Value = cFontSize
                                    cFontSize = SizeByte(FC, m_byteMarkupText)

                                    MC = MC + 1
                                    TL = TL + 5 + FC

                                Case 109    '"\fm"
                                    'FC = InStr(TL + 4, m_StrMarkupText, " ")
                                    'MFC = Mid(m_StrMarkupText, FC + 1, 6)
                                    'FC = CLng("&h" & MFC)
                                    'ReDim Preserve MarkupList(0 To MC)
                                    MarkupList(MC).lStyle = sMarking
                                    MarkupList(MC).prev_Value = cMarking
                                    cMarking = RGBByte(TL + 5, m_byteMarkupText)
                                    MC = MC + 1
                                    TL = TL + 11

                                Case 98    '"\fb"
                                    'FC = InStr(TL + 4, m_StrMarkupText, " ")
                                    'MFC = Mid(m_StrMarkupText, FC + 1, 6)
                                    'FC = CLng("&h" & MFC)
                                    'ReDim Preserve MarkupList(0 To MC)
                                    MarkupList(MC).lStyle = sLine
                                    MarkupList(MC).prev_Value = cLine
                                    cLine = RGBByte(TL + 5, m_byteMarkupText)
                                    MC = MC + 1
                                    TL = TL + 11
                            End Select

                        Case 115    '"\s"
                            'ReDim Preserve MarkupList(0 To MC)
                            MarkupList(MC).lStyle = sStrikeThrough
                            MarkupList(MC).prev_Value = cStrikeThrough
                            cStrikeThrough = Not cStrikeThrough
                            MC = MC + 1
                            TL = TL + 3


                    End Select
                End If

                GoTo NextChar

            Case 125    '"}"
                If MC > 0 Then
                    MC = MC - 1

                    Select Case MarkupList(MC).lStyle
                        Case sNone

                        Case sBold
                            cBold = CBool(MarkupList(MC).prev_Value)

                        Case sUnderline
                            cUnderline = CBool(MarkupList(MC).prev_Value)

                        Case sItalic
                            cItalic = CBool(MarkupList(MC).prev_Value)

                        Case sForeColor
                            cForeColor = CLng(MarkupList(MC).prev_Value)

                        Case sMarking
                            cMarking = CLng(MarkupList(MC).prev_Value)

                        Case sFontSize
                            cFontSize = CLng(MarkupList(MC).prev_Value)

                        Case sStrikeThrough
                            cStrikeThrough = CBool(MarkupList(MC).prev_Value)

                        Case sLine
                            cLine = CLng(MarkupList(MC).prev_Value)

                    End Select

                End If

                GoTo NextChar
        End Select

DoNotCheck:


        If cFontSize <> UserControl.FontSize Then
            If cFontSize = -1 Then
                UserControl.FontSize = m_StdFont.Size
            Else
                UserControl.FontSize = cFontSize
            End If

            GetTextMetrics UserControl.hdc, cTextMetric
            cDescendHeight = cTextMetric.tmDescent
        End If

        If cBold <> UserControl.FontBold Then
            If cBold = -1 Then
                UserControl.FontBold = m_StdFont.Bold
            Else
                UserControl.FontBold = cBold
            End If
        End If

        If cItalic <> UserControl.FontItalic Then
            If cItalic = -1 Then
                UserControl.FontItalic = m_StdFont.Italic
            Else
                UserControl.FontItalic = cItalic
            End If
        End If

        If cUnderline <> UserControl.FontUnderline Then
            If cUnderline = -1 Then
                UserControl.FontUnderline = m_StdFont.Underline
            Else
                UserControl.FontUnderline = cUnderline
            End If
        End If




        If TL <> TLength Then
            If Not (m_byteMarkupText(TL) = 13 Or m_byteMarkupText(TL) = 10) Then
                GetTextSize Chr(m_byteMarkupText(TL)), CharMap(NTC)
            Else
                GetTextSize " ", CharMap(NTC)
                CharMap(NTC).W = 0
            End If
            m_byteText(NTC) = m_byteMarkupText(TL)
        Else
            GetTextSize " ", CharMap(NTC)
            m_byteText(NTC) = 32
        End If

        If cLine <> -1 Then
            CharMap(NTC).W = CharMap(NTC).W + 2
        End If



        CharMap(NTC).d = cDescendHeight

        MarkupS(NTC).lBold = cBold
        MarkupS(NTC).lMarking = cMarking
        MarkupS(NTC).lForeColor = cForeColor
        MarkupS(NTC).lUnderline = cUnderline
        MarkupS(NTC).lItalic = cItalic
        MarkupS(NTC).lFontSize = cFontSize
        MarkupS(NTC).lStrikeThrough = cStrikeThrough
        MarkupS(NTC).lLine = cLine


        NTC = NTC + 1


NextChar:

    Next TL

    ReDim Preserve m_byteText(0 To NTC - 1)
    ReDim Preserve MarkupS(0 To NTC - 1)
    ReDim Preserve CharMap(0 To NTC - 1)

    'm_StrText = StrConv(m_byteText, vbUnicode)

DoneRefreshing:

    m_StrMarkupText = ""
    ReDim m_byteMarkupText(0)

    'RowNum = NRC
    m_bMarkupCalculated = True
    m_bMarkupCalculating = False


End Sub



Function Fmat(str As String, length As Long) As String
    Dim strLength As Long
    strLength = Len(str)

    If strLength > length Then
        Fmat = String(length, "X")
    ElseIf strLength < length Then
        Fmat = String(length - strLength, "0") & str
    Else
        Fmat = str
    End If

End Function



Function GetMidText(sString As String, sSearch As String, sSearch2 As String, Optional lStart As Long = 1) As String
    Dim tmp1 As Long
    Dim tmp2 As Long

    If lStart < 1 Then Exit Function

    tmp1 = InStr(lStart, sString, sSearch)
    If tmp1 = 0 Then Exit Function
    tmp1 = tmp1 + Len(sSearch)


    tmp2 = InStr(tmp1, sString, sSearch2)
    If tmp2 = 0 Then Exit Function

    GetMidText = Mid$(sString, tmp1, tmp2 - tmp1)
End Function


Private Sub UserControl_Paint()
    If Not m_bStarting Then Redraw
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackgroundColor", m_OleBackgroundColor, &HFFFFFF
        .WriteProperty "BorderColor", m_OleBorderColor, &HFFFFFF
        .WriteProperty "Font", m_StdFont, Ambient.Font
        .WriteProperty "ForeColor", m_OleForeColor, &H0
        .WriteProperty "Border", m_bBorder, True
        .WriteProperty "BorderThickness", m_lBorderThickness, 1
        .WriteProperty "MousePointer", m_MouMousePointer, 0
        .WriteProperty "Border", m_bBorder, True
        .WriteProperty "LineNumbers", m_bLineNumbers, False
        .WriteProperty "LineNumberForeColor", m_OleLineNumberForeColor, vbWhite
        .WriteProperty "LineNumberBackground", m_OleLineNumberBackground, vbBlack
        .WriteProperty "ConsoleColors", m_bConsoleColors, True
        
        .WriteProperty "RowLines", m_bRowLines, False
        .WriteProperty "RowLineColor", m_OleRowLineColor, &HEEEEEE
        .WriteProperty "RowNumberOnEveryLine", m_bRowNumberOnEveryLine, False
        .WriteProperty "WordWrap", m_bWordWrap, False
        .WriteProperty "MultiLine", m_bMultiLine, False
        .WriteProperty "HideCursor", m_bHideCursor, False
        .WriteProperty "AutoResize", m_bAutoResize, False
        .WriteProperty "Locked", m_bLocked, False
        .WriteProperty "ScrollBars", m_sScrollBars, ScrollBarStyle.lNone
        
    End With
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_OleBackgroundColor = .ReadProperty("BackgroundColor", &HFFFFFF)
        m_OleBorderColor = .ReadProperty("BorderColor", &HFFFFFF)
        Set Font = .ReadProperty("Font", Ambient.Font)
        m_OleForeColor = .ReadProperty("ForeColor", &H0)
        m_bBorder = .ReadProperty("Border", True)
        m_lBorderThickness = .ReadProperty("BorderThickness", 1)
        MousePointer = .ReadProperty("MousePointer", 0)
        m_bLineNumbers = .ReadProperty("LineNumbers", False)
        m_OleLineNumberForeColor = .ReadProperty("LineNumberForeColor", vbWhite)
        m_OleLineNumberBackground = .ReadProperty("LineNumberBackground", vbBlack)
        m_bConsoleColors = .ReadProperty("ConsoleColors", True)
        
        m_bRowLines = .ReadProperty("RowLines", False)
        m_OleRowLineColor = .ReadProperty("RowLineColor", &HEEEEEE)
        m_bRowNumberOnEveryLine = .ReadProperty("RowNumberOnEveryLine", False)
        m_bWordWrap = .ReadProperty("WordWrap", False)
        m_bMultiLine = .ReadProperty("MultiLine", False)
        m_bHideCursor = .ReadProperty("HideCursor", False)
        m_bAutoResize = .ReadProperty("AutoResize", False)
        m_bLocked = .ReadProperty("Locked", False)
        m_sScrollBars = .ReadProperty("ScrollBars", ScrollBarStyle.lNone)
        
    End With
    
    m_bStarting = False
    
    Redraw
End Sub






