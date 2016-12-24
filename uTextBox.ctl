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

'Private Type Sel_Normal
'    lStart As Long
'    lLength As Long
'    bValue As Boolean
'End Type

'Private Type Sel_Font
'    lStart As Long
'    lLength As Long
'    sFontName As String
'End Type

'Private Type Sel_Color
'    lStart As Long
'    lLength As Long
'    lColor As Long
'End Type

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
    W As Long
    H As Long
    d As Long
    x As Long
    y As Long
End Type

Private Type WHSL
    W As Long
    H As Long
    s As Long
    L As Long
End Type

Private Type NSS
    NumChars As Long
    StartY As Long
    StartChar As Long
End Type

Private Type MarkupStyles
    'lFontName As String
    lForeColor As Long
    lUnderline As Boolean
    lItalic As Boolean
    lBold As Boolean
    lMarking As Long
    lFontSize As Long
    lPartOfWord As Long
    lStrikeThrough As Boolean
    lLine As Long
End Type

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
Private m_SelEnd As Long
Private m_SelUpDownTheSame As Boolean

Private m_bRefreshing As Boolean
Private m_bRefreshedWhileBusy As Boolean

'Private m_LonFontCount As Long
'Private m_SelFont() As Sel_Font
'Private m_LonColorCount As Long
'Private m_SelColor() As Sel_Color
'Private m_LonMarkingCount As Long
'Private m_SelMarking() As Sel_Color
'Private m_LonUnderlineCount As Long
'Private m_SelUnderline() As Sel_Normal
'Private m_LonItalicCount As Long
'Private m_SelItalic() As Sel_Normal
'Private m_LonBoldCount As Long
'Private m_SelBold() As Sel_Normal

Private m_bLineNumbers As Boolean
Private m_bMarkupCalculated As Boolean
Private m_bMarkupCalculating As Boolean

Private m_bWordsCalculated As Boolean
Private m_bWordsCalculating As Boolean

Private m_bHideCursor As Boolean

Private m_bMultiLine As Boolean
Private m_bRowLines As Boolean
Private m_bAutoResize As Boolean

Private m_OleRowLineColor As OLE_COLOR
Private m_OleLineNumberBackground As OLE_COLOR
Private m_bRowNumberOnEveryLine As Boolean

Public Enum ScrollBarStyle
    lNone = 0
    lVertical = 1
    lHorizontal = 2
    lBoth = 3
End Enum

Private m_sScrollBars As ScrollBarStyle

Public m_lScrollLeft As Long
Public m_lScrollLeftMax As Long
Public m_lScrollTop As Long
Public m_lScrollTopMax As Long

'Private m_timer As clsTimer

Public Event Changed()
Public Event SelectionChanged()
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)

Private m_lUsercontrolHeight As Long
Private m_lUsercontrolWidth As Long
Private m_lUsercontrolLeft As Long
Private m_lUsercontrolTop As Long

Private m_bBlockNextKeyPress As Boolean 'for things like ctrl+space autocomplete

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As WH) As Long
Private Declare Function SetTextAlign Lib "gdi32.dll" (ByVal hdc As Long, ByVal wFlags As Long) As Long

Private Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function setCaretPos Lib "user32" Alias "SetCaretPos" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long



Public Function getWordFromChar(char As Long) As Long
    getWordFromChar = MarkupS(char).lPartOfWord
End Function

Public Function getWordLength(word As Long) As Long
    getWordLength = WordMap(word).L
End Function

Public Function getWordStart(word As Long) As Long
    getWordStart = WordMap(word).s
End Function


Public Sub setCharItallic(char As Long, bValue As Boolean)
    MarkupS(char).lItalic = bValue
End Sub

Public Sub setCharBold(char As Long, bValue As Boolean)
    MarkupS(char).lBold = bValue
End Sub

Public Sub setCharForeColor(char As Long, OleValue As OLE_COLOR)
    MarkupS(char).lForeColor = IIf(OleValue >= 0, OleValue, -1)
End Sub

Public Sub setCharBackColor(char As Long, OleValue As OLE_COLOR)
    MarkupS(char).lMarking = IIf(OleValue >= 0, OleValue, -1)
End Sub


Public Function getCharItallic(char As Long) As Boolean
    getCharItallic = MarkupS(char).lItalic
End Function

Public Function getCharBold(char As Long) As Boolean
    getCharBold = MarkupS(char).lBold
End Function

Public Function getCharForeColor(char As Long) As OLE_COLOR
    getCharForeColor = MarkupS(char).lForeColor
End Function

Public Function getCharBackColor(char As Long) As OLE_COLOR
   getCharBackColor = MarkupS(char).lMarking
End Function


Sub updateCaretPos()
    If Not Screen.ActiveControl Is Nothing Then
        If Not UserControl.Extender Is Screen.ActiveControl Then Exit Sub
    End If
    
    If m_bHideCursor Then Exit Sub
    
    CreateCaret UserControl.hWnd, 0, 2, CharMap(m_CursorPos).H

    setCaretPos CharMap(m_CursorPos).x - 1, CharMap(m_CursorPos).y - CharMap(m_CursorPos).H + CharMap(m_CursorPos).d
    ShowCaret UserControl.hWnd
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



Public Property Get MultiLine() As Boolean
    MultiLine = m_bMultiLine
End Property

Public Property Let MultiLine(ByVal bValue As Boolean)
    m_bMultiLine = bValue
    PropertyChanged "MultiLine"
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
    SelStart = m_SelEnd - m_SelStart
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


Public Property Get Border() As Boolean
    Border = m_bBorder
End Property

Public Property Let Border(ByVal bValue As Boolean)
    m_bBorder = bValue
    PropertyChanged "Border"
    If Not m_bStarting Then Redraw
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

Private Sub UserControl_DblClick()
    Dim word As Long
    
    word = MarkupS(m_CursorPos).lPartOfWord
    
    If word = -1 And m_CursorPos > 0 Then
        word = MarkupS(m_CursorPos - 1).lPartOfWord
    End If
        
    If word <> -1 Then
        m_SelStart = WordMap(word).s
        m_SelEnd = WordMap(word).s + WordMap(word).L
        If m_SelEnd > UBound(CharMap) Then m_SelEnd = UBound(CharMap)
        m_CursorPos = m_SelEnd
        If Not m_bStarting Then Redraw
        updateCaretPos
    End If
End Sub

Private Sub UserControl_GotFocus()
    updateCaretPos
End Sub

Private Sub UserControl_Initialize()
    m_bStarting = True

    'Set m_timer = New clsTimer

    'Dim lrand As Long
    Dim newChar As String

    Dim i As Long
    'Dim MS As String 'mid string

    Dim constString As String
    Const randomMarkup As Boolean = True
    
'
'    For i = 0 To 5
'        constString = constString & "Deze Textbox is gemaakt door Ricardo de Roode!                HierNogEvenEenLangWoord" & vbCrLf    '& vbCrLf
'    Next i
'
'
'    If randomMarkup Then
'        For i = 1 To Len(constString)
'            newChar = ""
'            '{\c FFFF00 hoi {\c FF00FF hallo dit is magenta gekleurde text} hoi}
'            If Mid$(constString, i, 1) <> " " And Mid$(constString, i, 1) <> vbCr And Mid$(constString, i, 1) <> vbLf Then
'                newChar = "{\c " & Fmat(Hex(CLng(Rnd * 255)), 2) & Fmat(Hex(CLng(Rnd * 255)), 2) & Fmat(Hex(CLng(Rnd * 255)), 2) & " "
'                newChar = newChar & "{\fb " & Fmat(Hex(CLng(Rnd * 255)), 2) & Fmat(Hex(CLng(Rnd * 255)), 2) & Fmat(Hex(CLng(Rnd * 255)), 2) & " "
'                newChar = newChar & "{\m " & Fmat(Hex(CLng(Rnd * 255)), 2) & Fmat(Hex(CLng(Rnd * 255)), 2) & Fmat(Hex(CLng(Rnd * 255)), 2) & " "
'                'newChar = newChar & "{\m FF00FF "
'                newChar = newChar & "{\fs " & Fix(Rnd * 16 + 8) & " "
'                'newChar = newChar & "{\i "
'
'                '            Select Case Round(Rnd * 3)
'                '                Case 0
'                '                    newChar = newChar & "{\i "
'                '                Case 1
'                '                    newChar = newChar & "{\b "
'                '                Case 2
'                '                    newChar = newChar & "{\u "
'                '                Case 3
'                '                    newChar = newChar & "{\s "
'                '            End Select
'
'                Select Case Mid$(constString, i, 1)
'                    Case "}", "{", "\"
'                        newChar = newChar & "\" & Mid$(constString, i, 1)
'
'                    Case Else
'                        newChar = newChar & Mid$(constString, i, 1)
'                End Select
'
'                'SnewChar = newChar & "}"
'                newChar = newChar & "}"
'                'newChar = newChar & "}"
'                newChar = newChar & "}"
'                newChar = newChar & "}"
'                newChar = newChar & "}"
'
'                m_StrMarkupText = m_StrMarkupText & newChar
'            Else
'                m_StrMarkupText = m_StrMarkupText & Mid$(constString, i, 1)
'
'            End If
'
'
'
'
'        Next i
'    Else
'        m_StrMarkupText = constString
'
'    End If


    m_OleRowLineColor = &HEEEEEE
    m_bRowLines = False

    m_bLineNumbers = False
    m_bRowLines = False
    m_OleLineNumberBackground = 0
    m_bRowNumberOnEveryLine = False
    m_lMouseDownPrevious = 99
    
End Sub

Sub DrawScrollBars(ByVal UW As Long, ByVal UH As Long, ByVal UHS As Long, ByVal UWS As Long, ByVal TSP As Long)
    Dim d1 As Double
    Dim d2 As Double
    Dim d3 As Double
    
    UH = UserControl.ScaleHeight
    UW = UserControl.ScaleWidth
    
    d1 = UWS / 15
    d2 = d1 * 1.73205
    d3 = d1 * 3
    
    If m_sScrollBars = lHorizontal Or m_sScrollBars = lBoth Then
        
        pts(0).x = UW - UWS
        pts(0).y = 0

        pts(1).x = UW
        pts(1).y = 0
        
        pts(2).x = UW
        pts(2).y = UH

        pts(3).x = UW - UWS
        pts(3).y = UH

        Polygon UserControl.hdc, pts(0), 4
            
        UserControl.Line (UW - UWS, UH - UHS)-(UW, UH - UHS), m_OleForeColor     'bottom
        UserControl.Line (UW - UWS, UHS - 1)-(UW, UHS - 1), m_OleForeColor    'top
        
        'triangle bottom
        UserControl.Line (Fix(UW - UWS / 2 - d3), Fix(UH - UHS / 2 - d2))-(Fix(UW - UWS / 2 + d3), Fix(UH - UHS / 2 - d2)) '_
        UserControl.Line (Fix(UW - UWS / 2 + d3), Fix(UH - UHS / 2 - d2))-(Fix(UW - UWS / 2 - 1), Fix(UH - UHS / 2 + d2)) ' /
        UserControl.Line (Fix(UW - UWS / 2 - d3), Fix(UH - UHS / 2 - d2))-(Fix(UW - UWS / 2 + 1), Fix(UH - UHS / 2 + d2)) '\
        

        'triangle top
        UserControl.Line (Fix(UW - UWS / 2 - d3), Fix(UHS / 2 + d2))-(Fix(UW - UWS / 2 + 1), Fix(UHS / 2 - d2)) '/
        UserControl.Line (Fix(UW - UWS / 2 + d3), Fix(UHS / 2 + d2))-(Fix(UW - UWS / 2 - 1), Fix(UHS / 2 - d2)) ' \
        UserControl.Line (Fix(UW - UWS / 2 - d3), Fix(UHS / 2 + d2))-(Fix(UW - UWS / 2 + d3), Fix(UHS / 2 + d2)) '_
        
    End If
    

    If m_sScrollBars = lVertical Or m_sScrollBars = lBoth Then
        
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
            pts(0).x = UWS + 2
            pts(0).y = UH - UHS + 2
    
            pts(1).x = pts(0).x
            pts(1).y = UH - 3
            
            pts(2).x = (UW - UWS * IIf(m_sScrollBars = lBoth, 3, 2) - 3) - (UW - UWS * IIf(m_sScrollBars = lBoth, 3, 2) - 3) * (1 / (m_lScrollLeftMax + UW) * m_lScrollLeftMax)
            If pts(2).x < 10 Then pts(2).x = 10
            pts(2).x = pts(2).x + pts(0).x
            
            pts(2).y = pts(1).y
    
            pts(3).x = pts(2).x
            pts(3).y = pts(0).y
            
            Polygon UserControl.hdc, pts(0), 4
        End If
        
        'triangle left
        UserControl.Line (Fix(UWS / 2 + d2), Fix(UH - UHS / 2 - d3))-(Fix(UWS / 2 + d2), Fix(UH - UHS / 2 + d3)) ' |
        UserControl.Line (Fix(UWS / 2 + d2), Fix(UH - UHS / 2 - d3))-(Fix(UWS / 2 - d2), Fix(UH - UHS / 2 + 1)) '/
        UserControl.Line (Fix(UWS / 2 + d2), Fix(UH - UHS / 2 + d3))-(Fix(UWS / 2 - d2), Fix(UH - UHS / 2 - 1))  '\
        
    End If
    
End Sub


Sub Redraw()
    If m_bRefreshing Then
        Exit Sub
    End If
    
    m_bRefreshing = True
    
    'm_timer.tStart

    If m_bMarkupCalculated = False Then ReCalculateMarkup
    If m_bWordsCalculated = False Then ReCalculateWords


    Dim i As Long
    Dim cc As Long    'Char Count
    Dim TL As Long    'text length
    
    Dim TW As Long    'text width
    Dim LNW As Long    'line number width
    Dim LNR As Long    'line number right
    Dim TextOffsetX As Long
    Dim TextOffsetY As Long
    Dim NRC As Long    'Number Row Count

    Dim UW As Long    'usercontrol width without scrollbars
    Dim UWS As Long    'usercontrol width
    Dim UH As Long    'usercontrol height without scrollbars
    Dim UHS As Long    'usercontrol height
    Dim RH As Long    'row height
    Dim RD As Long    'row d height

    Dim RL As Long    'row loop
    Dim TTW As Long    'temp text width
    Dim MTW As Long   'max text width

    Dim NLNR As Boolean    'Next Loop goto NextRow

    'currentStyle values
    Dim cForeColor As Long
    Dim cUnderline As Boolean
    Dim cItalic As Boolean
    Dim cBold As Boolean
    Dim cMarking As Long
    Dim cFontSize As Long
    Dim cStrikeThrough As Boolean
    Dim cLine As Long
    Dim TSP As Long    'text spacing
    Dim POWC As Long    'part of word checked

    UserControl.Cls

    UserControl.Font = m_StdFont
    UserControl.ForeColor = m_OleForeColor
    UserControl.BackColor = m_OleBackgroundColor
    UserControl.FontSize = m_StdFont.Size
    UserControl.FontStrikethru = m_StdFont.Strikethrough
    UserControl.FontUnderline = m_StdFont.Underline
    UserControl.FontItalic = m_StdFont.Italic
    UserControl.FontBold = m_StdFont.Bold

    UserControl.FillStyle = vbFSSolid
    UserControl.DrawStyle = 5
    UserControl.DrawMode = 13

    SetTextAlign UserControl.hdc, 24
    
    'TH = TextHeight("W")
    TW = TextWidth("9999")
    TSP = 6
    POWC = -1

    TL = UBound(m_byteText)
    'm_sScrollBars = lVertical

    If m_sScrollBars <> lNone Then
        UWS = 15
        UHS = 15
        UW = UserControl.ScaleWidth    ' - UWS
        UH = UserControl.ScaleHeight    ' - UHS
      
        If m_sScrollBars = lVertical Or m_sScrollBars = lBoth Then
            UH = UH - UHS
        End If
        
        If m_sScrollBars = lHorizontal Or m_sScrollBars = lBoth Then
            UW = UW - UWS - TSP
        Else
            UW = UW - TSP
        End If
    
    Else
        UW = UserControl.ScaleWidth - TSP
        UH = UserControl.ScaleHeight
    End If

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
    
    ReDim pts(0 To 3)
    ReDim RowMap(0 To 200)
    TTW = LNW
    RH = 0
    RD = 0

    'm_timer.tStart
    'On Error Resume Next
    
    For cc = 0 To TL
        'Debug.Print m_byteText(CC);
        
        If cBold <> MarkupS(cc).lBold Then
            cBold = MarkupS(cc).lBold
            UserControl.FontBold = cBold
        End If

        If cUnderline <> MarkupS(cc).lUnderline Then
            cUnderline = MarkupS(cc).lUnderline
            UserControl.FontUnderline = cUnderline
        End If

        If cItalic <> MarkupS(cc).lItalic Then
            cItalic = MarkupS(cc).lItalic
            UserControl.FontItalic = cItalic
        End If

        If cFontSize <> MarkupS(cc).lFontSize Then
            cFontSize = MarkupS(cc).lFontSize
            If cFontSize = -1 Then
                UserControl.FontSize = m_StdFont.Size
            Else
                UserControl.FontSize = cFontSize
            End If
        End If

        If cStrikeThrough <> MarkupS(cc).lStrikeThrough Then
            cStrikeThrough = MarkupS(cc).lStrikeThrough
            UserControl.FontStrikethru = cStrikeThrough
        End If


        If NLNR = True Or cc = 0 Then
            GoTo MakeNewRule
        End If
        
checkNextChar:
        
        Select Case m_byteText(cc)
            Case 13
                If m_bMultiLine Then NLNR = True
            Case 10
                CharMap(cc).x = TextOffsetX
                CharMap(cc).y = TextOffsetY
                GoTo NextChar
            Case 32
                'If TL = CC Then GoTo NextChar
                'If m_bWordWrap And TextOffsetX + CharMap(cc).W > UW Then
                '    GoTo NextChar  'TextOffsetX = LNW Or
                'End If
        End Select
        

        If MarkupS(cc).lPartOfWord <> -1 Then
            If POWC <> MarkupS(cc).lPartOfWord Then
                POWC = MarkupS(cc).lPartOfWord

                'does the current word fit?
                If m_bWordWrap And TextOffsetX + WordMap(POWC).W > UW And POWC > 0 Then
MakeNewRule:
                    TextOffsetX = LNW - m_lScrollLeft
                    TTW = TextOffsetX
                    RH = 0
                    RD = 0
                    
                    If WordWrap Then
                        If cc = 0 Then
                            POWC = 0
                        End If
                        
                        For RL = POWC To WordCount
                            TTW = TTW + WordMap(RL).W
                            If TTW > UW And RL > POWC Then Exit For
                            If WordMap(RL).H > RH Then RH = WordMap(RL).H
                        Next RL
                    Else
                        For RL = cc To UBound(m_byteText)
                            TTW = TTW + CharMap(RL).W
                            
                            If m_byteText(RL) = 10 Then Exit For
                            If CharMap(RL).H - CharMap(RL).d > RH Then RH = CharMap(RL).H - CharMap(RL).d
                            If CharMap(RL).d > RD Then RD = CharMap(RL).d
                            If TTW > MTW Then MTW = TTW
                        Next RL
                    End If
                    
                    If cc = 0 Then
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
                    
                    If m_bRowLines Then
                        If TextOffsetY < UH Then
                            UserControl.DrawStyle = 0
                            UserControl.Line (LNW, TextOffsetY)-(UW, TextOffsetY), m_OleRowLineColor
                            UserControl.DrawStyle = 5
                        End If
                    End If
                    
                    If m_bRowNumberOnEveryLine Or NLNR Or cc = 0 Then
                        If cc <> 0 Then NRC = NRC + 1
                        RowMap(NRC).StartY = TextOffsetY
                        RowMap(NRC).StartChar = cc
                    End If
                    
                    
                    
                    If NLNR = True Then
                        NLNR = False
                        GoTo checkNextChar
                    End If
                End If

                'if the word started on the previous row check to break again for really long words!
            ElseIf m_bWordWrap And TextOffsetX + CharMap(cc).W > UW Then
                GoTo MakeNewRule
            End If
        End If

        
        If TextOffsetY < UH And TextOffsetX - TSP < UW And TextOffsetX + CharMap(cc).W > 0 Then  '
            Dim jj As Long
            Dim kk As Long

            If cMarking <> MarkupS(cc).lMarking Then
                cMarking = MarkupS(cc).lMarking
                If cMarking <> -1 Then
                    UserControl.FillColor = MarkupS(cc).lMarking
                End If
            End If


            If cMarking <> -1 Then
                pts(0).x = TextOffsetX
                pts(0).y = TextOffsetY + CharMap(cc).d

                pts(1).x = TextOffsetX + CharMap(cc).W
                pts(1).y = pts(0).y

                pts(2).x = pts(1).x
                pts(2).y = pts(0).y - CharMap(cc).H 'TextOffsetY - CharMap(CC).H + CharMap(CC).d

                pts(3).x = pts(0).x
                pts(3).y = pts(2).y
                'UserControl.DrawMode = 15

                Polygon UserControl.hdc, pts(0), 4
                'UserControl.DrawMode = 13
            End If

            cLine = MarkupS(cc).lLine


            If cLine <> -1 Then
                If cLine <> cForeColor Then
                    cForeColor = cLine
                    UserControl.ForeColor = cLine
                End If

                Dim MS As String
                MS = ChrW(m_byteText(cc))
                For jj = 0 To 2
                    For kk = -2 To 0
                        If Not (jj = 0 And kk = 0) Then
                            UserControl.CurrentX = TextOffsetX + jj ' + 1
                            UserControl.CurrentY = TextOffsetY + kk '- 1    '- CharMap(CC).H
                            UserControl.Print MS;
                        End If
                    Next kk
                Next jj
                UserControl.CurrentX = TextOffsetX + 1
                UserControl.CurrentY = TextOffsetY - 1    ' - CharMap(CC).H
            Else
                UserControl.CurrentX = TextOffsetX
                UserControl.CurrentY = TextOffsetY    ' - CharMap(CC).H
            End If

            If cForeColor <> MarkupS(cc).lForeColor Then
                cForeColor = MarkupS(cc).lForeColor
                If cForeColor = -1 Then
                    UserControl.ForeColor = m_OleForeColor
                Else
                    UserControl.ForeColor = cForeColor
                End If
            End If

            UserControl.Print ChrW(m_byteText(cc));

            If cc >= m_SelStart And cc < m_SelEnd Then

                pts(0).x = TextOffsetX
                pts(0).y = TextOffsetY + CharMap(cc).d

                pts(1).x = TextOffsetX + CharMap(cc).W
                pts(1).y = pts(0).y

                pts(2).x = pts(1).x
                pts(2).y = TextOffsetY - RH + IIf(m_bMultiLine, CharMap(cc).d, 0)

                pts(3).x = TextOffsetX
                pts(3).y = pts(2).y
                
                UserControl.DrawMode = 6 '6
                Polygon UserControl.hdc, pts(0), 4
                UserControl.DrawMode = 13
            End If

        
        ElseIf TextOffsetY >= UH Then
            GoTo DoneRefreshing
        End If
        
        CharMap(cc).x = TextOffsetX
        CharMap(cc).y = TextOffsetY
        
        RowMap(NRC).NumChars = RowMap(NRC).NumChars + 1

        TextOffsetX = TextOffsetX + CharMap(cc).W

NextChar:

    Next cc
DoneRefreshing:
    
    m_lUsercontrolHeight = UH
    m_lUsercontrolWidth = UW
    m_lUsercontrolTop = TSP
    m_lUsercontrolLeft = LNW + TSP
    
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
    UserControl.FillColor = m_OleBackgroundColor
    
    ReDim Preserve RowMap(0 To NRC)
    
    If m_bLineNumbers Then
        pts(0).x = 0:         pts(0).y = 0
        pts(1).x = LNR + TSP: pts(1).y = 0
        pts(2).x = pts(1).x:  pts(2).y = UH
        pts(3).x = 0:         pts(3).y = UH
        Polygon UserControl.hdc, pts(0), 4
        
        For i = 0 To NRC
            TW = UserControl.TextWidth(i + 1)
            UserControl.CurrentX = LNR - TW
            If RowMap(i).StartY < UH Then
                UserControl.CurrentY = RowMap(i).StartY    ' - TH
                UserControl.Print CStr(i + 1)
                UserControl.Line (TSP, RowMap(i).StartY)-(LNR, RowMap(i).StartY), m_OleRowLineColor
            End If
        Next i
    End If
    
    m_lScrollLeftMax = m_lScrollLeft + (MTW - UW)
    If m_lScrollLeftMax > 0 And m_lScrollLeft > m_lScrollLeftMax Then m_lScrollLeft = m_lScrollLeftMax

    If m_sScrollBars <> lNone Then
        DrawScrollBars UW, UH, UHS, UWS, TSP
    End If
    
    If m_bAutoResize Then
        'If m_lScrollLeftMax <> 0 Then
            UserControl.Width = ScaleX(MTW + TSP, vbPixels, vbTwips)
        'End If
    End If
    
    
    If m_bBorder Then
        UserControl.Line (0, 0)-(0, UserControl.ScaleHeight), m_OleBorderColor
        UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), m_OleBorderColor
        UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), m_OleBorderColor
        UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), m_OleBorderColor

        UserControl.Line (UserControl.ScaleWidth, 0)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), m_OleBorderColor
    End If

    UserControl.Refresh

    'DoEvents
    m_bRefreshing = False
    'If m_bRefreshedWhileBusy Then
    '    m_bRefreshedWhileBusy = False
    '    Redraw
    'End If
End Sub


Sub ReCalculateWords()
    Dim WC As Long    'word count
    Dim WH As Long    'word height
    Dim WW As Long    'word width
    Dim WL As Long    'word length
    Dim BT As Long    'ByteText
    Dim UB As Long    'ubound bytetext
    Dim TL As Long    'text length
    If m_bMarkupCalculating Then Exit Sub
    m_bMarkupCalculating = True

    On Error GoTo endff
    ReDim WordMap(0 To UBound(m_byteText) + 2)
    
    UB = UBound(m_byteText)
    For TL = 0 To UB
        BT = m_byteText(TL)
        
        If TL < UB And (BT = 32 Or BT = 10 Or (BT >= 40 And BT <= 47) Or BT = 58 Or BT = 59) Then      ' a space  Or m_byteText(TL) = 13
            If WL >= 0 Then
                WordMap(WC).H = WH
                WordMap(WC).W = WW
                WordMap(WC).L = WL
                'If m_byteText(TL) <> 10 Then
                WC = WC + 1
                'End If
                WH = 0
                WW = 0
                WL = 0

                WordMap(WC).s = TL + 1
                MarkupS(TL).lPartOfWord = -1
            End If
        Else
            MarkupS(TL).lPartOfWord = WC
            If CharMap(TL).H > WH Then
                WH = CharMap(TL).H
            End If
            WW = WW + CharMap(TL).W
            WL = WL + 1

        End If
    Next TL


    WordMap(WC).H = WH
    WordMap(WC).W = WW
    WordMap(WC).L = WL

    WordCount = WC
endff:

    m_bMarkupCalculating = False
    m_bWordsCalculated = True
End Sub


'Sub ReCalculateWords1()
'    Dim WC As Long 'word count
'    Dim WH As Long 'word height
'    Dim WW As Long 'word width
'    Dim WL As Long 'word length
'
'    Dim TL As Long 'text length
'    'Dim MS As String 'mid string
'
'    'TLength = Len(m_StrText)
'    On Error GoTo endff
'    ReDim WordMap(0 To UBound(m_byteText) + 2)
'
'    'WordMap(0).S = 0
'
'    For TL = 0 To UBound(m_byteText)
'        'MSL = Asc(Mid$(m_StrText, TL + 1, 1))
'
'        If m_byteText(TL) = 32 Or m_byteText(TL) = 13 Or m_byteText(TL) = 10 Or m_byteText(TL) = 45 Then ' a space
'            If WordMap(WC).L > 0 Then
''                WordMap(WC).H = WH
''                WordMap(WC).W = WW
''                WordMap(WC).L = WL
'                WC = WC + 1
''                WH = 0
''                WW = 0
''                WL = 0
'
'                WordMap(WC).S = TL + 1
'                MarkupS(TL).lPartOfWord = -1
'            End If
'        Else
'            MarkupS(TL).lPartOfWord = WC
'            If CharMap(TL).H > WordMap(WC).H Then
'                WordMap(WC).H = CharMap(TL).H
'            End If
'            WordMap(WC).W = WordMap(WC).W + CharMap(TL).W
'            WordMap(WC).L = WordMap(WC).L + 1
'
'        End If
'
'
'
'    Next TL
'
''    WordMap(WC).H = WH
''    WordMap(WC).W = WW
''    WordMap(WC).L = WL
'
'    'ReDim Preserve WordMap(0 To WC + 2)
'
'    WordCount = WC
'endff:
'End Sub


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
    
    'Redraw
End Sub


Function AddCharAtCursor(Optional sChar As String = "", Optional noevents As Boolean = False) As Boolean
    Dim lLength As Long
    Dim i As Long

    Dim lInsertLength As Long
    Dim lLengthDifference As Long

    Dim CursorToEnd As Long

    lInsertLength = Len(sChar)
    If lInsertLength = 0 And m_SelStart = m_SelEnd Then Exit Function

    If m_SelStart <> m_SelEnd Then
        lLengthDifference = lInsertLength - (m_SelEnd - m_SelStart)
    Else
        lLengthDifference = lInsertLength
    End If

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

    For i = 1 To lInsertLength
        m_byteText(m_SelStart + i - 1) = Asc(Mid$(sChar, i, 1))
        'If m_SelStart + i - 2 > 0 Then
        'CharMap(m_SelStart + i - 1) = CharMap(m_SelStart + i - 2)
        'End If
        'MarkupS(m_SelStart + i - 1) = MarkupS(m_SelStart + i - 2)
        'Else

        'End If
        If m_SelStart + i - 2 >= 0 Then
            MarkupS((m_SelStart + i - 1)) = MarkupS((m_SelStart + i - 2))
        Else
            With MarkupS(m_SelStart + i - 1)
                .lStrikeThrough = m_StdFont.Strikethrough
                .lFontSize = -1
                .lUnderline = m_StdFont.Underline
                .lItalic = m_StdFont.Italic
                .lBold = m_StdFont.Bold
                .lMarking = -1
                .lForeColor = -1
                .lLine = -1
            End With
        End If

    Next i

    CheckCharSize m_SelStart, lInsertLength

    m_SelStart = m_SelStart + lInsertLength
    m_SelEnd = m_SelStart
    m_CursorPos = m_SelStart

    'm_byteText(m_CursorPos) = Asc(sChar)
    'm_bMarkupCalculated = False
    m_bWordsCalculated = False

    'UserControl_KeyDown vbKeyRight, 0

    AddCharAtCursor = True
    If Not noevents Then RaiseEvent Changed
End Function


Sub CheckCharSize(lStart As Long, lLength As Long)
    Dim i As Long
    Dim uSize As Long
    
    Dim cForeColor As Long
    Dim cUnderline As Boolean
    Dim cItalic As Boolean
    Dim cBold As Boolean
    Dim cMarking As Long
    Dim cFontSize As Long
    Dim cStrikeThrough As Boolean
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
    cUnderline = m_StdFont.Underline
    cItalic = m_StdFont.Italic
    cBold = m_StdFont.Bold
    cFontSize = -1
    cMarking = -1
    cLine = -1
    cStrikeThrough = m_StdFont.Strikethrough

    GetTextMetrics UserControl.hdc, cTextMetric
    cDescendHeight = cTextMetric.tmDescent
    
    'uSize = UBound(MarkupS)
    
    For i = lStart To lStart + lLength
        With MarkupS(i)
            If .lFontSize <> cFontSize Then
                cFontSize = .lFontSize
                If .lFontSize = -1 Then
                    UserControl.FontSize = m_StdFont.Size
                Else
                    UserControl.FontSize = cFontSize
                End If

                GetTextMetrics UserControl.hdc, cTextMetric
                cDescendHeight = cTextMetric.tmDescent
            End If

            If .lBold <> cBold Then
                cBold = .lBold
                If cBold = -1 Then
                    UserControl.FontBold = m_StdFont.Bold
                Else
                    UserControl.FontBold = cBold
                End If
            End If

            If .lItalic <> cItalic Then
                cItalic = .lItalic
                If cItalic = -1 Then
                    UserControl.FontItalic = m_StdFont.Italic
                Else
                    UserControl.FontItalic = cItalic
                End If
            End If

            If .lUnderline <> cUnderline Then
                cUnderline = .lUnderline
                If .lUnderline = -1 Then
                    UserControl.FontUnderline = m_StdFont.Underline
                Else
                    UserControl.FontUnderline = cUnderline
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

Function getCharAtCursor(x As Long, y As Long) As Long
Dim i As Long

    Dim TTW     As Long    'total text width
    Dim CTW  As Long    'current text width
    Dim CR      As Long    'char row
    Dim UB      As Long    'ubound of rowmap
    Dim EOR As Long 'end of row
    Dim TS As Long 'total size

    UB = UBound(RowMap)
    CR = UB
    For i = 0 To UB    'number of rows
        If y < RowMap(i).StartY Then
            CR = i
            Exit For
        End If
    Next i

    If CR = UB Then
        EOR = UBound(CharMap)
    Else
        EOR = RowMap(CR + 1).StartChar - 1
    End If
    
    If CharMap(RowMap(CR).StartChar).x > x Then
        getCharAtCursor = RowMap(CR).StartChar
        Exit Function
    ElseIf CharMap(RowMap(CR).StartChar + RowMap(CR).NumChars - 1).x < x Then
        getCharAtCursor = EOR
        Exit Function
    End If
    
    'TS = CharMap(RowMap(CR).StartChar).X
    For i = RowMap(CR).StartChar To EOR
        If m_byteText(i) <> 10 And m_byteText(i) <> 13 Then
            If x > CharMap(i).x And x <= CharMap(i).x + CharMap(i).W Then
                If x < CharMap(i).x + CharMap(i).W / 2 Then
                    getCharAtCursor = i
                Else
                    getCharAtCursor = i + 1
                    If getCharAtCursor > UBound(m_byteText) Then getCharAtCursor = UBound(m_byteText)
                End If
                Exit Function
            End If
        End If

    Next i
End Function


Private Sub UserControl_LostFocus()
    DestroyCaret
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim tmpSwapSel As Long
    Dim mustRedraw As Boolean
    
    m_lMouseDown = m_lMouseDown Or Button
    m_lMouseDownX = x ' - m_lScrollLeft
    m_lMouseDownY = y ' - m_lScrollTop
    m_lMouseX = x ' - m_lScrollLeft
    m_lMouseY = y ' - m_lScrollTop
    
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
    
    updateCaretPos
    
    If Not m_bStarting And mustRedraw Then Redraw
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim tmpSwapSel As Long
    Dim mustRedraw As Boolean
    
    m_lMouseX = x ' - m_lScrollLeft
    m_lMouseY = y ' - m_lScrollTop
    
    If m_lMouseDown <> lNone Then
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
        
        If CharMap(m_CursorPos).x >= m_lUsercontrolWidth Then
            m_lScrollLeft = m_lScrollLeft + (CharMap(m_CursorPos).x - m_lUsercontrolWidth)
            mustRedraw = True
        ElseIf CharMap(m_CursorPos).x <= m_lUsercontrolLeft Then
            m_lScrollLeft = m_lScrollLeft + (CharMap(m_CursorPos).x - m_lUsercontrolLeft)
            mustRedraw = True
        End If
        
        If Not m_bStarting And mustRedraw Then Redraw
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_lMouseDown = m_lMouseDown And Not Button
    'Debug.Print m_lMouseDown
    If Not m_bStarting Then Redraw
End Sub


Function getNextCharUpDown(U As Boolean, STS As Boolean) As Long 'up, selectionTheSame
    Dim i As Long

    Dim TTW     As Long    'total text width
    Static CTW  As Long    'current text width
    Dim CR      As Long    'current word
    Dim UB      As Long    'ubound of rowmap
    
    UB = UBound(RowMap)
    CR = UB
    
    For i = 0 To UB    'number of rows
        If m_CursorPos < RowMap(i).StartChar Then
            CR = i - 1
            Exit For
        End If
    Next i
    
    
    If Not STS Then
        CTW = 0
        For i = RowMap(CR).StartChar To m_CursorPos - 1
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
    For i = RowMap(CR).StartChar To RowMap(CR).NumChars + RowMap(CR).StartChar
        TTW = TTW + CharMap(i).W
        If (TTW > CTW Or i = RowMap(CR).NumChars + RowMap(CR).StartChar) And m_byteText(i) <> 13 And m_byteText(i) <> 10 Then
            getNextCharUpDown = i
            Exit Function
        End If
    Next i

    getNextCharUpDown = RowMap(CR).StartChar

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
    
    getSelectionChanged True
    
    RaiseEvent KeyDown(KeyCode, Shift)
    
    If KeyCode = 0 And Shift = 0 Then m_bBlockNextKeyPress = True
    
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            If m_lMouseDown <> lNone Then Exit Sub
            m_CursorPos = getNextCharUpDown(KeyCode = vbKeyUp, m_SelUpDownTheSame)
            m_SelStart = m_CursorPos
            m_SelEnd = m_CursorPos
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

                        If KeyCode = vbKeyX Then
                            mustRedraw = AddCharAtCursor()
                        End If

                    Case vbKeyV
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
            If m_bMultiLine Then mustRedraw = AddCharAtCursor(vbCrLf)

        Case vbKeyBack
            If m_SelStart = m_SelEnd Then
                If m_SelStart > 0 Then
                    m_SelStart = m_SelStart - 1
                Else
                    Exit Sub
                End If
            End If

            mustRedraw = AddCharAtCursor()

        Case vbKeyDelete
            If m_SelEnd >= UBound(m_byteText) Then
                Exit Sub
            End If

            If m_SelStart = m_SelEnd Then
                m_SelEnd = m_SelEnd + 1
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
    

    If CharMap(m_CursorPos).x >= m_lUsercontrolWidth Then
        m_lScrollLeft = m_lScrollLeft + (CharMap(m_CursorPos).x - m_lUsercontrolWidth)
        mustRedraw = True
    ElseIf CharMap(m_CursorPos).x <= m_lUsercontrolLeft Then
        m_lScrollLeft = m_lScrollLeft + (CharMap(m_CursorPos).x - m_lUsercontrolLeft)
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

    WordPart = MarkupS(m_CursorPos).lPartOfWord
    If WordPart = -1 Then
        For i = m_CursorPos To UBound(CharMap)
            WordPart = MarkupS(i).lPartOfWord
            If WordPart <> -1 Then
                getNextWordFromCursor = WordMap(WordPart).s
                Exit Function
            End If
        Next i
        getNextWordFromCursor = UBound(CharMap)
    Else
        WordPart = WordPart + 1
        If WordPart > WordCount Then
            getNextWordFromCursor = WordMap(WordCount).s + WordMap(WordCount).L
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

    WordPart = MarkupS(m_CursorPos).lPartOfWord
    If WordPart = -1 Then
        For i = m_CursorPos To 0 Step -1
            WordPart = MarkupS(i).lPartOfWord
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


Private Sub Usercontrol_Resize()
    If Not m_bStarting Then Redraw

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
    m_SelEnd = m_SelStart + WordMap(wordNr).L
    
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
    Dim cUnderline As Boolean
    Dim cItalic As Boolean
    Dim cBold As Boolean
    Dim cMarking As Long
    Dim cFontSize As Long
    Dim cStrikeThrough As Boolean
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



Function Fmat(str As String, Length As Long) As String
    Dim strLength As Long
    strLength = Len(str)

    If strLength > Length Then
        Fmat = String(Length, "X")
    ElseIf strLength < Length Then
        Fmat = String(Length - strLength, "0") & str
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
        .WriteProperty "MousePointer", m_MouMousePointer, 0
        .WriteProperty "Border", m_bBorder, True
        .WriteProperty "LineNumbers", m_bLineNumbers, False

        .WriteProperty "RowLines", m_bRowLines, False
        .WriteProperty "RowLineColor", m_OleRowLineColor, &HEEEEEE
        .WriteProperty "RowNumberOnEveryLine", m_bRowNumberOnEveryLine, False
        .WriteProperty "WordWrap", m_bWordWrap, False
        .WriteProperty "MultiLine", m_bMultiLine, False
        .WriteProperty "HideCursor", m_bHideCursor, False
        .WriteProperty "AutoResize", m_bAutoResize, False
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
        MousePointer = .ReadProperty("MousePointer", 0)
        m_bBorder = .ReadProperty("Border", True)
        m_bLineNumbers = .ReadProperty("LineNumbers", False)

        m_bRowLines = .ReadProperty("RowLines", False)
        m_OleRowLineColor = .ReadProperty("RowLineColor", &HEEEEEE)
        m_bRowNumberOnEveryLine = .ReadProperty("RowNumberOnEveryLine", False)
        m_bWordWrap = .ReadProperty("WordWrap", False)
        m_bMultiLine = .ReadProperty("MultiLine", False)
        m_bHideCursor = .ReadProperty("HideCursor", False)
        m_bAutoResize = .ReadProperty("AutoResize", False)
        m_sScrollBars = .ReadProperty("ScrollBars", ScrollBarStyle.lNone)
        
    End With
    m_bStarting = False
    Redraw
End Sub






