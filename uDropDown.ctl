VERSION 5.00
Begin VB.UserControl uDropDown 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   FillStyle       =   0  'Solid
   MousePointer    =   1  'Arrow
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
End
Attribute VB_Name = "uDropDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function CreatePen& Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long)
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long


Private Declare Function GetLastError Lib "kernel32" () As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const GWL_EXSTYLE As Long = (-20)
Private Const SW_HIDE As Long = 0
Private Const SW_SHOW As Long = 5


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Todo:
'ScrollbarHeight instellen zodat je niet zo'n hele kleine scrollbar krijgt
'Scroll Interval speed setting
'scroll Speed setting
'settingform:
'   - must be able to add and remove items form the shotcutlist
'   - only show the comport box when arduino is selected (checkbox in the setting?)
'   - add and remove games from the game tab
'   - make an info label that has all the commands in it like: {TEXT} and {COMM} and {PATH32} and {PATH2560}
'   - be able to create shortcuts as a system hook to execute some commands without the need of typing a or b or y or g.
'   - Add setting to set the hight of the form when in the taskbar and when in the top of the screen
'
'dropdownbox:
'   - when selected option has one or more vbcrlf then only show only first row if it does not fit
'   - check if touchscreen is able to click everything nicely
'   -

'
'
'
'


Private Type Item
    Text As String
    ItemData As Long
    ItemColor As OLE_COLOR
    TextAlignment As AlignmentConstants
End Type


Public Event ItemChange(ItemIndex As Long)
Public Event OnDropdown()

Private WithEvents m_picMenu As PictureBox
Attribute m_picMenu.VB_VarHelpID = -1
Private WithEvents m_tmrFocus As Timer
Attribute m_tmrFocus.VB_VarHelpID = -1
Private WithEvents m_tmrScroll As Timer
Attribute m_tmrScroll.VB_VarHelpID = -1

Private m_OleBackgroundColor As OLE_COLOR
Private m_OleForeColor As OLE_COLOR
Private m_OleBorderColor As OLE_COLOR
Private m_OleSelectionBorderColor As OLE_COLOR
Private m_OleSelectionBackgroundColor As OLE_COLOR

Private m_OleBackgroundColorDisabled As OLE_COLOR
Private m_OleForeColorDisabled As OLE_COLOR
Private m_OleBorderColorDisabled As OLE_COLOR
Private m_OleSelectionBorderColorDisabled As OLE_COLOR
Private m_OleSelectionBackgroundColorDisabled As OLE_COLOR

Private m_StrText As String
Private m_bBorder As Boolean
Private m_bStarting As Boolean
Private m_MouMousePointer As MousePointerConstants
Private m_bRefreshing As Boolean
Private m_bPicButtonDown As Boolean
Private m_bMenuDown As Boolean
Private m_bGotFocus As Boolean

Private m_StdStandardFont As New StdFont
Private m_PoiMouse As POINTAPI
Private m_bEnabled As Boolean

Private m_bRefreshingMenu As Boolean
Private m_PoiMenuMouse As POINTAPI
Private m_LonItemHeight As Long
Private m_LonItemCount As Long
Private m_LonItemAtTop As Long
Private m_LonItemsVisible As Long

Private m_LonScrollBarWidth As Long
Private m_bScrollArrowDown As Boolean
Private m_bScrollArrowUp As Boolean
Private m_LonScrollTop As Long
Private m_LonScrollHeight As Long
Private m_LonScrollMax As Long
Private m_bScrollHandleVisible As Boolean
Private m_bScrollHandleDown As Boolean
Private m_LonScrollHandleDragY As Long

Private Items() As Item

Private m_LonListIndex As Long
Private m_LonListIndexMouseOver As Long
Private m_StdFont As StdFont
Private m_StdFontWebdings As StdFont
Private m_LonDotsTextWidth As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Function ShortenText(ByRef StrValue As String, LonLength As Long) As String
    Dim tmpStrPrint As String
    Dim j As Long

    tmpStrPrint = "..."
    If LonLength < m_LonDotsTextWidth Then
        ShortenText = ""
        Exit Function
    End If

    If m_picMenu.TextWidth(StrValue) > LonLength Then
        For j = 1 To Len(StrValue)
            tmpStrPrint = Mid$(StrValue, 1, j)
            If m_picMenu.TextWidth(tmpStrPrint) + m_LonDotsTextWidth > LonLength Then
                tmpStrPrint = Mid$(StrValue, 1, j - 1) & "..."
                Exit For
            End If

        Next j
        ShortenText = tmpStrPrint
    Else
        ShortenText = StrValue
    End If
End Function


Sub Clear()
    m_LonItemCount = 0
    m_LonListIndex = -1
    m_LonListIndexMouseOver = -1
    m_StrText = ""
    ReDim Items(0 To 0) As Item

    If Not m_bStarting Then Redraw
End Sub

Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property

Public Property Let Enabled(ByVal bValue As Boolean)
    m_bEnabled = bValue
    PropertyChanged "Enabled"
    If Not m_bStarting Then Redraw
End Property

Public Property Get Font() As StdFont
    Set Font = m_StdFont
End Property

Public Property Set Font(ByVal StdValue As StdFont)
    Set m_StdFont = StdValue
    Set UserControl.Font = m_StdFont
    Set m_picMenu.Font = m_StdFont
    PropertyChanged "Font"
    m_LonDotsTextWidth = m_picMenu.TextWidth("...")
    If Not m_bStarting Then Redraw
End Property


Public Property Get ListCount() As Long
    ListCount = m_LonItemCount
End Property


Public Property Get ListIndex() As Long
    ListIndex = m_LonListIndex
End Property

Public Property Let ListIndex(Index As Long)
    If Index < 0 Or Index > m_LonItemCount - 1 Then Exit Property 'Err.Raise 19444, "", "Array Out of Bound": Exit Property
    
    m_LonListIndex = Index
    m_StrText = Items(m_LonListIndex).Text
    RaiseEvent ItemChange(m_LonListIndex)
    If Not m_bStarting Then Redraw
End Property


Public Property Get List(Index As Long) As String
    If Index < 0 Or Index > m_LonItemCount - 1 Then Err.Raise 19444, "", "Array Out of Bound": Exit Property
    List = Items(Index).Text
End Property

Public Property Let List(Index As Long, ByVal StrValue As String)
    If Index < 0 Or Index > m_LonItemCount - 1 Then Err.Raise 19444, "", "Array Out of Bound": Exit Property
    Items(Index).Text = StrValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get ItemData(Index As Long) As Long
    If Index < 0 Or Index > m_LonItemCount - 1 Then Err.Raise 19444, "", "Array Out of Bound": Exit Property
    ItemData = Items(Index).ItemData
End Property

Public Property Let ItemData(Index As Long, ByVal LonValue As Long)
    If Index < 0 Or Index > m_LonItemCount - 1 Then Err.Raise 19444, "", "Array Out of Bound": Exit Property
    Items(Index).ItemData = LonValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get ItemAlignment(Index As Long) As Long
    If Index < 0 Or Index > m_LonItemCount - 1 Then Err.Raise 19444, "", "Array Out of Bound": Exit Property
    ItemAlignment = Items(Index).TextAlignment
End Property

Public Property Let ItemAlignment(Index As Long, ByVal AliValue As AlignmentConstants)
    If Index < 0 Or Index > m_LonItemCount - 1 Then Err.Raise 19444, "", "Array Out of Bound": Exit Property
    Items(Index).TextAlignment = AliValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get ItemColor(Index As Long) As OLE_COLOR
    If Index < 0 Or Index > m_LonItemCount - 1 Then Err.Raise 19444, "", "Array Out of Bound": Exit Property
    ItemColor = Items(Index).ItemColor
End Property

Public Property Let ItemColor(Index As Long, ByVal OleValue As OLE_COLOR)
    If Index < 0 Or Index > m_LonItemCount - 1 Then Err.Raise 19444, "", "Array Out of Bound": Exit Property
    Items(Index).ItemColor = OleValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get ItemsVisible() As Long
    ItemsVisible = m_LonItemsVisible
End Property

Public Property Let ItemsVisible(ByVal LonValue As Long)
    m_LonItemsVisible = LonValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get ScrollBarWidth() As Long
    ScrollBarWidth = m_LonScrollBarWidth
End Property

Public Property Let ScrollBarWidth(ByVal LonValue As Long)
    m_LonScrollBarWidth = LonValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get ItemHeight() As Long
    ItemHeight = m_LonItemHeight
End Property

Public Property Let ItemHeight(ByVal LonValue As Long)
    m_LonItemHeight = LonValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get MousePointer() As MousePointerConstants
    MousePointer = m_MouMousePointer
End Property

Public Property Let MousePointer(ByVal MouValue As MousePointerConstants)
    m_MouMousePointer = MouValue
End Property


Public Property Get Border() As Boolean
    Border = m_bBorder
End Property

Public Property Let Border(ByVal bValue As Boolean)
    m_bBorder = bValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get Text() As String
    Text = m_StrText
End Property

Public Property Let Text(ByVal StrValue As String)
    m_StrText = StrValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get SelectionBackgroundColor() As OLE_COLOR
    SelectionBackgroundColor = m_OleSelectionBackgroundColor
End Property

Public Property Let SelectionBackgroundColor(ByVal OleValue As OLE_COLOR)
    m_OleSelectionBackgroundColor = OleValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get SelectionBorderColor() As OLE_COLOR
    SelectionBorderColor = m_OleSelectionBorderColor
End Property

Public Property Let SelectionBorderColor(ByVal OleValue As OLE_COLOR)
    m_OleSelectionBorderColor = OleValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_OleBorderColor
End Property

Public Property Let BorderColor(ByVal OleValue As OLE_COLOR)
    m_OleBorderColor = OleValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_OleForeColor
End Property

Public Property Let ForeColor(ByVal OleValue As OLE_COLOR)
    m_OleForeColor = OleValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get BackgroundColor() As OLE_COLOR
    BackgroundColor = m_OleBackgroundColor
End Property

Public Property Let BackgroundColor(ByVal OleValue As OLE_COLOR)
    m_OleBackgroundColor = OleValue
    If Not m_bStarting Then Redraw
End Property










Public Property Get SelectionBackgroundColorDisabled() As OLE_COLOR
    SelectionBackgroundColorDisabled = m_OleSelectionBackgroundColorDisabled
End Property

Public Property Let SelectionBackgroundColorDisabled(ByVal OleValue As OLE_COLOR)
    m_OleSelectionBackgroundColorDisabled = OleValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get SelectionBorderColorDisabled() As OLE_COLOR
    SelectionBorderColorDisabled = m_OleSelectionBorderColorDisabled
End Property

Public Property Let SelectionBorderColorDisabled(ByVal OleValue As OLE_COLOR)
    m_OleSelectionBorderColorDisabled = OleValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get BorderColorDisabled() As OLE_COLOR
    BorderColorDisabled = m_OleBorderColorDisabled
End Property

Public Property Let BorderColorDisabled(ByVal OleValue As OLE_COLOR)
    m_OleBorderColorDisabled = OleValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get ForeColorDisabled() As OLE_COLOR
    ForeColorDisabled = m_OleForeColorDisabled
End Property

Public Property Let ForeColorDisabled(ByVal OleValue As OLE_COLOR)
    m_OleForeColorDisabled = OleValue
    If Not m_bStarting Then Redraw
End Property


Public Property Get BackgroundColorDisabled() As OLE_COLOR
    BackgroundColorDisabled = m_OleBackgroundColorDisabled
End Property

Public Property Let BackgroundColorDisabled(ByVal OleValue As OLE_COLOR)
    m_OleBackgroundColorDisabled = OleValue
    If Not m_bStarting Then Redraw
End Property










Sub RedrawPause()
    m_bStarting = True
End Sub

Sub RedrawResume()
    m_bStarting = False
    Redraw
End Sub

Public Function AddItem(sText As String, Optional lItemData As Long = 0, Optional Index As Long = -1, Optional lItemColor As OLE_COLOR = -1, Optional lAlignment As AlignmentConstants = vbLeftJustify) As Long

    If Index = -1 Then
        ReDim Preserve Items(0 To m_LonItemCount) As Item
        With Items(m_LonItemCount)
            .Text = sText
            .ItemData = lItemData
            .ItemColor = lItemColor
            .TextAlignment = lAlignment
        End With

        AddItem = m_LonItemCount
    Else

        ' We let VB evaluate the size of each item using LenB().
        ReDim Preserve Items(0 To m_LonItemCount) As Item
        If Index < UBound(Items) Then
            CopyMemory ByVal VarPtr(Items(Index + 1)), ByVal VarPtr(Items(Index)), (UBound(Items) - Index) * LenB(Items(Index))
            With Items(Index)
                .Text = sText
                .ItemData = lItemData
                .ItemColor = lItemColor
                .TextAlignment = lAlignment
            End With
        End If

    End If

    m_LonItemCount = m_LonItemCount + 1

End Function

Public Sub RemoveItem(Index As Long)

    If Index < 0 Or Index >= m_LonItemCount Then Err.Raise 19444, "", "Array Out of Bound": Exit Sub

    ' We let VB evaluate the size of each item using LenB().
    'm_LonItemCount = m_LonItemCount - 1

    If Index < m_LonItemCount - 1 Then
        CopyMemory ByVal VarPtr(Items(Index)), ByVal VarPtr(Items(Index + 1)), (UBound(Items) - Index) * LenB(Items(Index)) + 1

        ReDim Preserve Items(0 To UBound(Items) - 1)
        m_LonItemCount = m_LonItemCount - 1
    Else
        ReDim Preserve Items(0 To Index) As Item
    End If


End Sub



Sub CheckScrollButtons(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim tmpL As Long

    m_PoiMenuMouse.x = x
    m_PoiMenuMouse.y = y

    m_bScrollArrowUp = False
    m_bScrollArrowDown = False

    If Button = 1 And m_bScrollHandleVisible Then
        If x >= m_picMenu.ScaleWidth - m_LonScrollBarWidth + 2 And x <= m_picMenu.ScaleWidth - 3 Then
            'mouse is above scrollbar

            If y >= 2 And y < m_LonScrollBarWidth - 2 Then
                m_bScrollArrowUp = True
            End If

            If y >= m_picMenu.ScaleHeight - m_LonScrollBarWidth + 3 And y < m_picMenu.ScaleHeight - 1 Then
                m_bScrollArrowDown = True
            End If

            tmpL = m_LonScrollTop + ((m_LonScrollMax - m_LonScrollHeight) / (m_LonItemCount - m_LonItemsVisible) * m_LonItemAtTop)
            If y >= tmpL And y < tmpL + m_LonScrollHeight Then
                m_bScrollHandleDown = True

            End If

        End If
    End If
End Sub

Private Sub m_picMenu_DblClick()
    m_picMenu_MouseDown 1, 0, CInt(m_PoiMenuMouse.x), CInt(m_PoiMenuMouse.y)
End Sub

Private Sub m_picMenu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    CheckScrollButtons Button, Shift, x, y
    If m_bScrollArrowUp Or m_bScrollArrowDown Then
        m_tmrScroll.Interval = 500
        m_tmrScroll.Enabled = True
    End If

    If m_bScrollHandleDown Then
        m_LonScrollHandleDragY = y - (m_LonScrollTop + ((m_LonScrollMax - m_LonScrollHeight) / (m_LonItemCount - m_LonItemsVisible) * m_LonItemAtTop))
    End If

    RedrawMenu
End Sub

Private Sub m_picMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim tmpL As Long

    CheckScrollButtons Button, Shift, x, y

    If m_bScrollHandleDown Then
        tmpL = (y - m_LonScrollHandleDragY - m_LonScrollTop) / ((m_LonScrollMax - m_LonScrollHeight) / (m_LonItemCount - m_LonItemsVisible))
        If tmpL < 0 Then tmpL = 0
        If tmpL > m_LonItemCount - m_LonItemsVisible Then tmpL = m_LonItemCount - m_LonItemsVisible

        m_LonItemAtTop = tmpL
    End If

    RedrawMenu
End Sub

Private Sub m_picMenu_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    CheckScrollButtons Button, Shift, x, y
    If m_bScrollArrowUp Then
        m_LonItemAtTop = m_LonItemAtTop - 1
        If m_LonItemAtTop < 0 Then m_LonItemAtTop = 0
    ElseIf m_bScrollArrowDown Then
        m_LonItemAtTop = m_LonItemAtTop + 1
        If m_LonItemAtTop > m_LonItemCount - m_LonItemsVisible Then m_LonItemAtTop = m_LonItemCount - m_LonItemsVisible
    ElseIf m_LonListIndexMouseOver <> -1 And m_bScrollHandleDown = False Then
        m_LonListIndex = m_LonListIndexMouseOver
        m_StrText = Items(m_LonListIndex).Text
        RaiseEvent ItemChange(m_LonListIndex)
        CloseMenu
        Redraw
    End If

    m_tmrScroll.Enabled = False
    m_bScrollArrowUp = False
    m_bScrollArrowDown = False
    m_bScrollHandleDown = False
    RedrawMenu
End Sub

Private Sub m_tmrFocus_Timer()
    On Error GoTo Not_Supported
    
    Dim i As Object
    Dim j As Long
    
    Set i = UserControl.Parent

    'If LCase(TypeName(i)) = "form" Then
        If GetForegroundWindow() <> i.hWnd Then
            CloseMenu
        End If
    'End If
    
    'Debug.Print "GetForegroundWindow():" & GetForegroundWindow(); "i.hwnd: " & i.hwnd

    Exit Sub
Not_Supported:
    
    
End Sub


Private Sub m_tmrScroll_Timer()
    m_tmrScroll.Interval = 1
    If m_bScrollArrowUp Then
        m_LonItemAtTop = m_LonItemAtTop - 1
        If m_LonItemAtTop < 0 Then m_LonItemAtTop = 0
    ElseIf m_bScrollArrowDown Then
        m_LonItemAtTop = m_LonItemAtTop + 1
        If m_LonItemAtTop > m_LonItemCount - m_LonItemsVisible Then m_LonItemAtTop = m_LonItemCount - m_LonItemsVisible
    End If
    Redraw
End Sub

Private Sub UserControl_DblClick()
    UserControl_MouseDown 1, 0, CInt(m_PoiMouse.x), CInt(m_PoiMouse.y)
End Sub

Private Sub UserControl_EnterFocus()
    m_bGotFocus = True
    Redraw
End Sub

Private Sub UserControl_ExitFocus()
    m_bGotFocus = False
    m_bMenuDown = False
    CloseMenu
    Redraw
End Sub

Private Sub UserControl_Initialize()
    m_bStarting = True
    
    m_OleBorderColor = &HFFFFFF
    m_OleBackgroundColor = &HFFFFFF
    m_OleForeColor = &H0
    m_OleSelectionBackgroundColor = &HFFFFFF
    m_OleSelectionBorderColor = &HFFFFFF
    
    m_OleBorderColorDisabled = &HFFFFFF
    m_OleBackgroundColorDisabled = &HFFFFFF
    m_OleForeColorDisabled = &H0
    m_OleSelectionBackgroundColorDisabled = &HFFFFFF
    m_OleSelectionBorderColorDisabled = &HFFFFFF
    
    m_StrText = "uFrame"
    m_bBorder = True
    m_StdStandardFont = UserControl.Font
    m_MouMousePointer = 0
    m_LonItemsVisible = 5
    m_LonListIndex = -1
    m_LonScrollBarWidth = 20
    m_bEnabled = True
    
    Set m_picMenu = UserControl.Controls.Add("VB.PictureBox", "m_picMenu")
    Set m_tmrFocus = UserControl.Controls.Add("VB.Timer", "m_tmrFocus")
    Set m_tmrScroll = UserControl.Controls.Add("VB.Timer", "m_tmrScroll")

    Usercontrol_Resize
    m_picMenu.BorderStyle = 0
    m_picMenu.AutoRedraw = True
    m_picMenu.ScaleMode = vbPixels
    m_picMenu.Visible = False
    m_picMenu.FontSize = 8
    
    Set m_StdFontWebdings = New StdFont
    m_StdFontWebdings.name = "Webdings"
    m_StdFontWebdings.Size = 8
    
    SetParent m_picMenu.hWnd, GetParent(0)

    SetWindowLong m_picMenu.hWnd, -20, GetWindowLong(m_picMenu.hWnd, -20) Or &H80&



    'Call SetWindowLong(m_picMenu.hwnd, GWL_EXSTYLE, GetWindowLong(m_picMenu.hwnd, GWL_EXSTYLE) - WS_EX_APPWINDOW)



    m_LonItemCount = 0

    '
    '    ReDim Items(0 To 5) As Item
    '    m_LonItemCount = 6
    '
    '    Items(0).ItemData = 70
    '    Items(0).Text = "1. dp"
    '    Items(0).TextAlignment = vbLeftJustify
    '
    '    Items(1).ItemData = 60
    '    Items(1).Text = "2. kp"
    '    Items(1).TextAlignment = vbCenter
    '
    '    Items(2).ItemData = 40
    '    Items(2).Text = "3."
    '    Items(2).TextAlignment = vbRightJustify
    '
    '    Items(3).ItemData = 90
    '    Items(3).Text = "4. kijken hoeveel ruimte ik heb voor de karakters"
    '    Items(3).TextAlignment = vbCenter
    '
    '    Items(4).ItemData = 40
    '    Items(4).Text = "5. een appel"
    '
    '    Items(5).ItemData = 40
    '    Items(5).Text = "6. en een peer"
    '
End Sub



Private Function CalculatePosition() As POINTAPI
    Dim tmpMenuPosition As RECT

    GetWindowRect UserControl.hWnd, tmpMenuPosition

    CalculatePosition.x = tmpMenuPosition.Left * Screen.TwipsPerPixelX
    CalculatePosition.y = tmpMenuPosition.Top * Screen.TwipsPerPixelY

End Function

Sub Redraw()
    Dim tmpTextHeight As Long
    Dim tmpTextWidth As Long

    Dim tmpPicText As String
    Dim tmpPicTextOffset As Long

    If m_bRefreshing Then Exit Sub
    m_bRefreshing = True

    tmpTextHeight = UserControl.TextHeight(m_StrText)
    tmpTextWidth = UserControl.TextWidth(m_StrText)

    UserControl.Cls

    UserControl.BackColor = IIf(m_bEnabled, m_OleBackgroundColor, m_OleBackgroundColorDisabled)

    UserControl.DrawStyle = 0
    
    
    If m_bBorder Then
        UserControl.Line (0, 0)-(0, UserControl.ScaleHeight), IIf(m_bEnabled, m_OleBorderColor, m_OleBorderColorDisabled)
        UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), IIf(m_bEnabled, m_OleBorderColor, m_OleBorderColorDisabled)
        UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), IIf(m_bEnabled, m_OleBorderColor, m_OleBorderColorDisabled)
        UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), IIf(m_bEnabled, m_OleBorderColor, m_OleBorderColorDisabled)

        UserControl.Line (UserControl.ScaleWidth - m_LonScrollBarWidth, 0)-(UserControl.ScaleWidth - m_LonScrollBarWidth, UserControl.ScaleHeight - 1), IIf(m_bEnabled, m_OleBorderColor, m_OleBorderColorDisabled)
    End If

    
    
    
    ReDim pts(0 To 3)
    pts(0).x = 3: pts(0).y = 3
    pts(1).x = 3: pts(1).y = UserControl.ScaleHeight - 3
    pts(2).x = UserControl.ScaleWidth - m_LonScrollBarWidth - 2: pts(2).y = UserControl.ScaleHeight - 3
    pts(3).x = UserControl.ScaleWidth - m_LonScrollBarWidth - 2: pts(3).y = 3
    UserControl.ForeColor = IIf(m_bEnabled, m_OleSelectionBorderColor, m_OleSelectionBorderColorDisabled)
    If m_LonListIndex > -1 Then
        UserControl.FillColor = IIf(Items(m_LonListIndex).ItemColor <> -1, Items(m_LonListIndex).ItemColor, IIf(m_bEnabled, m_OleSelectionBackgroundColor, m_OleSelectionBackgroundColorDisabled))
    Else
        UserControl.FillColor = IIf(m_bEnabled, m_OleSelectionBackgroundColor, m_OleSelectionBackgroundColorDisabled)
    End If
    UserControl.DrawStyle = 5
    UserControl.FillStyle = 0
    Polygon UserControl.hdc, pts(0), 4


    If m_bGotFocus And Not m_bMenuDown And m_bEnabled Then
        pts(0).x = pts(0).x - 1
        pts(0).y = pts(0).y - 1
        pts(1).x = pts(1).x - 1
        pts(3).y = pts(3).y - 1
        UserControl.DrawStyle = 2
        UserControl.FillStyle = 1
        UserControl.ForeColor = IIf(m_bEnabled, m_OleSelectionBorderColor, m_OleSelectionBorderColorDisabled)
        Polygon UserControl.hdc, pts(0), 4
    End If


    If m_bPicButtonDown Then
        tmpPicTextOffset = 1
        UserControl.ForeColor = IIf(m_bEnabled, m_OleBackgroundColor, m_OleBackgroundColorDisabled)
    Else
        tmpPicTextOffset = 0
        UserControl.ForeColor = IIf(m_bEnabled, m_OleBorderColor, m_OleBorderColorDisabled)
    End If
    UserControl.FillColor = IIf(m_bEnabled, m_OleBackgroundColor, m_OleBackgroundColorDisabled)


    ReDim pts(0 To 3)
    UserControl.DrawStyle = 0
    pts(0).x = UserControl.ScaleWidth - m_LonScrollBarWidth + 2: pts(0).y = 2
    pts(1).x = UserControl.ScaleWidth - m_LonScrollBarWidth + 2: pts(1).y = UserControl.ScaleHeight - 3
    pts(2).x = UserControl.ScaleWidth - 3: pts(2).y = UserControl.ScaleHeight - 3
    pts(3).x = UserControl.ScaleWidth - 3: pts(3).y = 2
    Polygon UserControl.hdc, pts(0), 4


    If m_bMenuDown Then
        tmpPicText = "5"
    Else
        tmpPicText = "6"
    End If

    UserControl.ForeColor = IIf(m_bEnabled, m_OleForeColor, m_OleForeColorDisabled)
    UserControl.DrawStyle = 0

    Set UserControl.Font = m_StdFontWebdings
    UserControl.CurrentX = Fix(UserControl.ScaleWidth - m_LonScrollBarWidth / 2 - UserControl.TextWidth(tmpPicText) / 2) + 1 + tmpPicTextOffset
    UserControl.CurrentY = Fix(UserControl.ScaleHeight / 2 - UserControl.TextHeight(tmpPicText) / 2) + tmpPicTextOffset - 1
    UserControl.Print tmpPicText

    Set UserControl.Font = m_StdFont
    UserControl.CurrentX = 7
    UserControl.CurrentY = Fix(UserControl.ScaleHeight / 2 - tmpTextHeight / 2)
    UserControl.Print ShortenText(m_StrText, UserControl.ScaleWidth - UserControl.ScaleHeight - 9)

    If m_bMenuDown Then
        RedrawMenu
    End If

    m_bRefreshing = False
End Sub

Sub RedrawMenu()
    Dim i As Long

    Dim tmpPrintTop As Long
    Dim tmpPrintRow As Long
    Dim tmpArrowUpOffset As Long
    Dim tmpArrowDownOffset As Long
    Dim tmpSplit() As String
    Dim tmpTextHeight As Long
    Dim tmpSplitLength As Long
    Dim tmpShortText As String
    Dim tmpLeft As Long

    If m_bRefreshingMenu Then Exit Sub
    m_bRefreshingMenu = True

    m_picMenu.Width = UserControl.Width

    ReDim pts(0 To 3)
    m_picMenu.DrawStyle = 0
    m_picMenu.FillStyle = 0
    pts(0).x = 0: pts(0).y = 0
    pts(1).x = m_picMenu.ScaleWidth - 1: pts(1).y = 0
    pts(2).x = m_picMenu.ScaleWidth - 1: pts(2).y = m_picMenu.ScaleHeight - 1
    pts(3).x = 0: pts(3).y = m_picMenu.ScaleHeight - 1
    
    
    m_picMenu.FillColor = IIf(m_bEnabled, m_OleBackgroundColor, m_OleBackgroundColorDisabled)
    m_picMenu.ForeColor = IIf(m_bEnabled, m_OleBorderColor, m_OleBorderColorDisabled)
    Polygon m_picMenu.hdc, pts(0), 4
    m_picMenu.ForeColor = IIf(m_bEnabled, m_OleForeColor, m_OleForeColorDisabled)

    tmpPrintTop = 1
    m_LonListIndexMouseOver = -1
    Dim tmpScrollbarWidth As Long

    If m_bScrollHandleVisible Then
        tmpScrollbarWidth = m_LonScrollBarWidth
    Else
        tmpScrollbarWidth = 1
    End If

    While tmpPrintTop < m_picMenu.ScaleHeight - m_LonItemHeight And i < m_LonItemCount And (i + m_LonItemAtTop) < m_LonItemCount

        pts(0).x = 2: pts(0).y = tmpPrintTop + 1
        pts(1).x = m_picMenu.ScaleWidth - tmpScrollbarWidth - 2: pts(1).y = tmpPrintTop + 1
        pts(2).x = m_picMenu.ScaleWidth - tmpScrollbarWidth - 2: pts(2).y = tmpPrintTop + m_LonItemHeight
        pts(3).x = 2: pts(3).y = tmpPrintTop + m_LonItemHeight

        If m_PoiMenuMouse.x >= 0 And m_PoiMenuMouse.x <= m_picMenu.ScaleWidth - tmpScrollbarWidth And _
           m_PoiMenuMouse.y > tmpPrintTop And m_PoiMenuMouse.y < tmpPrintTop + m_LonItemHeight + 1 Then
            m_LonListIndexMouseOver = m_LonItemAtTop + i
        Else
            m_picMenu.ForeColor = IIf(m_bEnabled, m_OleBackgroundColor, m_OleBackgroundColorDisabled)
        End If

        m_picMenu.FillColor = IIf(Items(m_LonItemAtTop + i).ItemColor = -1, IIf(m_bEnabled, m_OleSelectionBackgroundColor, m_OleSelectionBackgroundColorDisabled), Items(m_LonItemAtTop + i).ItemColor)
        Polygon m_picMenu.hdc, pts(0), 4

        m_picMenu.ForeColor = IIf(m_bEnabled, m_OleForeColor, m_OleForeColorDisabled)

        tmpTextHeight = m_picMenu.TextHeight(Items(m_LonItemAtTop + i).Text)
        tmpSplit = Split(Items(m_LonItemAtTop + i).Text, vbCrLf)





        If tmpTextHeight < m_LonItemHeight Then
            For tmpSplitLength = 0 To UBound(tmpSplit)
                tmpShortText = ShortenText(tmpSplit(tmpSplitLength), m_picMenu.ScaleWidth - m_LonScrollBarWidth - 5)

                Select Case Items(m_LonItemAtTop + i).TextAlignment
                    Case AlignmentConstants.vbLeftJustify
                        tmpLeft = 3

                    Case AlignmentConstants.vbCenter
                        tmpLeft = (m_picMenu.ScaleWidth - m_LonScrollBarWidth - 5) / 2 - m_picMenu.TextWidth(tmpShortText) / 2 + 2

                    Case AlignmentConstants.vbRightJustify
                        tmpLeft = (m_picMenu.ScaleWidth - m_LonScrollBarWidth - 2) - m_picMenu.TextWidth(tmpShortText)
                End Select

                m_picMenu.CurrentX = tmpLeft
                m_picMenu.CurrentY = tmpPrintTop + m_LonItemHeight / 2 + (tmpTextHeight / (UBound(tmpSplit) + 1) * (tmpSplitLength - ((UBound(tmpSplit) + 1) / 2))) + 1

                m_picMenu.Print tmpShortText
            Next tmpSplitLength
        Else

            tmpShortText = ShortenText(tmpSplit(0), m_picMenu.ScaleWidth - m_LonScrollBarWidth - 5)

            Select Case Items(m_LonItemAtTop + i).TextAlignment
                Case AlignmentConstants.vbLeftJustify
                    tmpLeft = 3

                Case AlignmentConstants.vbCenter
                    tmpLeft = (m_picMenu.ScaleWidth - m_LonScrollBarWidth - 5) / 2 - m_picMenu.TextWidth(tmpShortText) / 2 + 3

                Case AlignmentConstants.vbRightJustify
                    tmpLeft = (m_picMenu.ScaleWidth - m_LonScrollBarWidth - 2) - m_picMenu.TextWidth(tmpShortText)

            End Select

            m_picMenu.CurrentX = tmpLeft
            m_picMenu.CurrentY = tmpPrintTop + m_LonItemHeight / 2 - m_picMenu.TextHeight(tmpSplit(0)) / 2

            m_picMenu.Print tmpShortText
        End If

        i = i + 1
        tmpPrintTop = tmpPrintTop + m_LonItemHeight

    Wend
    
    If m_bScrollHandleVisible Then
        'The ScrollBar
        m_picMenu.DrawStyle = 0
        pts(0).x = m_picMenu.ScaleWidth - m_LonScrollBarWidth: pts(0).y = 0
        pts(1).x = m_picMenu.ScaleWidth - 1: pts(1).y = 0
        pts(2).x = m_picMenu.ScaleWidth - 1: pts(2).y = m_picMenu.ScaleHeight - 1
        pts(3).x = m_picMenu.ScaleWidth - m_LonScrollBarWidth: pts(3).y = m_picMenu.ScaleHeight - 1
        m_picMenu.ForeColor = IIf(m_bEnabled, m_OleBorderColor, m_OleBorderColorDisabled)
        m_picMenu.FillColor = IIf(m_bEnabled, m_OleBackgroundColor, m_OleBackgroundColorDisabled)
        Polygon m_picMenu.hdc, pts(0), 4

        'Upper Arrow+Button
        If Not m_bScrollArrowUp Then
            tmpArrowUpOffset = 0
            m_picMenu.DrawStyle = 0
        Else
            tmpArrowUpOffset = 1
            m_picMenu.DrawStyle = 2
        End If
        pts(0).x = m_picMenu.ScaleWidth - m_LonScrollBarWidth + 2: pts(0).y = 2
        pts(1).x = m_picMenu.ScaleWidth - 3: pts(1).y = 2
        pts(2).x = m_picMenu.ScaleWidth - 3: pts(2).y = m_LonScrollBarWidth - 3
        pts(3).x = m_picMenu.ScaleWidth - m_LonScrollBarWidth + 2: pts(3).y = m_LonScrollBarWidth - 3
        Polygon m_picMenu.hdc, pts(0), 4



        'Lower Arrow+Button
        If Not m_bScrollArrowDown Then
            tmpArrowDownOffset = 0
            m_picMenu.DrawStyle = 0
        Else
            tmpArrowDownOffset = 1
            m_picMenu.DrawStyle = 2
        End If
        pts(0).x = m_picMenu.ScaleWidth - m_LonScrollBarWidth + 2: pts(0).y = m_picMenu.ScaleHeight - m_LonScrollBarWidth + 2
        pts(1).x = m_picMenu.ScaleWidth - 3: pts(1).y = pts(0).y
        pts(2).x = pts(1).x: pts(2).y = m_picMenu.ScaleHeight - 3
        pts(3).x = pts(0).x: pts(3).y = m_picMenu.ScaleHeight - 3
        Polygon m_picMenu.hdc, pts(0), 4


        'Middle Bar
        m_picMenu.DrawStyle = IIf(m_bScrollHandleDown = True, 2, 0)
        pts(0).x = m_picMenu.ScaleWidth - m_LonScrollBarWidth + 2: pts(0).y = m_LonScrollTop + ((m_LonScrollMax - m_LonScrollHeight) / (m_LonItemCount - m_LonItemsVisible) * m_LonItemAtTop)
        pts(1).x = m_picMenu.ScaleWidth - 3: pts(1).y = pts(0).y
        pts(2).x = m_picMenu.ScaleWidth - 3: pts(2).y = pts(0).y + m_LonScrollHeight
        pts(3).x = m_picMenu.ScaleWidth - m_LonScrollBarWidth + 2: pts(3).y = pts(0).y + m_LonScrollHeight
        Polygon m_picMenu.hdc, pts(0), 4

        m_picMenu.ForeColor = IIf(m_bEnabled, m_OleBorderColor, m_OleBorderColorDisabled)
        Set m_picMenu.Font = m_StdFontWebdings
        m_picMenu.CurrentX = Round(m_picMenu.ScaleWidth - (m_LonScrollBarWidth / 2 + m_picMenu.TextWidth("5") / 2)) + tmpArrowUpOffset
        m_picMenu.CurrentY = Round(m_LonScrollBarWidth / 2 - m_picMenu.TextHeight("5") / 2 - 1) + tmpArrowUpOffset
        m_picMenu.Print "5"

        m_picMenu.CurrentX = Round(m_picMenu.ScaleWidth - (m_LonScrollBarWidth / 2 + m_picMenu.TextWidth("6") / 2)) + tmpArrowDownOffset
        m_picMenu.CurrentY = Round(m_picMenu.ScaleHeight - m_LonScrollBarWidth / 2 - m_picMenu.TextHeight("6") / 2 - 1) + tmpArrowDownOffset
        m_picMenu.Print "6"

    End If

    Set m_picMenu.Font = m_StdFont



    m_bRefreshingMenu = False
End Sub


Sub OpenMenu()
    Dim tmpP As POINTAPI

    If m_LonItemCount = 0 Then Exit Sub

    tmpP = CalculatePosition
    m_picMenu.Height = (m_LonItemsVisible * m_LonItemHeight + 4) * Screen.TwipsPerPixelY
    m_picMenu.Left = tmpP.x

    If tmpP.y + UserControl.Height - Screen.TwipsPerPixelY + m_picMenu.Height > Screen.Height Then
        m_picMenu.Top = tmpP.y - m_picMenu.Height + Screen.TwipsPerPixelY
    Else
        m_picMenu.Top = tmpP.y + UserControl.Height - Screen.TwipsPerPixelY
    End If

    m_picMenu.Visible = True
    m_picMenu.ZOrder 0

    m_LonScrollTop = m_LonScrollBarWidth - 1
    m_LonScrollMax = m_picMenu.ScaleHeight - m_LonScrollTop * 2 - 1
    m_LonScrollHeight = m_LonScrollMax / m_LonItemCount * m_LonItemsVisible  'UserControl.ScaleHeight - 6
    If m_LonScrollHeight < 6 Then m_LonScrollHeight = 6

    m_bScrollHandleVisible = m_LonScrollHeight < m_LonScrollMax

    SetTopMostWindow m_picMenu.hWnd, True
    m_tmrFocus.Interval = 100
    m_tmrFocus.Enabled = True
    m_bMenuDown = True
End Sub

Sub CloseMenu()
    m_picMenu.Visible = False
    m_bMenuDown = False
    m_tmrFocus.Enabled = False
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode

        Case vbKeyDown
            If m_LonListIndex < m_LonItemCount - 1 Then
                m_LonListIndex = m_LonListIndex + 1
                m_StrText = Items(m_LonListIndex).Text
                RaiseEvent ItemChange(m_LonListIndex)
            End If

        Case vbKeyUp
            If m_LonListIndex > 0 Then
                m_LonListIndex = m_LonListIndex - 1
                m_StrText = Items(m_LonListIndex).Text
                RaiseEvent ItemChange(m_LonListIndex)
            End If


    End Select

    Redraw
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not m_bEnabled Then
        Redraw
        Exit Sub
    End If
    
    CheckMenuOpenButton Button, Shift, x, y

    'If Not m_bPicButtonDown Then
        'If m_bGotFocus Then
    If m_bMenuDown = False Then
        RaiseEvent OnDropdown
        OpenMenu
    Else
        CloseMenu
    End If
        ' End If
    'End If
    Redraw
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_PoiMouse.x = x
    m_PoiMouse.y = y

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not m_bEnabled Then
        Redraw
        Exit Sub
    End If
    
    CheckMenuOpenButton Button, Shift, x, y

'    If m_bPicButtonDown Then
'        m_bMenuDown = Not m_bMenuDown
'        If m_bMenuDown = True Then
'            RaiseEvent OnDropdown
'            OpenMenu
'        Else
'            CloseMenu
'        End If
'    End If

    m_bPicButtonDown = False
    Redraw
End Sub


Sub CheckMenuOpenButton(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_bPicButtonDown = False

    If x > UserControl.ScaleWidth - UserControl.ScaleHeight + 2 And x < UserControl.ScaleWidth - 2 Then
        If y > 1 And y < UserControl.ScaleHeight - 1 Then
            m_bPicButtonDown = True
        End If
    End If


End Sub



Private Sub Usercontrol_Resize()

'    m_picOpen.Width = UserControl.ScaleHeight - 2
'    m_picOpen.Left = UserControl.ScaleWidth - m_picOpen.Width - 1
'    m_picOpen.Top = 1
'    m_picOpen.Height = m_picOpen.Width

    If Not m_bStarting Then Redraw
End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_bStarting = True
    With PropBag
        m_OleBackgroundColor = .ReadProperty("BackgroundColor", &HFFFFFF)
        m_OleBorderColor = .ReadProperty("BorderColor", &HFFFFFF)
        m_OleForeColor = .ReadProperty("ForeColor", &H0)
        m_OleSelectionBackgroundColor = .ReadProperty("SelectionBackgroundColor", &H0)
        m_OleSelectionBorderColor = .ReadProperty("SelectionBorderColor", &H0)
        
        m_OleBackgroundColorDisabled = .ReadProperty("BackgroundColorDisabled", &HFFFFFF)
        m_OleBorderColorDisabled = .ReadProperty("BorderColorDisabled", &HFFFFFF)
        m_OleForeColorDisabled = .ReadProperty("ForeColorDisabled", &H0)
        m_OleSelectionBackgroundColorDisabled = .ReadProperty("SelectionBackgroundColorDisabled", &H0)
        m_OleSelectionBorderColorDisabled = .ReadProperty("SelectionBorderColorDisabled", &H0)
        
        Set Font = .ReadProperty("Font", Ambient.Font)
        m_StrText = .ReadProperty("Text", "Button")
        m_bBorder = .ReadProperty("Border", True)
        MousePointer = .ReadProperty("MousePointer", 0)
        m_LonItemHeight = .ReadProperty("ItemHeight", 25)
        m_LonItemsVisible = .ReadProperty("VisibleItems", 5)
        m_LonScrollBarWidth = .ReadProperty("ScrollBarWidth", 20)
    End With
    m_bStarting = False
    Redraw
End Sub

Private Sub UserControl_Terminate()
    SetParent m_picMenu.hWnd, GetParent(UserControl.hWnd)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackgroundColor", m_OleBackgroundColor, &HFFFFFF
        .WriteProperty "BorderColor", m_OleBorderColor, &HFFFFFF
        .WriteProperty "ForeColor", m_OleForeColor, &H0
        .WriteProperty "SelectionBackgroundColor", m_OleSelectionBackgroundColor, &H0
        .WriteProperty "SelectionBorderColor", m_OleSelectionBorderColor, &H0
        
        .WriteProperty "BackgroundColorDisabled", m_OleBackgroundColorDisabled, &HFFFFFF
        .WriteProperty "BorderColorDisabled", m_OleBorderColorDisabled, &HFFFFFF
        .WriteProperty "ForeColorDisabled", m_OleForeColorDisabled, &H0
        .WriteProperty "SelectionBackgroundColorDisabled", m_OleSelectionBackgroundColorDisabled, &H0
        .WriteProperty "SelectionBorderColorDisabled", m_OleSelectionBorderColorDisabled, &H0
        
        .WriteProperty "Font", m_StdFont, Ambient.Font
        .WriteProperty "Text", m_StrText, "Button"
        .WriteProperty "Border", m_bBorder, True
        .WriteProperty "MousePointer", m_MouMousePointer, 0
        .WriteProperty "ItemHeight", m_LonItemHeight, 25
        .WriteProperty "VisibleItems", m_LonItemsVisible, 5
        .WriteProperty "ScrollBarWidth", m_LonScrollBarWidth, 20
    End With
End Sub





