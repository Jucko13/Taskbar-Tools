VERSION 5.00
Begin VB.UserControl uButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1080
   ScaleHeight     =   51
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   72
End
Attribute VB_Name = "uButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type




Public Event Click(Button As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnter()
Public Event MouseLeave()

Public Enum ButtonOnClick
    [Move Text]
    [Move Border Out]
    [Move Border In]
    [Border In and Text]
    [Border Out and Text]
    [None]
End Enum


'Private cImage As c32bppDIB
'Private m_GDItoken As Long


Private m_bStarting As Boolean
Private m_StrCaption As String
Private m_bButtonDown As Boolean

Private m_OleBackgroundColor As OLE_COLOR
Private m_OleForeColor As OLE_COLOR
Private m_OleBorderColor As OLE_COLOR
Private m_OleCaptionBorderColor As OLE_COLOR
Private m_OleMouseOverBackgroundColor As OLE_COLOR
Private m_OleFocusColor As OLE_COLOR

Private m_OleBackgroundColorDisabled As OLE_COLOR
Private m_OleForeColorDisabled As OLE_COLOR
Private m_OleBorderColorDisabled As OLE_COLOR
Private m_OleCaptionBorderColorDisabled As OLE_COLOR
Private m_OleMouseOverBackgroundColorDisabled As OLE_COLOR
Private m_OleFocusColorDisabled As OLE_COLOR

Private m_bFocusVisible As Boolean
Private m_bBorder As Boolean
Private m_ButButtonAnimation As ButtonOnClick
Private m_StdPicture As StdPicture
Private m_StdPictureMouseOver As StdPicture
Private m_bAlignPictureInCorner As Boolean
Private m_MouMousePointer As MousePointerConstants
Private m_bRefreshing As Boolean
Private m_StdFont As stdole.StdFont
Private m_bCaptionBorder As Boolean

Private m_intCaptionOffsetLeft As Integer
Private m_intCaptionOffsetTop As Integer
Private m_bEnabled As Boolean

Private WithEvents m_tmrMouseOver As Timer
Attribute m_tmrMouseOver.VB_VarHelpID = -1
Private m_PoiMousePosition As POINTAPI
Private m_bMouseOver As Boolean

Private m_bHasFocus As Boolean


Public Property Get MouseOverBackgroundColor() As OLE_COLOR
    MouseOverBackgroundColor = m_OleMouseOverBackgroundColor
End Property

Public Property Let MouseOverBackgroundColor(ByVal OleValue As OLE_COLOR)
    m_OleMouseOverBackgroundColor = OleValue
    PropertyChanged "MouseOverBackgroundColor"
    If Not m_bStarting Then Redraw
End Property

Public Property Get CaptionBorderColor() As OLE_COLOR
    CaptionBorderColor = m_OleCaptionBorderColor
End Property

Public Property Let CaptionBorderColor(ByVal OleValue As OLE_COLOR)
    m_OleCaptionBorderColor = OleValue
    PropertyChanged "CaptionBorderColor"
    If Not m_bStarting Then Redraw
End Property


Public Property Get FocusVisible() As Boolean
    FocusVisible = m_bFocusVisible
End Property

Public Property Let FocusVisible(ByVal bValue As Boolean)
    m_bFocusVisible = bValue
    PropertyChanged "FocusVisible"
    If Not m_bStarting Then Redraw
End Property



Public Property Get FocusColor() As OLE_COLOR
    FocusColor = m_OleFocusColor
End Property

Public Property Let FocusColor(ByVal OleValue As OLE_COLOR)
    m_OleFocusColor = OleValue
    PropertyChanged "FocusColor"
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


Public Property Get AlignPictureInCorner() As Boolean
    AlignPictureInCorner = m_bAlignPictureInCorner
End Property

Public Property Let AlignPictureInCorner(ByVal bValue As Boolean)
    m_bAlignPictureInCorner = bValue
    PropertyChanged "AlignPictureInCorner"
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









Public Property Get MouseOverBackgroundColorDisabled() As OLE_COLOR
    MouseOverBackgroundColorDisabled = m_OleMouseOverBackgroundColorDisabled
End Property

Public Property Let MouseOverBackgroundColorDisabled(ByVal OleValue As OLE_COLOR)
    m_OleMouseOverBackgroundColorDisabled = OleValue
    PropertyChanged "MouseOverBackgroundColorDisabled"
    If Not m_bStarting Then Redraw
End Property

Public Property Get CaptionBorderColorDisabled() As OLE_COLOR
    CaptionBorderColorDisabled = m_OleCaptionBorderColorDisabled
End Property

Public Property Let CaptionBorderColorDisabled(ByVal OleValue As OLE_COLOR)
    m_OleCaptionBorderColorDisabled = OleValue
    PropertyChanged "CaptionBorderColorDisabled"
    If Not m_bStarting Then Redraw
End Property

Public Property Get FocusColorDisabled() As OLE_COLOR
    FocusColorDisabled = m_OleFocusColorDisabled
End Property

Public Property Let FocusColorDisabled(ByVal OleValue As OLE_COLOR)
    m_OleFocusColorDisabled = OleValue
    PropertyChanged "FocusColorDisabled"
    If Not m_bStarting Then Redraw
End Property

Public Property Get BorderColorDisabled() As OLE_COLOR
    BorderColorDisabled = m_OleBorderColorDisabled
End Property

Public Property Let BorderColorDisabled(ByVal OleValue As OLE_COLOR)
    m_OleBorderColorDisabled = OleValue
    PropertyChanged "BorderColorDisabled"
    If Not m_bStarting Then Redraw
End Property

Public Property Get ForeColorDisabled() As OLE_COLOR
    ForeColorDisabled = m_OleForeColorDisabled
End Property

Public Property Let ForeColorDisabled(ByVal OleValue As OLE_COLOR)
    m_OleForeColorDisabled = OleValue
    PropertyChanged "ForeColorDisabled"
    If Not m_bStarting Then Redraw
End Property

Public Property Get BackgroundColorDisabled() As OLE_COLOR
    BackgroundColorDisabled = m_OleBackgroundColorDisabled
End Property

Public Property Let BackgroundColorDisabled(ByVal OleValue As OLE_COLOR)
    m_OleBackgroundColorDisabled = OleValue
    PropertyChanged "BackgroundColorDisabled"
    If Not m_bStarting Then Redraw
End Property










Public Property Get CaptionOffsetTop() As Integer
    CaptionOffsetTop = m_intCaptionOffsetTop
End Property

Public Property Let CaptionOffsetTop(ByVal intValue As Integer)
    m_intCaptionOffsetTop = intValue
    PropertyChanged "CaptionOffsetTop"
    If Not m_bStarting Then Redraw
End Property

Public Property Get CaptionOffsetLeft() As Integer
    CaptionOffsetLeft = m_intCaptionOffsetLeft
End Property

Public Property Let CaptionOffsetLeft(ByVal intValue As Integer)
    m_intCaptionOffsetLeft = intValue
    PropertyChanged "CaptionOffsetLeft"
    If Not m_bStarting Then Redraw
End Property


Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property

Public Property Let Enabled(ByVal bValue As Boolean)
    m_bEnabled = bValue
    PropertyChanged "Enabled"
    If Not m_bStarting Then Redraw
End Property




Public Property Get CaptionBorder() As Boolean
    CaptionBorder = m_bCaptionBorder
End Property

Public Property Let CaptionBorder(ByVal bValue As Boolean)
    m_bCaptionBorder = bValue
    PropertyChanged "CaptionBorder"
    If Not m_bStarting Then Redraw
End Property


Public Property Get Font() As StdFont
    Set Font = m_StdFont
End Property

Public Property Set Font(ByVal StdValue As StdFont)
    Set m_StdFont = StdValue
    PropertyChanged "Font"
    If Not m_bStarting Then Redraw
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = m_MouMousePointer
End Property

Public Property Let MousePointer(ByVal MouValue As MousePointerConstants)
    m_MouMousePointer = MouValue
    UserControl.MousePointer = MouValue
    PropertyChanged "MousePointer"
End Property


Public Property Get PictureMouseOver() As StdPicture
    Set PictureMouseOver = m_StdPictureMouseOver
End Property

Public Property Set PictureMouseOver(ByVal StdValue As StdPicture)
    Set m_StdPictureMouseOver = StdValue

    PropertyChanged "PictureMouseOver"
    If Not m_bStarting Then Redraw
End Property


Public Property Get Picture() As StdPicture
    Set Picture = m_StdPicture
End Property

Public Property Set Picture(ByVal StdValue As StdPicture)
    Set m_StdPicture = StdValue

    PropertyChanged "Picture"
    If Not m_bStarting Then Redraw
End Property

Public Property Get ButtonAnimation() As ButtonOnClick
    ButtonAnimation = m_ButButtonAnimation
End Property

Public Property Let ButtonAnimation(ByVal ButValue As ButtonOnClick)
    m_ButButtonAnimation = ButValue
    PropertyChanged "ButtonAnimation"
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



Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = -518
    Caption = m_StrCaption
End Property

Public Property Let Caption(ByVal StrValue As String)
    m_StrCaption = StrValue
    PropertyChanged "Caption"
    If Not m_bStarting Then Redraw
End Property

Function isMouseOverControl() As Boolean
    GetCursorPos m_PoiMousePosition
    isMouseOverControl = CBool(WindowFromPoint(m_PoiMousePosition.X, m_PoiMousePosition.Y) = UserControl.hWnd)
End Function

Private Sub m_tmrMouseOver_Timer()
    If Not isMouseOverControl Then
        m_tmrMouseOver.Enabled = False
        m_bMouseOver = False
        RaiseEvent MouseLeave
        Redraw
    End If
End Sub

Private Sub UserControl_DblClick()
    m_bButtonDown = True
    If Not m_bStarting Then Redraw
End Sub

Private Sub UserControl_EnterFocus()
    m_bHasFocus = True
    If Not m_bStarting Then Redraw
End Sub

Private Sub UserControl_ExitFocus()
    m_bHasFocus = False
    If Not m_bStarting Then Redraw
End Sub

Private Sub UserControl_Initialize()
    m_bStarting = True
    m_OleBackgroundColor = &HE18700
    m_OleForeColor = &H800000
    m_StrCaption = "Button"
    m_OleBorderColor = &HFFFFFF
    m_bBorder = True
    'm_ButButtonAnimation = [Move Border In]
    m_MouMousePointer = 0
    m_bEnabled = True
    
    Set m_tmrMouseOver = UserControl.Controls.Add("VB.Timer", "m_tmrMouseOver")

    Set m_StdFont = UserControl.Font

    m_bCaptionBorder = False
    m_OleCaptionBorderColor = &HFFFFFF
    m_intCaptionOffsetLeft = 0
    m_intCaptionOffsetTop = 0
    m_bStarting = False

    Redraw
    m_OleMouseOverBackgroundColor = &H0
End Sub

Sub Redraw()
    If m_bRefreshing Then Exit Sub
    m_bRefreshing = True

    Dim tmpTextWidth As Long
    Dim tmpTextHeight As Long
    Dim tmpBorderOffset As Long
    Dim tmpTextOffset As Long
    Dim tmpTextStartY As Long
    Dim tmpCaptionSplit() As String
    Dim tmpFocusColor As OLE_COLOR
    Dim tmpX As Long
    Dim tmpY As Long
    Dim i As Long
    Dim j As Long
    Dim K As Long
    Dim t_StdPicture As StdPicture
    
    
    Set UserControl.Font = m_StdFont

    UserControl.AutoRedraw = True
    UserControl.Cls

    If m_bMouseOver Then
        UserControl.BackColor = IIf(m_bEnabled, m_OleMouseOverBackgroundColor, m_OleMouseOverBackgroundColorDisabled)
        Set t_StdPicture = m_StdPictureMouseOver
    Else
        UserControl.BackColor = IIf(m_bEnabled, m_OleBackgroundColor, m_OleBackgroundColorDisabled)
        Set t_StdPicture = m_StdPicture
    End If

    If m_bButtonDown Then
        Select Case m_ButButtonAnimation
            Case [Move Border In], [Border In and Text]
                tmpBorderOffset = 1

            Case [Move Border Out], [Border Out and Text]
                tmpBorderOffset = -1

        End Select


        Select Case m_ButButtonAnimation

            Case [Move Text], [Border Out and Text], [Border In and Text]
                tmpTextOffset = 1
        End Select
    End If

    If m_bBorder Then
        UserControl.Line (tmpBorderOffset, tmpBorderOffset)-(UserControl.ScaleWidth - 1 - tmpBorderOffset, tmpBorderOffset), IIf(m_bEnabled, m_OleBorderColor, m_OleBorderColorDisabled)
        UserControl.Line (UserControl.ScaleWidth - 1 - tmpBorderOffset, tmpBorderOffset)-(UserControl.ScaleWidth - 1 - tmpBorderOffset, UserControl.ScaleHeight - tmpBorderOffset), IIf(m_bEnabled, m_OleBorderColor, m_OleBorderColorDisabled)

        UserControl.Line (tmpBorderOffset, UserControl.ScaleHeight - 1 - tmpBorderOffset)-(UserControl.ScaleWidth - tmpBorderOffset, UserControl.ScaleHeight - 1 - tmpBorderOffset), IIf(m_bEnabled, m_OleBorderColor, m_OleBorderColorDisabled)
        UserControl.Line (tmpBorderOffset, tmpBorderOffset)-(tmpBorderOffset, UserControl.ScaleHeight - 1 - tmpBorderOffset), IIf(m_bEnabled, m_OleBorderColor, m_OleBorderColorDisabled)
    End If
    
    If m_bHasFocus And m_bFocusVisible Then
        tmpFocusColor = IIf(m_bEnabled, m_OleFocusColor, m_OleFocusColorDisabled)
        UserControl.DrawStyle = vbDot
        
        UserControl.Line (2, 2)-(UserControl.ScaleWidth - 3, 2), tmpFocusColor
        UserControl.Line (UserControl.ScaleWidth - 3, 2)-(UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3), tmpFocusColor

        UserControl.Line (2, UserControl.ScaleHeight - 3)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3), tmpFocusColor
        UserControl.Line (2, 2)-(2, UserControl.ScaleHeight - 3), tmpFocusColor
        
        UserControl.DrawStyle = vbSolid
    End If

    If Not t_StdPicture Is Nothing Then
        If m_bAlignPictureInCorner Then
            t_StdPicture.Render UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
            0, m_StdPicture.Height, t_StdPicture.Width, -t_StdPicture.Height, 0
            
        Else
            t_StdPicture.Render UserControl.hdc, _
                          Round(UserControl.ScaleWidth / 2 - ScaleX(t_StdPicture.Width, vbTwips, vbPixels) / 2) + tmpTextOffset, _
                          Round(UserControl.ScaleHeight / 2 - (ScaleY(t_StdPicture.Height, vbTwips, vbPixels) + tmpTextHeight) / 2) + tmpTextOffset, _
                          ScaleX(t_StdPicture.Width, vbTwips, vbPixels), ScaleY(t_StdPicture.Height, vbTwips, vbPixels), 0, 0, t_StdPicture.Width, t_StdPicture.Height, 0
            
        End If
        
        tmpX = Fix(UserControl.ScaleWidth / 2 - tmpTextWidth / 2) + tmpTextOffset
        tmpY = Fix((UserControl.ScaleHeight / 2 - (t_StdPicture.Height + tmpTextHeight) / 2)) + t_StdPicture.Height + tmpTextOffset
    End If
    
    tmpCaptionSplit = Split(m_StrCaption, vbCrLf)
    
    tmpTextWidth = UserControl.TextWidth(m_StrCaption)
    tmpTextHeight = UserControl.TextHeight(m_StrCaption)
        
    
    tmpTextStartY = UserControl.ScaleHeight / 2 - tmpTextHeight / 2
    
    For K = 0 To UBound(tmpCaptionSplit)
        
        tmpX = Fix(UserControl.ScaleWidth / 2 - UserControl.TextWidth(tmpCaptionSplit(K)) / 2) + tmpTextOffset
        tmpY = tmpTextStartY + tmpTextOffset
    
        tmpX = tmpX + m_intCaptionOffsetLeft
        tmpY = tmpY + m_intCaptionOffsetTop
    
        If m_bCaptionBorder Then
            UserControl.ForeColor = IIf(m_bEnabled, m_OleCaptionBorderColor, m_OleCaptionBorderColorDisabled)
            For i = -1 To 1
                For j = -1 To 1
                    If i <> 0 Or j <> 0 Then
                        UserControl.CurrentX = tmpX + i
                        UserControl.CurrentY = tmpY + j
                        UserControl.Print tmpCaptionSplit(K)
                    End If
                Next j
            Next i
        End If
    
        UserControl.CurrentX = tmpX
        UserControl.CurrentY = tmpY
        UserControl.ForeColor = IIf(m_bEnabled, m_OleForeColor, m_OleForeColorDisabled)
        UserControl.Print tmpCaptionSplit(K)
        
        tmpTextStartY = tmpTextStartY + UserControl.TextHeight(tmpCaptionSplit(K))
    Next K
    
    m_bRefreshing = False
End Sub


Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then RaiseEvent Click(-1, -1, -1)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_bButtonDown = True

    Redraw
    RaiseEvent MouseDown(Button, Shift, X, Y)
    'Debug.Print "down"
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 0 Or Y < 0 Or X > UserControl.ScaleWidth Or Y > UserControl.ScaleHeight Then
        m_bButtonDown = False
    Else
        m_bButtonDown = IIf(Not Button = 0, True, False)

        If m_bMouseOver = False Then
            m_bMouseOver = True
            m_tmrMouseOver.Interval = 40
            m_tmrMouseOver.Enabled = True
            RaiseEvent MouseEnter
        End If
    End If

    RaiseEvent MouseMove(Button, Shift, X, Y)
    Redraw
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)

    If X > UserControl.ScaleWidth Or Y > UserControl.ScaleHeight Or X < 0 Or Y < 0 Then
        GoTo NoClick
    End If


    If m_bButtonDown Then
        RaiseEvent Click(Button, X, Y)
    End If

NoClick:
    m_bButtonDown = False
    If Not m_bStarting Then Redraw
    'Debug.Print "up"
End Sub

Private Sub Usercontrol_Resize()
    If Not m_bStarting Then Redraw
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_bStarting = True
    With PropBag
        m_OleBackgroundColor = .ReadProperty("BackgroundColor", &HE18700)
        m_OleBorderColor = .ReadProperty("BorderColor", &HFFFFFF)
        m_OleForeColor = .ReadProperty("ForeColor", &H800000)
        m_OleMouseOverBackgroundColor = .ReadProperty("MouseOverBackgroundColor", &H0)
        m_OleCaptionBorderColor = .ReadProperty("CaptionBorderColor", &HFFFFFF)
        m_OleFocusColor = .ReadProperty("FocusColor", &HFFFFFF)
        
        m_OleBackgroundColorDisabled = .ReadProperty("BackgroundColorDisabled", &HE18700)
        m_OleBorderColorDisabled = .ReadProperty("BorderColorDisabled", &HFFFFFF)
        m_OleForeColorDisabled = .ReadProperty("ForeColorDisabled", &H800000)
        m_OleMouseOverBackgroundColorDisabled = .ReadProperty("MouseOverBackgroundColorDisabled", &H0)
        m_OleCaptionBorderColorDisabled = .ReadProperty("CaptionBorderColorDisabled", &HFFFFFF)
        m_OleFocusColorDisabled = .ReadProperty("FocusColorDisabled", &HFFFFFF)
        
        
        m_bFocusVisible = .ReadProperty("FocusVisible", True)
        m_StrCaption = .ReadProperty("Caption", "Button")
        m_bBorder = .ReadProperty("Border", True)
        m_ButButtonAnimation = .ReadProperty("BorderAnimation", [Move Border In])
        MousePointer = .ReadProperty("MousePointer", 0)
        Set m_StdFont = .ReadProperty("Font", Ambient.Font)
        m_bCaptionBorder = .ReadProperty("CaptionBorder", False)
        Set m_StdPicture = .ReadProperty("Picture", Nothing)
        Set m_StdPictureMouseOver = .ReadProperty("PictureMouseOver", Nothing)
        m_bAlignPictureInCorner = .ReadProperty("AlignPictureInCorner", False)
        
        m_intCaptionOffsetLeft = .ReadProperty("CaptionOffsetLeft", 0)
        m_intCaptionOffsetTop = .ReadProperty("CaptionOffsetTop", 0)
        m_bEnabled = .ReadProperty("Enabled", True)
        
    End With
    m_bStarting = False
    Redraw

End Sub

Private Sub UserControl_Terminate()
    m_bStarting = True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackgroundColor", m_OleBackgroundColor, &HE18700
        .WriteProperty "BorderColor", m_OleBorderColor, &HFFFFFF
        .WriteProperty "ForeColor", m_OleForeColor, &H800000
        .WriteProperty "MouseOverBackgroundColor", m_OleMouseOverBackgroundColor, &H0
        .WriteProperty "CaptionBorderColor", m_OleCaptionBorderColor, &HFFFFFF
        .WriteProperty "FocusColor", m_OleFocusColor, &HFFFFFF
        
        .WriteProperty "BackgroundColorDisabled", m_OleBackgroundColorDisabled, &HE18700
        .WriteProperty "BorderColorDisabled", m_OleBorderColorDisabled, &HFFFFFF
        .WriteProperty "ForeColorDisabled", m_OleForeColorDisabled, &H800000
        .WriteProperty "MouseOverBackgroundColorDisabled", m_OleMouseOverBackgroundColorDisabled, &H0
        .WriteProperty "CaptionBorderColorDisabled", m_OleCaptionBorderColorDisabled, &HFFFFFF
        .WriteProperty "FocusColorDisabled", m_OleFocusColorDisabled, &HFFFFFF
        
        
        .WriteProperty "FocusVisible", m_bFocusVisible, True
        .WriteProperty "Caption", m_StrCaption, "Button"
        .WriteProperty "Border", m_bBorder, True
        .WriteProperty "BorderAnimation", m_ButButtonAnimation, [Move Border In]
        .WriteProperty "Picture", m_StdPicture, Nothing
        .WriteProperty "PictureMouseOver", m_StdPictureMouseOver, Nothing
        .WriteProperty "AlignPictureInCorner", m_bAlignPictureInCorner, False
        
        .WriteProperty "MousePointer", m_MouMousePointer, 0
        .WriteProperty "Font", m_StdFont, Ambient.Font
        .WriteProperty "CaptionBorder", m_bCaptionBorder, False
        .WriteProperty "CaptionOffsetLeft", m_intCaptionOffsetLeft, 0
        .WriteProperty "CaptionOffsetTop", m_intCaptionOffsetTop, 0
        
        .WriteProperty "Enabled", m_bEnabled, True

    End With

End Sub



