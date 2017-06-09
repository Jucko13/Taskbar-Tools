VERSION 5.00
Begin VB.UserControl uFrame 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   ControlContainer=   -1  'True
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   199
End
Attribute VB_Name = "uFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_OleBackgroundColor As OLE_COLOR
Private m_OleForeColor As OLE_COLOR
Private m_StrCaption As String
Private m_OleBorderColor As OLE_COLOR
Private m_bBorder As Boolean
Private m_bStarting As Boolean
Private m_MouMousePointer As MousePointerConstants
Private m_bRefreshing As Boolean

Private m_StdFont As StdFont

Public Property Get ScaleWidth() As Single
    ScaleWidth = UserControl.Width '* Screen.TwipsPerPixelX
End Property

Public Property Get ScaleHeight() As Single
    ScaleHeight = UserControl.Height '* Screen.TwipsPerPixelY
End Property


Public Property Get Font() As StdFont
    Set Font = m_StdFont
End Property

Public Property Set Font(ByVal StdValue As StdFont)
    Set m_StdFont = StdValue
    Set UserControl.Font = m_StdFont
    PropertyChanged "Font"
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

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_OleBorderColor
End Property

Public Property Let BorderColor(ByVal OleValue As OLE_COLOR)
    m_OleBorderColor = OleValue
    If Not m_bStarting Then Redraw
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = -518
    Caption = m_StrCaption
End Property

Public Property Let Caption(ByVal StrValue As String)
    m_StrCaption = StrValue
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


Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_Initialize()
    m_bStarting = True
    m_OleForeColor = &H0
    m_StrCaption = "uFrame"
    m_OleBorderColor = &HFFFFFF
    m_OleBackgroundColor = &HFFFFFF
    m_bBorder = True

    m_MouMousePointer = 0
    m_bStarting = False
    Redraw

End Sub


Sub Redraw()
    If m_bRefreshing Then Exit Sub
    m_bRefreshing = True

    Dim tmpTextHeight As Long
    Dim tmpTextWidth As Long

    tmpTextHeight = UserControl.TextHeight(m_StrCaption)
    tmpTextWidth = UserControl.TextWidth(m_StrCaption)

    UserControl.Cls
    UserControl.BackColor = m_OleBackgroundColor
    UserControl.ForeColor = m_OleForeColor

    If m_bBorder Then
        UserControl.Line (0, tmpTextHeight / 2)-(0, UserControl.ScaleHeight), m_OleBorderColor

        If Not Len(m_StrCaption) = 0 Then
            UserControl.Line (0, tmpTextHeight / 2)-(6, tmpTextHeight / 2), m_OleBorderColor
            UserControl.Line (6 + 2 + tmpTextWidth + 3, tmpTextHeight / 2)-(UserControl.ScaleWidth, tmpTextHeight / 2), m_OleBorderColor
        Else
            UserControl.Line (0, tmpTextHeight / 2)-(UserControl.ScaleWidth, tmpTextHeight / 2), m_OleBorderColor
        End If

        UserControl.Line (UserControl.ScaleWidth - 1, tmpTextHeight / 2)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), m_OleBorderColor
        UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), m_OleBorderColor
    End If
    UserControl.CurrentX = 6 + 2
    UserControl.CurrentY = 0
    UserControl.Print m_StrCaption
    
    
    If Not uDontDrawDots Then
        Dim X As Long
        Dim Y As Long
        
        For X = 0 To UserControl.ScaleWidth Step 3
            For Y = Int(tmpTextHeight / 2) + 4 To UserControl.ScaleHeight Step 3
                UserControl.PSet (X, Y), m_OleBorderColor
            Next Y
        Next X
    End If
    
    m_bRefreshing = False
End Sub


Private Sub UserControl_Resize()
    Redraw
End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_OleBackgroundColor = .ReadProperty("BackgroundColor", &HFFFFFF)
        m_OleBorderColor = .ReadProperty("BorderColor", &HFFFFFF)
        m_OleForeColor = .ReadProperty("ForeColor", &H0)
        m_StrCaption = .ReadProperty("Caption", "Button")
        m_bBorder = .ReadProperty("Border", True)
        MousePointer = .ReadProperty("MousePointer", 0)
        Set Font = .ReadProperty("Font", Ambient.Font)
    End With
    Redraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackgroundColor", m_OleBackgroundColor, &HFFFFFF
        .WriteProperty "BorderColor", m_OleBorderColor, &HFFFFFF
        .WriteProperty "ForeColor", m_OleForeColor, &H0
        .WriteProperty "Caption", m_StrCaption, "Button"
        .WriteProperty "Border", m_bBorder, True
        .WriteProperty "MousePointer", m_MouMousePointer, 0
        .WriteProperty "Font", m_StdFont, Ambient.Font
    End With
End Sub





