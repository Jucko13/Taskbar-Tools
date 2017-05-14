VERSION 5.00
Begin VB.UserControl uCheckBox 
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
Attribute VB_Name = "uCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'checselectedbordercolor toevoegen

'als de checkbox height aangepast word dan veranderd de box van grootte en klopt de verhouding niet meer.


Public Enum uCheckboxConstants
    u_unChecked = 0
    u_Checked = 1
    u_PartialChecked = 2
    u_Cross = 3
End Enum

Public Enum uCheckSizes
    u_Smalllest = 0
    u_Small = 1
    u_Normal = 2
    u_Big = 3
    u_Biggest = 4
End Enum

Public Event ActivateNextState(ByRef u_Cancel As Boolean, ByRef u_NewState As uCheckboxConstants)
Public Event Changed(ByRef u_NewState As uCheckboxConstants)

Private m_OleBackgroundColor As OLE_COLOR
Private m_OleForeColor As OLE_COLOR

Private m_bBorder As Boolean
Private m_OleBorderColor As OLE_COLOR
Private m_LonBorderThickness As Long

Private m_bStarting As Boolean
Private m_MouMousePointer As MousePointerConstants
Private m_bRefreshing As Boolean
Private m_OleCheckBorderColor As OLE_COLOR
Private m_LonCheckBorderThickness As Long
Private m_OleCheckBackgroundColor As OLE_COLOR
Private m_OleCheckSelectionColor As OLE_COLOR
Private m_UChValue As uCheckboxConstants
Private m_UChCheckSize As uCheckSizes
Private m_SinCheckSize As Long

Private m_StrCaption As String
Private m_bCaptionBorder As Boolean
Private m_OleCaptionBorderColor As OLE_COLOR
Private m_intCaptionOffsetLeft As Integer
Private m_intCaptionOffsetTop As Integer
Private m_StdFont As StdFont


Public Property Get BorderThickness() As Long
    BorderThickness = m_LonBorderThickness
End Property

Public Property Let BorderThickness(ByVal LonValue As Long)
    m_LonBorderThickness = LonValue
    PropertyChanged "BorderThickness"
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

Public Property Get Font() As StdFont
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = m_StdFont
End Property

Public Property Set Font(ByRef StdValue As StdFont)
    Set m_StdFont = StdValue
    Set UserControl.Font = m_StdFont
    PropertyChanged "Font"
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

Public Property Get CaptionBorder() As Boolean
    CaptionBorder = m_bCaptionBorder
End Property

Public Property Let CaptionBorder(ByVal bValue As Boolean)
    m_bCaptionBorder = bValue
    PropertyChanged "CaptionBorder"
    If Not m_bStarting Then Redraw
End Property


Public Property Get CheckSize() As uCheckSizes
    CheckSize = m_UChCheckSize
End Property

Public Property Let CheckSize(ByVal UChValue As uCheckSizes)
    m_UChCheckSize = UChValue

    Select Case m_UChCheckSize
        Case uCheckSizes.u_Smalllest
            m_SinCheckSize = 6

        Case uCheckSizes.u_Small
            m_SinCheckSize = 8

        Case uCheckSizes.u_Normal
            m_SinCheckSize = 12

        Case uCheckSizes.u_Big
            m_SinCheckSize = 20

        Case uCheckSizes.u_Biggest
            m_SinCheckSize = 32
    End Select
    PropertyChanged "CheckSize"
    If Not m_bStarting Then Redraw
End Property


Public Property Get Value() As uCheckboxConstants
    Value = m_UChValue
End Property

Public Property Let Value(ByVal UChValue As uCheckboxConstants)
    RaiseEvent Changed(UChValue)
    m_UChValue = UChValue
    PropertyChanged "Value"
    If Not m_bStarting Then Redraw
End Property

Public Property Get CheckSelectionColor() As OLE_COLOR
    CheckSelectionColor = m_OleCheckSelectionColor
End Property

Public Property Let CheckSelectionColor(ByVal OleValue As OLE_COLOR)
    m_OleCheckSelectionColor = OleValue
    PropertyChanged "CheckSelectionColor"
    If Not m_bStarting Then Redraw
End Property

Public Property Get CheckBackgroundColor() As OLE_COLOR
    CheckBackgroundColor = m_OleCheckBackgroundColor
End Property

Public Property Let CheckBackgroundColor(ByVal OleValue As OLE_COLOR)
    m_OleCheckBackgroundColor = OleValue
    PropertyChanged "CheckBackgroundColor"
    If Not m_bStarting Then Redraw
End Property

Public Property Get CheckBorderThickness() As Long
    CheckBorderThickness = m_LonCheckBorderThickness
End Property

Public Property Let CheckBorderThickness(ByVal LonValue As Long)
    m_LonCheckBorderThickness = LonValue
    If m_LonCheckBorderThickness < 0 Then
        m_LonCheckBorderThickness = 0
    End If
    PropertyChanged "CheckBorderThickness"
    If Not m_bStarting Then Redraw
End Property

Public Property Get CheckBorderColor() As OLE_COLOR
    CheckBorderColor = m_OleCheckBorderColor
End Property

Public Property Let CheckBorderColor(ByVal OleValue As OLE_COLOR)
    m_OleCheckBorderColor = OleValue
    PropertyChanged "CheckBorderColor"
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

Private Sub UserControl_Initialize()
    m_bStarting = True
    m_OleForeColor = &H0
    m_StrCaption = "uFrame"
    m_OleBorderColor = &HFFFFFF
    m_OleBackgroundColor = &HFFFFFF
    m_bBorder = True

    m_OleCheckBorderColor = &H0
    m_LonCheckBorderThickness = 1
    m_OleCheckBackgroundColor = &HFFFFFF
    m_OleCheckSelectionColor = &HFF00FF

    m_MouMousePointer = 0
    Set Font = UserControl.Font
    m_UChValue = uCheckboxConstants.u_unChecked
    m_UChCheckSize = uCheckSizes.u_Normal
    m_bCaptionBorder = False
    m_OleCaptionBorderColor = &HFFFFFF


    m_bStarting = False
    m_intCaptionOffsetLeft = 0
    m_intCaptionOffsetTop = 0
    m_LonBorderThickness = 1
End Sub


Sub Redraw()
    If m_bRefreshing Then Exit Sub
    m_bRefreshing = True

    Dim tmpTextHeight As Long
    Dim tmpTextWidth As Long
    Dim tmpX As Single
    Dim tmpY As Single
    Dim tmpHeight As Single
    Dim tmpWidth As Single
    Dim tmpOffsetAdj1 As Single
    Dim tmpOffsetAdj2 As Single
    Dim tmpOffsetAdj3 As Single
    Dim tmpCapX As Long
    Dim tmpCapY As Long
    Dim i As Long
    Dim j As Long

    tmpTextHeight = UserControl.TextHeight(m_StrCaption)
    tmpTextWidth = UserControl.TextWidth(m_StrCaption)

    UserControl.Cls

    ReDim pts(0 To 20) As POINTAPI

    If m_bBorder And m_LonBorderThickness > 0 Then
        UserControl.BackColor = m_OleBorderColor

        If m_LonBorderThickness < UserControl.ScaleHeight / 2 Then
            UserControl.FillColor = m_OleBackgroundColor
            UserControl.FillStyle = 0
            UserControl.DrawStyle = 5

            pts(0).x = m_LonBorderThickness
            pts(0).y = m_LonBorderThickness

            pts(1).x = UserControl.ScaleWidth - m_LonBorderThickness
            pts(1).y = m_LonBorderThickness

            pts(2).x = UserControl.ScaleWidth - m_LonBorderThickness
            pts(2).y = UserControl.ScaleHeight - m_LonBorderThickness

            pts(3).x = m_LonBorderThickness
            pts(3).y = UserControl.ScaleHeight - m_LonBorderThickness

            Polygon UserControl.hdc, pts(0), 4
        End If

    ElseIf Not m_bBorder Or m_LonBorderThickness < 1 Then
        UserControl.BackColor = m_OleBackgroundColor
    End If


    UserControl.DrawStyle = 0

    tmpWidth = m_SinCheckSize
    tmpHeight = tmpWidth / 2

    tmpOffsetAdj1 = (tmpHeight / 5)
    tmpOffsetAdj2 = (tmpHeight / 2.5)
    tmpOffsetAdj3 = (tmpHeight / (10 / 3))

    tmpX = Fix(UserControl.ScaleHeight / 2) - tmpHeight
    tmpY = Fix(UserControl.ScaleHeight / 2)

    pts(0).x = tmpX
    pts(0).y = tmpY - tmpHeight

    pts(1).x = tmpX + tmpWidth
    pts(1).y = tmpY - tmpHeight

    pts(2).x = tmpX + tmpWidth
    pts(2).y = tmpY + tmpHeight

    pts(3).x = tmpX
    pts(3).y = tmpY + tmpHeight

    If m_LonCheckBorderThickness > 0 Then
        pts(4).x = pts(0).x - m_LonCheckBorderThickness
        pts(4).y = pts(0).y - m_LonCheckBorderThickness

        pts(5).x = pts(1).x + m_LonCheckBorderThickness
        pts(5).y = pts(1).y - m_LonCheckBorderThickness

        pts(6).x = pts(2).x + m_LonCheckBorderThickness
        pts(6).y = pts(2).y + m_LonCheckBorderThickness

        pts(7).x = pts(3).x - m_LonCheckBorderThickness
        pts(7).y = pts(3).y + m_LonCheckBorderThickness

        UserControl.ForeColor = m_OleCheckBorderColor
        UserControl.FillColor = m_OleCheckBorderColor

        Polygon UserControl.hdc, pts(4), 4
    End If

    UserControl.ForeColor = m_OleCheckBackgroundColor
    UserControl.FillColor = m_OleCheckBackgroundColor
    Polygon UserControl.hdc, pts(0), 4


    tmpCapX = tmpX + tmpWidth + 5 + m_intCaptionOffsetLeft
    tmpCapY = tmpY - Fix(UserControl.TextHeight(m_StrCaption) / 2) - 1 + m_intCaptionOffsetTop
    If m_bCaptionBorder Then
        UserControl.ForeColor = m_OleCaptionBorderColor
        For i = -1 To 1
            For j = -1 To 1
                If i <> 0 Or j <> 0 Then
                    UserControl.CurrentX = tmpCapX + i
                    UserControl.CurrentY = tmpCapY + j
                    UserControl.Print m_StrCaption
                End If
            Next j
        Next i
    End If

    UserControl.ForeColor = m_OleForeColor
    UserControl.CurrentX = tmpCapX
    UserControl.CurrentY = tmpCapY

    UserControl.Print m_StrCaption

    UserControl.ForeColor = m_OleCheckSelectionColor
    UserControl.FillColor = m_OleCheckSelectionColor

    If m_UChValue = u_Checked Then
        pts(0).x = tmpX - tmpOffsetAdj1
        pts(0).y = tmpY - tmpHeight - tmpOffsetAdj1

        pts(1).x = tmpX + (tmpWidth / 2) + tmpOffsetAdj1
        pts(1).y = tmpY + tmpOffsetAdj1

        '(IIf(UserControl.ScaleHeight Mod 4 = 0, 1, 0)) + IIf((UserControl.ScaleHeight - 1) Mod 4 = 0, 1, 0)

        pts(2).x = tmpX + tmpWidth
        pts(2).y = tmpY - tmpHeight + tmpOffsetAdj2
        If m_UChCheckSize = u_Smalllest Then pts(2).y = pts(2).y + 1

        pts(3).x = pts(2).x + tmpOffsetAdj3
        pts(3).y = pts(2).y + tmpOffsetAdj3

        pts(4).x = pts(1).x
        pts(4).y = pts(1).y + (tmpOffsetAdj1 * 3)

        pts(5).x = pts(0).x - tmpOffsetAdj3
        pts(5).y = tmpY - tmpHeight + (tmpHeight / 10)

        Polygon UserControl.hdc, pts(0), 6
    ElseIf m_UChValue = u_PartialChecked Then
        pts(0).x = tmpX + 1
        pts(0).y = tmpY - tmpHeight + 1

        pts(1).x = tmpX + tmpWidth - 1
        pts(1).y = tmpY - tmpHeight + 1

        pts(2).x = tmpX + tmpWidth - 1
        pts(2).y = tmpY + tmpHeight - 1

        pts(3).x = tmpX + 1
        pts(3).y = tmpY + tmpHeight - 1

        Polygon UserControl.hdc, pts(0), 4

    ElseIf m_UChValue = u_Cross Then
        pts(0).x = tmpX
        pts(0).y = tmpY - tmpHeight - tmpOffsetAdj2

        pts(1).x = tmpX + (tmpWidth / 2)
        pts(1).y = tmpY - tmpOffsetAdj2

        pts(2).x = tmpX + tmpWidth
        pts(2).y = tmpY - tmpHeight - tmpOffsetAdj2

        pts(3).x = tmpX + tmpWidth + tmpOffsetAdj2
        pts(3).y = tmpY - tmpHeight

        pts(4).x = tmpX + (tmpWidth / 2) + tmpOffsetAdj2
        pts(4).y = tmpY

        pts(5).x = tmpX + tmpWidth + tmpOffsetAdj2
        pts(5).y = tmpY + tmpHeight

        pts(6).x = tmpX + tmpWidth
        pts(6).y = tmpY + tmpHeight + tmpOffsetAdj2

        pts(7).x = tmpX + tmpWidth / 2
        pts(7).y = tmpY + tmpOffsetAdj2

        pts(8).x = tmpX
        pts(8).y = tmpY + tmpHeight + tmpOffsetAdj2

        pts(9).x = tmpX - tmpOffsetAdj2
        pts(9).y = tmpY + tmpHeight

        pts(10).x = tmpX + tmpWidth / 2 - tmpOffsetAdj2
        pts(10).y = tmpY

        pts(11).x = tmpX - tmpOffsetAdj2
        pts(11).y = tmpY - tmpHeight

        Polygon UserControl.hdc, pts(0), 12
    End If


    m_bRefreshing = False
End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim tmpCancel As Boolean
    Dim tmpNewState As uCheckboxConstants

    tmpCancel = False
    tmpNewState = m_UChValue
    RaiseEvent ActivateNextState(tmpCancel, tmpNewState)

    If tmpCancel Then
        m_UChValue = tmpNewState
    Else
        If m_UChValue = u_Checked Then
            m_UChValue = u_unChecked
        ElseIf m_UChValue = u_unChecked Then
            m_UChValue = u_Checked
        End If
    End If
    
    RaiseEvent Changed(m_UChValue)
    
    If Not m_bStarting Then Redraw
End Sub

Private Sub Usercontrol_Resize()
    If Not m_bStarting Then Redraw
End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_bStarting = True
    With PropBag
        m_OleBackgroundColor = .ReadProperty("BackgroundColor", &HFFFFFF)
        m_bBorder = .ReadProperty("Border", True)
        m_OleBorderColor = .ReadProperty("BorderColor", &HFFFFFF)
        m_LonBorderThickness = .ReadProperty("BorderThickness", 1)

        m_StrCaption = .ReadProperty("Caption", "Button")
        m_bCaptionBorder = .ReadProperty("CaptionBorder", False)
        m_OleCaptionBorderColor = .ReadProperty("CaptionBorderColor", &HFFFFFF)
        m_intCaptionOffsetLeft = .ReadProperty("CaptionOffsetLeft", 0)
        m_intCaptionOffsetTop = .ReadProperty("CaptionOffsetTop", 0)

        m_OleCheckBackgroundColor = .ReadProperty("CheckBackgroundColor", &HFFFFFF)
        m_OleCheckBorderColor = .ReadProperty("CheckBorderColor", &H0)
        m_LonCheckBorderThickness = .ReadProperty("CheckBorderThickness", 1)
        m_OleCheckSelectionColor = .ReadProperty("CheckSelectionColor", &HFF00FF)
        CheckSize = .ReadProperty("CheckSize", uCheckSizes.u_Normal)

        Set Font = .ReadProperty("Font", Ambient.Font)
        m_OleForeColor = .ReadProperty("ForeColor", &H0)
        MousePointer = .ReadProperty("MousePointer", 0)
        m_UChValue = .ReadProperty("Value", uCheckboxConstants.u_unChecked)


    End With
    m_bStarting = False
    Redraw


End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackgroundColor", m_OleBackgroundColor, &HFFFFFF
        .WriteProperty "Border", m_bBorder, True
        .WriteProperty "BorderColor", m_OleBorderColor, &HFFFFFF
        .WriteProperty "BorderThickness", m_LonBorderThickness, 1

        .WriteProperty "Caption", m_StrCaption, "Button"
        .WriteProperty "CaptionBorder", m_bCaptionBorder, False
        .WriteProperty "CaptionBorderColor", m_OleCaptionBorderColor, &HFFFFFF
        .WriteProperty "CaptionOffsetLeft", m_intCaptionOffsetLeft, 0
        .WriteProperty "CaptionOffsetTop", m_intCaptionOffsetTop, 0

        .WriteProperty "CheckBackgroundColor", m_OleCheckBackgroundColor, &HFFFFFF
        .WriteProperty "CheckBorderColor", m_OleCheckBorderColor, &H0
        .WriteProperty "CheckBorderThickness", m_LonCheckBorderThickness, 1
        .WriteProperty "CheckSelectionColor", m_OleCheckSelectionColor, &HFF00FF
        .WriteProperty "CheckSize", m_UChCheckSize, uCheckSizes.u_Normal

        .WriteProperty "Font", m_StdFont, Ambient.Font
        .WriteProperty "ForeColor", m_OleForeColor, &H0
        .WriteProperty "MousePointer", m_MouMousePointer, 0
        .WriteProperty "Value", m_UChValue, uCheckboxConstants.u_unChecked
    End With


End Sub





