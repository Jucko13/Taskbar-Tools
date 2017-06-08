VERSION 5.00
Begin VB.UserControl uLoadBar 
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
Attribute VB_Name = "uLoadBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum uCaptionType
    u_Percentage = 0
    u_Value = 1
    u_Custom = 99
End Enum

Public Enum uBarType
    u_Bar_Horizontal = 0
    u_Bar_Vertical = 1
    u_Bar_Square = 2
End Enum

Private WithEvents m_tmrLoad As Timer
Attribute m_tmrLoad.VB_VarHelpID = -1

Private m_uCaptionTypeStyle As uCaptionType
Private m_StrCaption As String

Private m_bBorder As Boolean
Private m_bStarting As Boolean
Private m_MouMousePointer As MousePointerConstants
Private m_bRefreshing As Boolean

Private m_OleBackgroundColor As OLE_COLOR
Private m_OleForeColor As OLE_COLOR
Private m_OleBorderColor As OLE_COLOR
Private m_OleBarColor As OLE_COLOR

Private m_LonMinValue As Long
Private m_LonMaxValue As Long
Private m_LonValue As Long
Private m_LonBarWidth As Long
Private m_bLoading As Boolean
Private m_DouLoadPosition As Double
Private m_bCaptionBorder As Boolean
Private m_OleCaptionBorderColor As OLE_COLOR
Private m_LonLoadingSpeed As Long
Private m_StdFont As StdFont
Private m_UBaBarType As uBarType

Private Const Pi As Double = 3.14159265359


Public Property Get BarType() As uBarType
    BarType = m_UBaBarType
End Property

Public Property Let BarType(ByVal UBaValue As uBarType)
    m_UBaBarType = UBaValue
    PropertyChanged "BarType"
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


Public Property Get CaptionBorderColor() As OLE_COLOR
    CaptionBorderColor = m_OleCaptionBorderColor
End Property

Public Property Let CaptionBorderColor(ByVal OleValue As OLE_COLOR)
    m_OleCaptionBorderColor = OleValue
    PropertyChanged "CaptionBorderColor"
    If Not m_bStarting Then Redraw
End Property

Public Property Get LoadingSpeed() As Long
    LoadingSpeed = m_LonLoadingSpeed
End Property

Public Property Let LoadingSpeed(ByVal LonValue As Long)
    m_LonLoadingSpeed = LonValue
    If m_LonLoadingSpeed < 0 Then
        m_LonLoadingSpeed = 0
    ElseIf m_LonLoadingSpeed > 10 Then
        m_LonLoadingSpeed = 10
    End If

    PropertyChanged "LoadingSpeed"
End Property




Public Property Get CaptionBorder() As Boolean
    CaptionBorder = m_bCaptionBorder
End Property

Public Property Let CaptionBorder(ByVal bValue As Boolean)
    m_bCaptionBorder = bValue
    PropertyChanged "CaptionBorder"
    If Not m_bStarting Then Redraw
End Property



Public Property Get CaptionType() As uCaptionType
    CaptionType = m_uCaptionTypeStyle
End Property

Public Property Let CaptionType(ByVal uCaptionTypeValue As uCaptionType)
    m_uCaptionTypeStyle = uCaptionTypeValue
    PropertyChanged "CaptionType"
    If Not m_bStarting Then Redraw
End Property



Public Property Get Loading() As Boolean
    Loading = m_bLoading
End Property

Public Property Let Loading(ByVal bValue As Boolean)
    m_bLoading = bValue
    m_tmrLoad.Enabled = bValue
    PropertyChanged "Loading"
    If Not m_bStarting Then Redraw
End Property


Public Property Get BarWidth() As Long
    BarWidth = m_LonBarWidth
End Property

Public Property Let BarWidth(ByVal LonValue As Long)
    m_LonBarWidth = LonValue
    If m_LonBarWidth < 0 Then
        m_LonBarWidth = 0
    End If

    PropertyChanged "BarWidth"
    If Not m_bStarting Then Redraw
End Property

Public Property Get Value() As Long
    Value = m_LonValue
End Property

Public Property Let Value(ByVal LonValue As Long)
    m_LonValue = LonValue
    If m_LonValue < m_LonMinValue Then
        m_LonValue = m_LonMinValue
    End If

    If m_LonValue > m_LonMaxValue Then
        m_LonValue = m_LonMaxValue
    End If

    PropertyChanged "Value"
    If Not m_bStarting Then Redraw
End Property

Public Property Get MaxValue() As Long
    MaxValue = m_LonMaxValue
End Property

Public Property Let MaxValue(ByVal LonValue As Long)
    m_LonMaxValue = LonValue

    If m_LonMaxValue < 0 Then
        m_LonMaxValue = 0
    ElseIf m_LonMaxValue < m_LonMinValue Then
        m_LonMinValue = m_LonMaxValue
        PropertyChanged "MinValue"
    End If


    If m_LonValue > m_LonMaxValue Then
        m_LonValue = m_LonMaxValue
        PropertyChanged "Value"
    End If

    PropertyChanged "MaxValue"
    If Not m_bStarting Then Redraw
End Property

Public Property Get MinValue() As Long
    MinValue = m_LonMinValue
End Property

Public Property Let MinValue(ByVal LonValue As Long)
    m_LonMinValue = LonValue
    If m_LonMinValue < 0 Then
        m_LonMinValue = 0
    ElseIf m_LonMinValue > m_LonMaxValue Then
        m_LonMaxValue = m_LonMinValue
        PropertyChanged "MaxValue"
    End If

    If m_LonValue < m_LonMinValue Then
        m_LonValue = m_LonMinValue
        PropertyChanged "Value"
    End If

    PropertyChanged "MinValue"
    If Not m_bStarting Then Redraw
End Property

Public Property Get BarColor() As OLE_COLOR
    BarColor = m_OleBarColor
End Property

Public Property Let BarColor(ByVal OleValue As OLE_COLOR)
    m_OleBarColor = OleValue
    PropertyChanged "BarColor"
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
Attribute BackgroundColor.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute BackgroundColor.VB_MemberFlags = "4"
    BackgroundColor = m_OleBackgroundColor
End Property

Public Property Let BackgroundColor(ByVal OleValue As OLE_COLOR)
    m_OleBackgroundColor = OleValue
    PropertyChanged "BackgroundColor"
    If Not m_bStarting Then Redraw
End Property

Private Sub m_tmrLoad_Timer()
    If m_UBaBarType = u_Bar_Square Then
        m_DouLoadPosition = m_DouLoadPosition + (Pi * (m_LonLoadingSpeed / 400))
        If m_DouLoadPosition > 2 * Pi Then
            m_DouLoadPosition = 0
        End If

    Else
        m_DouLoadPosition = m_DouLoadPosition + m_LonLoadingSpeed
        If m_DouLoadPosition > 61 Then
            m_DouLoadPosition = (CLng(m_DouLoadPosition) Mod 61)
        End If

    End If

    Redraw
End Sub

Private Sub UserControl_Initialize()
    Set m_tmrLoad = UserControl.Controls.Add("VB.Timer", "m_tmrLoad")
    m_tmrLoad.Enabled = False
    m_tmrLoad.Interval = 30

    m_bStarting = True
    m_OleForeColor = &H800000
    m_StrCaption = ""
    m_OleBorderColor = &HFFFFFF
    m_OleBackgroundColor = &HE18700
    m_OleBarColor = &HFFD091
    m_bBorder = True

    m_MouMousePointer = 0

    m_LonMinValue = 0
    m_LonMaxValue = 100
    m_LonValue = 50
    m_LonBarWidth = 8

    m_uCaptionTypeStyle = 0
    m_bLoading = False
    m_LonLoadingSpeed = 2
    m_bCaptionBorder = False


    m_bStarting = False


    m_OleCaptionBorderColor = &HFFFFFF
    m_UBaBarType = uBarType.u_Bar_Square
End Sub


Sub Redraw()
    If m_bRefreshing Then Exit Sub
    m_bRefreshing = True
    Dim i As Long
    Dim j As Long

    Dim tmpTextHeight As Long
    Dim tmpTextWidth As Long
    Dim tmpX As Single
    Dim tmpY As Single
    Dim tmpConnectX As Single
    Dim tmpConnectY As Single
    Dim bToggle As Boolean
    Dim lCounter As Long
    Dim tmpPi As Double
    Dim partSize As Double
    
    If Not m_StdFont Is Nothing Then
        Set UserControl.Font = m_StdFont
    End If

    tmpTextHeight = UserControl.TextHeight(m_StrCaption)
    tmpTextWidth = UserControl.TextWidth(m_StrCaption)

    UserControl.Cls
    UserControl.BackColor = m_OleBackgroundColor
    UserControl.ForeColor = m_OleForeColor

    On Error GoTo Repair_Array
    If UBound(pts) < 41 Then
Repair_Array:
        ReDim pts(0 To 41) As POINTAPI
    End If
    On Error GoTo 0


    If m_LonMaxValue <= 0 Then m_LonMaxValue = m_LonMinValue + 1

    If m_bLoading Then GoTo Skip_Plotting


Set_Progress:
    If m_LonValue = m_LonMinValue Then GoTo Skip_Plotting


    partSize = m_LonMaxValue / 8

    pts(0).X = UserControl.ScaleWidth / 2
    pts(0).Y = -3

    UserControl.DrawWidth = 1
    UserControl.FillStyle = 0
    UserControl.FillColor = m_OleBarColor
    UserControl.ForeColor = IIf(m_bBorder = True, m_OleBorderColor, m_OleBackgroundColor)

    Select Case m_UBaBarType
        Case uBarType.u_Bar_Square
            If m_LonValue <= partSize Then    ' 12.5% or smaller
                tmpX = Tan((2 * Pi) / m_LonMaxValue * m_LonValue) * (UserControl.ScaleWidth / 2 + 3)
                pts(1).X = pts(0).X + tmpX: pts(1).Y = -3
                pts(2).X = UserControl.ScaleWidth / 2: pts(2).Y = UserControl.ScaleHeight / 2
                Polygon UserControl.hdc, pts(0), 3
            ElseIf m_LonValue <= (partSize * 3) Then
                pts(1).X = pts(0).X + (UserControl.ScaleWidth / 2) + 3: pts(1).Y = -3
                tmpY = Tan((2 * Pi) / m_LonMaxValue * (m_LonValue - partSize * 2)) * (UserControl.ScaleHeight / 2 + 3)
                pts(2).X = pts(1).X: pts(2).Y = UserControl.ScaleHeight / 2 + tmpY
                pts(3).X = UserControl.ScaleWidth / 2: pts(3).Y = UserControl.ScaleHeight / 2
                Polygon UserControl.hdc, pts(0), 4
            ElseIf m_LonValue <= (partSize * 5) Then
                pts(1).X = UserControl.ScaleWidth + 3: pts(1).Y = -3
                pts(2).X = pts(1).X: pts(2).Y = UserControl.ScaleHeight + 3
                tmpX = Tan((2 * Pi) / m_LonMaxValue * (m_LonValue - partSize * 4)) * ((UserControl.ScaleWidth + 6) / 2)
                pts(3).X = UserControl.ScaleWidth / 2 - tmpX: pts(3).Y = pts(2).Y
                pts(4).X = UserControl.ScaleWidth / 2: pts(4).Y = UserControl.ScaleHeight / 2
                Polygon UserControl.hdc, pts(0), 5
            ElseIf m_LonValue <= (partSize * 7) Then
                pts(1).X = UserControl.ScaleWidth + 3: pts(1).Y = -3
                pts(2).X = pts(1).X: pts(2).Y = UserControl.ScaleHeight + 3
                pts(3).X = -3: pts(3).Y = UserControl.ScaleHeight + 3
                tmpY = Tan((2 * Pi) / m_LonMaxValue * (m_LonValue - partSize * 6)) * ((UserControl.ScaleHeight + 6) / 2)
                pts(4).X = pts(3).X: pts(4).Y = UserControl.ScaleHeight / 2 - tmpY
                pts(5).X = UserControl.ScaleWidth / 2: pts(5).Y = UserControl.ScaleHeight / 2
                Polygon UserControl.hdc, pts(0), 6
            ElseIf m_LonValue = m_LonMaxValue Then
                pts(0).X = -3: pts(0).Y = -3
                pts(1).X = UserControl.ScaleWidth + 3: pts(1).Y = -3
                pts(2).X = UserControl.ScaleWidth + 3: pts(2).Y = UserControl.ScaleHeight + 3
                pts(3).X = -3: pts(3).Y = UserControl.ScaleHeight + 3
                Polygon UserControl.hdc, pts(0), 4
            Else
                pts(1).X = UserControl.ScaleWidth + 3: pts(1).Y = -3
                pts(2).X = pts(1).X: pts(2).Y = UserControl.ScaleHeight + 3
                pts(3).X = -3: pts(3).Y = UserControl.ScaleHeight + 3
                pts(4).X = -3: pts(4).Y = UserControl.ScaleHeight + 3
                pts(5).X = -3: pts(5).Y = -3
                tmpX = Tan((2 * Pi) / m_LonMaxValue * (m_LonValue - partSize * 7)) * ((UserControl.ScaleWidth) / 2)
                pts(6).X = tmpX: pts(6).Y = -3
                pts(7).X = UserControl.ScaleWidth / 2: pts(7).Y = UserControl.ScaleHeight / 2
                Polygon UserControl.hdc, pts(0), 8
            End If


        Case uBarType.u_Bar_Horizontal
            pts(0).X = -3: pts(0).Y = pts(0).X
            pts(1).X = (UserControl.ScaleWidth) / m_LonMaxValue * m_LonValue: pts(1).Y = pts(0).X
            pts(2).X = pts(1).X: pts(2).Y = UserControl.ScaleHeight + 3
            pts(3).X = pts(0).X: pts(3).Y = pts(2).Y
            Polygon UserControl.hdc, pts(0), 4


        Case uBarType.u_Bar_Vertical
            pts(0).X = -3: pts(0).Y = UserControl.ScaleHeight + 3
            pts(1).X = pts(0).X: pts(1).Y = UserControl.ScaleHeight - (UserControl.ScaleHeight) / m_LonMaxValue * m_LonValue - 1
            pts(2).X = UserControl.ScaleWidth + 3: pts(2).Y = pts(1).Y
            pts(3).X = pts(2).X: pts(3).Y = pts(0).Y
            Polygon UserControl.hdc, pts(0), 4


    End Select

Skip_Plotting:
    If m_bLoading Then

        Select Case m_UBaBarType
            Case uBarType.u_Bar_Square
                tmpX = Round(UserControl.ScaleWidth / 2)
                tmpY = Round(UserControl.ScaleHeight / 2)

                UserControl.ForeColor = IIf(m_bBorder = True, m_OleBorderColor, m_OleBackgroundColor)
                UserControl.FillColor = m_OleBarColor
                UserControl.FillStyle = 0

                tmpPi = (2 * Pi) / 16
                For i = 0 To 16
                    tmpConnectX = Cos(tmpPi * i + m_DouLoadPosition) * UserControl.ScaleWidth
                    tmpConnectY = Sin(tmpPi * i + m_DouLoadPosition) * UserControl.ScaleHeight

                    If Not bToggle Then
                        pts(lCounter).X = tmpX
                        pts(lCounter).Y = tmpY
                        lCounter = lCounter + 1

                        pts(lCounter).X = tmpConnectX + tmpX
                        pts(lCounter).Y = tmpConnectY + tmpY
                        lCounter = lCounter + 1
                    Else
                        pts(lCounter).X = tmpConnectX + tmpX
                        pts(lCounter).Y = tmpConnectY + tmpY
                        lCounter = lCounter + 1

                        pts(lCounter).X = tmpX
                        pts(lCounter).Y = tmpY
                        lCounter = lCounter + 1
                    End If
                    bToggle = Not bToggle

                Next i

                Polygon UserControl.hdc, pts(0), lCounter

            Case uBarType.u_Bar_Horizontal
                lCounter = 0

                tmpConnectX = -3 + UserControl.ScaleWidth + 6
                tmpConnectY = ((UserControl.ScaleWidth + 3) / 30) + 1

                UserControl.ForeColor = IIf(m_bBorder = True, m_OleBorderColor, m_OleBackgroundColor)
                UserControl.FillColor = m_OleBarColor
                UserControl.FillStyle = 0

                If UBound(pts) < (tmpConnectY + 2) * 5 Then
                    ReDim pts(0 To (tmpConnectY + 2) * 5) As POINTAPI
                End If
    
                For i = -tmpConnectY To 1

                    pts(lCounter).X = -3 + tmpConnectX + (60 * i) + m_DouLoadPosition
                    pts(lCounter).Y = UserControl.ScaleHeight + 3
                    lCounter = lCounter + 1

                    pts(lCounter).X = pts(lCounter - 1).X + 30
                    pts(lCounter).Y = pts(lCounter - 1).Y
                    lCounter = lCounter + 1

                    pts(lCounter).X = pts(lCounter - 1).X + UserControl.ScaleHeight + 6
                    pts(lCounter).Y = -3
                    lCounter = lCounter + 1


                    pts(lCounter).X = pts(lCounter - 1).X + 30
                    pts(lCounter).Y = pts(lCounter - 1).Y
                    lCounter = lCounter + 1


                    pts(lCounter).X = pts(lCounter - 3).X + 30
                    pts(lCounter).Y = pts(lCounter - 4).Y
                    lCounter = lCounter + 1
                Next i

                Polygon UserControl.hdc, pts(0), lCounter

            Case uBarType.u_Bar_Vertical

        End Select
    End If



    If m_bBorder Then
        UserControl.FillStyle = 1
        UserControl.ForeColor = m_OleBorderColor

        pts(0).X = 0
        pts(0).Y = 0

        pts(1).X = UserControl.ScaleWidth - 1
        pts(1).Y = 0

        pts(2).X = UserControl.ScaleWidth - 1
        pts(2).Y = UserControl.ScaleHeight - 1

        pts(3).X = 0
        pts(3).Y = UserControl.ScaleHeight - 1

        Polygon UserControl.hdc, pts(0), 4
    End If



    UserControl.FillStyle = 0
    UserControl.ForeColor = IIf(m_bBorder = True, m_OleBorderColor, m_OleBackgroundColor)
    UserControl.FillColor = m_OleBackgroundColor

    If m_LonBarWidth > 0 And UserControl.ScaleWidth > (m_LonBarWidth + 1) * 2 And UserControl.ScaleHeight > (m_LonBarWidth + 1) * 2 Then
        pts(0).X = m_LonBarWidth
        pts(0).Y = m_LonBarWidth

        pts(1).X = UserControl.ScaleWidth - m_LonBarWidth - 1
        pts(1).Y = m_LonBarWidth

        pts(2).X = UserControl.ScaleWidth - m_LonBarWidth - 1
        pts(2).Y = UserControl.ScaleHeight - m_LonBarWidth - 1

        pts(3).X = m_LonBarWidth
        pts(3).Y = UserControl.ScaleHeight - m_LonBarWidth - 1

        Polygon UserControl.hdc, pts(0), 4
    End If




    Dim tmpCaptionToPrint As String
    tmpCaptionToPrint = ""

    Select Case m_uCaptionTypeStyle

        Case uCaptionType.u_Percentage
            tmpCaptionToPrint = Round(100 / m_LonMaxValue * m_LonValue) & " %"

        Case uCaptionType.u_Value
            tmpCaptionToPrint = m_LonValue & "/" & m_LonMaxValue

        Case uCaptionType.u_Custom
            tmpCaptionToPrint = m_StrCaption


    End Select


    tmpX = Round(UserControl.ScaleWidth / 2 - UserControl.TextWidth(tmpCaptionToPrint) / 2)
    tmpY = Round(UserControl.ScaleHeight / 2 - UserControl.TextHeight(tmpCaptionToPrint) / 2)

    If m_bCaptionBorder And tmpCaptionToPrint <> "" Then
        UserControl.ForeColor = m_OleCaptionBorderColor

        For i = -1 To 1
            For j = -1 To 1
                If Not (i = 0 And j = 0) Then
                    UserControl.CurrentX = tmpX + i
                    UserControl.CurrentY = tmpY + j
                    UserControl.Print tmpCaptionToPrint
                End If
            Next j
        Next i
    End If

    UserControl.ForeColor = m_OleForeColor
    UserControl.CurrentX = tmpX
    UserControl.CurrentY = tmpY
    UserControl.Print tmpCaptionToPrint


End_of_sub:
    m_bRefreshing = False
End Sub


Private Sub Usercontrol_Resize()
    If Not m_bStarting Then Redraw
End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_bStarting = True
    With PropBag
        m_OleBackgroundColor = .ReadProperty("BackgroundColor", &HE18700)
        m_OleBarColor = .ReadProperty("BarColor", &HFFD091)
        m_UBaBarType = .ReadProperty("BarType", uBarType.u_Bar_Square)
        m_LonBarWidth = .ReadProperty("BarWidth", 8)
        m_bBorder = .ReadProperty("Border", True)
        m_OleBorderColor = .ReadProperty("BorderColor", &HFFFFFF)
        m_StrCaption = .ReadProperty("Caption", "Button")
        m_bCaptionBorder = .ReadProperty("CaptionBorder", False)
        m_OleCaptionBorderColor = .ReadProperty("CaptionBorderColor", &HFFFFFF)
        m_uCaptionTypeStyle = .ReadProperty("CaptionType", uCaptionType.u_Percentage)
        Set Font = .ReadProperty("Font", Ambient.Font)
        m_OleForeColor = .ReadProperty("ForeColor", &H800000)
        Loading = .ReadProperty("Loading", False)
        m_LonLoadingSpeed = .ReadProperty("LoadingSpeed", 2)
        m_LonMaxValue = .ReadProperty("MaxValue", 100)
        m_LonMinValue = .ReadProperty("MinValue", 0)
        MousePointer = .ReadProperty("MousePointer", 0)
        m_LonValue = .ReadProperty("Value", 50)
    End With
    m_bStarting = False
    Redraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackgroundColor", m_OleBackgroundColor, &HE18700
        .WriteProperty "BarColor", m_OleBarColor, &HFFD091
        .WriteProperty "BarType", m_UBaBarType, uBarType.u_Bar_Square
        .WriteProperty "BarWidth", m_LonBarWidth, 8
        .WriteProperty "Border", m_bBorder, True
        .WriteProperty "BorderColor", m_OleBorderColor, &HFFFFFF
        .WriteProperty "Caption", m_StrCaption, "Button"
        .WriteProperty "CaptionBorder", m_bCaptionBorder, False
        .WriteProperty "CaptionBorderColor", m_OleCaptionBorderColor, &HFFFFFF
        .WriteProperty "CaptionType", m_uCaptionTypeStyle, uCaptionType.u_Percentage
        .WriteProperty "Font", m_StdFont, Ambient.Font
        .WriteProperty "ForeColor", m_OleForeColor, &H800000
        .WriteProperty "Loading", m_bLoading, False
        .WriteProperty "LoadingSpeed", m_LonLoadingSpeed, 2
        .WriteProperty "MaxValue", m_LonMaxValue, 100
        .WriteProperty "MinValue", m_LonMinValue, 0
        .WriteProperty "MousePointer", m_MouMousePointer, 0
        .WriteProperty "Value", m_LonValue, 50
    End With

End Sub





