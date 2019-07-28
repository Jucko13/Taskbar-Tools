VERSION 5.00
Begin VB.UserControl uGraph 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0024211E&
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12015
   ForeColor       =   &H0000FF00&
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   801
End
Attribute VB_Name = "uGraph"
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

Private Type vLine
    lPoints() As Double
    lColor As Long
    lThickness As Long
    Visible As Boolean
End Type
    
Dim ScaleX As Double '1 = 1 punt per pixel
Dim ScaleY As Double '1 = 1 punt per pixel
Dim lHeight As Long
Dim lWidth As Long

Dim LineEveryX As Double
Dim LineEveryY As Double

Dim MaxY As Double
Dim MinY As Double
Dim Range As Double

Dim DragX As Double
Dim tmpDragX As Long

Dim MostItems As Long

Dim MessureRate As Long

Dim Lines(0 To 8) As vLine

Dim newItemAdded As Boolean 'alleen gaan scrollen als er een nieuw item is
Dim GrafiekEenheid As String
Dim Dragging As Boolean
Dim DraggingX As Double
Dim DragTmpX As Double

Dim mouseX As Double

Private unitNames() As String
Private Const unitNamesConst As String = ",k,m,t"
Private mouseMoveEvent As Boolean
Private Const offsetX As Long = 30

Private WithEvents m_tmrMouseOver As Timer
Attribute m_tmrMouseOver.VB_VarHelpID = -1
Private m_PoiMousePosition As POINTAPI
Private m_bMouseOver As Boolean
Public Event MouseEnter()
Public Event MouseLeave()


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


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    DraggingX = DragX
    DragTmpX = X
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mouseMoveEvent = True Then Exit Sub
    mouseMoveEvent = True
    
    If Dragging = True Then
        DragX = DraggingX + ((DragTmpX - X) * ScaleX)
        If ScaleX > 1 Then
            DragX = DragX - (DragX Mod ScaleX)
        End If
        If DragX < 0 Then DragX = 0
        If DragX > tmpDragX Then DragX = tmpDragX
        'Scroll.Value = 30000 / tmpDragX * DragX
        
    End If
    
    mouseX = X
    
    If m_bMouseOver = False Then
        m_bMouseOver = True
        m_tmrMouseOver.Interval = 40
        m_tmrMouseOver.Enabled = True
        RaiseEvent MouseEnter
    End If
        
    'lnLine.Y1 = 0
    'lnLine.Y2 = picGraph.Height
    'lnLine.X1 = x
    'lnLine.X2 = x
    

    'lnLine.Visible = True
    
    Redraw
    
    DoEvents
    mouseMoveEvent = False
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dragging = False
End Sub

'Private Sub Scroll_Change()
'    Dim i As Long
'    Static StillScrolling As Boolean
'
'    If StillScrolling Then Exit Sub
'    StillScrolling = True
'
'    If tmpDragX < 0 Then
'        DragX = 0
'    Else
'        DragX = Round((tmpDragX / 30000) * Scroll.Value, 0)
'        If ScaleX > 1 Then
'            DragX = DragX - (DragX Mod ScaleX)
'        End If
'    End If
'
'    'For i = 0 To UBound(Lines)
'    '    lblLine(i).Visible = False
'    'Next i
'    'lnLine.Visible = False
'
'    Redraw
'
'    StillScrolling = False
'End Sub

Private Sub UserControl_Initialize()
    Dim i As Long

    For i = 0 To UBound(Lines)
        ReDim Lines(i).lPoints(0) As Double
        Lines(i).lColor = vbRed
        Lines(i).lThickness = 1
        'If i > 0 Then
        '    Load lblLine(i)
        '    lblLine(i).Visible = False
        'End If
    Next i
    
    unitNames = Split(unitNamesConst, ",")
    
    LineEveryX = 10
    LineEveryY = 8
    
    Set m_tmrMouseOver = UserControl.Controls.Add("VB.Timer", "m_tmrMouseOver")
    
    MessureRate = 20
    MaxY = 1
    MinY = 0
    Range = (MaxY) - (MinY)
    ScaleX = 0.5
    GrafiekEenheid = "ms"
    UserControl_Resize
End Sub

Sub SetGrafiekEenheid(NieuweEenheid As String, Optional bRedraw As Boolean = False)
    GrafiekEenheid = NieuweEenheid
    
    If bRedraw Then
        Redraw
    End If
End Sub

Sub SetMaxY(dValue As Double, Optional bRedraw As Boolean = True)
    MaxY = dValue
    
    Range = (MaxY) - (MinY)
    
    If bRedraw Then
        Redraw
    End If
End Sub

Sub SetMinY(dValue As Double, Optional bRedraw As Boolean = True)
    MinY = dValue
    
    Range = (MaxY) - (MinY)
    
    If bRedraw Then
        Redraw
    End If
End Sub


Sub SetMessureRate(lValue As Long)
    MessureRate = 1000 / lValue
End Sub

Sub SetLineEveryNy(lValue As Long)
    LineEveryY = lValue
End Sub

Sub SetLineEveryNx(lValue As Long)
    LineEveryX = lValue
End Sub

Sub ScrollToLastItem(LineNumber As Long, Optional ForceScroll As Boolean = False)
    Dim tmpValue As Long
    If newItemAdded = False And ForceScroll = False Then Exit Sub
    'tmpValue = tmpDragX
    
    tmpDragX = (MostItems - Fix(lWidth * ScaleX))
    If tmpDragX < 0 Then tmpDragX = 0
    
    If tmpDragX < 0 Then
        DragX = 0
    Else
        DragX = tmpDragX 'Round((tmpDragX / 30000) * Scroll.Value, 0)
        If ScaleX > 1 Then
            DragX = DragX - (DragX Mod ScaleX)
        End If
    End If
    
    'If Scroll.Value <> Scroll.Max Then
    '    Scroll.Value = Scroll.Max
    'Else
    '    Scroll_Change
    'End If
    
    'tmpDragX = tmpValue
    newItemAdded = False
    
    Dim lEnd As Single
    
    If (lWidth * ScaleX) + DragX < UserControl.ScaleWidth - 1 Then
        lEnd = (lWidth * ScaleX)
        If lEnd > UBound(Lines(0).lPoints) Then
            lEnd = UBound(Lines(0).lPoints)
        End If
    
        'picGraph_MouseMove 0, 0, lEnd / ScaleX, 0
    Else
        'picGraph_MouseMove 0, 0, picGraph.Width, 0
    End If
    
End Sub

Sub Clear(Optional LineNumber As Long = -1)
    'ReDim Lines(LineNumber).lPoints(0) As Double
    Dim i As Long
    
    MostItems = 0
    
    If LineNumber = -1 Then
        For i = 0 To UBound(Lines)
            ReDim Lines(i).lPoints(0)
        Next i
        
    Else
        ReDim Lines(LineNumber).lPoints(0) As Double
        For i = 0 To UBound(Lines)
            If UBound(Lines(i).lPoints) > MostItems Then
                MostItems = UBound(Lines(i).lPoints)
            End If
        Next i
    End If
    
    Redraw
End Sub

Function AddItems(LineNumber As Long, Items() As Double, Optional bRedraw As Boolean = True)
    Dim i As Long
    
    For i = LBound(Items) To UBound(Items)
        AddItem LineNumber, Items(i), False
    Next i
    
    If bRedraw Then
        Redraw
    End If
End Function


Function AddItem(LineNumber As Long, ItemValue As Double, Optional bRedraw As Boolean = True)
    Dim NewId As Long
    NewId = -1
    newItemAdded = True
    With Lines(LineNumber)
        NewId = UBound(.lPoints) + 1
        ReDim Preserve .lPoints(0 To NewId) As Double
        .lPoints(NewId) = ItemValue
    End With
    
    If NewId > MostItems Then
        MostItems = NewId
        If (MostItems) > (lWidth * ScaleX) Then
            tmpDragX = ((MostItems) - (lWidth * ScaleX))
        Else
            tmpDragX = 0
        End If
    End If
    
    If ItemValue > MaxY Then
        MaxY = ItemValue
        Range = (MaxY) - (MinY)
        LineEveryY = Range
    End If
    
    If ItemValue < MinY Then
        MinY = ItemValue
        Range = (MaxY) - (MinY)
        LineEveryY = Range
    End If
    
    AddItem = NewId
    If bRedraw Then
        Redraw
    End If
End Function

Property Let LineColor(LineNumber As Long, LineColor As Long)
    Lines(LineNumber).lColor = LineColor
    'lblLine(LineNumber).ForeColor = LineColor
End Property

Property Let LineThickness(LineNumber As Long, LineThickness As Long)
    Lines(LineNumber).lThickness = LineThickness
End Property

Property Let LineVisible(LineNumber As Long, lVisible As Boolean)
    Lines(LineNumber).Visible = lVisible
End Property

Sub Redraw()
    Dim highestPoint As Long
    Dim tmpHighestPoint As Long
    
    lHeight = UserControl.ScaleHeight - 1
    lWidth = UserControl.ScaleWidth - offsetX
    ScaleY = lHeight / Range
    tmpDragX = ((MostItems) - (lWidth * ScaleX))
    If tmpDragX < 0 Then tmpDragX = 0
    

    
    
    UserControl.Picture = LoadPicture()
    
    DrawGrid
    Dim i As Long
    For i = 0 To UBound(Lines)
        If Lines(i).Visible = True Then
            'If i = 5 Then
            '    DrawPoints Lines(i), True
            'Else
                
                tmpHighestPoint = DrawPoints(Lines(i), False)
                If tmpHighestPoint > highestPoint Then
                    highestPoint = tmpHighestPoint
                End If
                
            'End If
        End If
    Next i
    

    DrawLine
    
    'DoEvents
    
    If highestPoint <> MaxY And highestPoint > 0 Then
        MaxY = highestPoint
        Range = MaxY
    End If
    
    'lnLine.Redraw
End Sub

Sub DrawLine()
    
    Dim i As Long
    Dim tmpLeft As Long
    Dim tmpTop As Long
    Dim tmptest As Long
    
    
    If Not m_bMouseOver Then
        mouseX = offsetX + MostItems / ScaleX - DragX / ScaleX
    End If
    
    UserControl.Line (mouseX, 0)-(mouseX, UserControl.ScaleHeight), vbBlue
    
    tmptest = DragX + ((mouseX - offsetX) * ScaleX)
    
    'Debug.Print Lines(1).lPoints(tmptest)
    
    If tmptest < 0 Then tmptest = 0
    

    'On Error Resume Next
    
    For i = 0 To UBound(Lines)
        If Lines(i).Visible = True Then
            Dim tmpLineText As String
            Dim tmpLineTextWidth As Long
            Dim tmpLineTextHeight As Long
            
            
            If tmptest <= UBound(Lines(i).lPoints) Then
                tmpLineText = Round(Lines(i).lPoints(tmptest)) & " Bps" 'Format$(Format$(Lines(i).lPoints(tmptest), "0"), "@@@") & " Bps"
            Else
                tmpLineText = "0 Bps"
            End If
            
            tmpLineTextWidth = UserControl.TextWidth(tmpLineText)
            tmpLineTextHeight = UserControl.TextHeight(tmpLineText)
            
            If mouseX + tmpLineTextWidth < UserControl.ScaleWidth - 5 Then
                tmpLeft = mouseX + 4 ' lblLine(i).Width
            Else
                tmpLeft = (mouseX - tmpLineTextWidth) - 4
            End If
            
            
'            ReDim pts(0 To 3)
'            pts(0).x = tmpLeft
'            pts(0).y = tmpTop
'
'            pts(1).x = pts(0).x + tmpLineTextWidth
'            pts(1).y = pts(0).y
'
'            pts(2).x = pts(1).x
'            pts(2).y = tmpLineTextHeight + pts(0).y
'
'            pts(3).x = pts(0).x
'            pts(3).y = pts(2).y
'
'            picGraph.ForeColor = Lines(i).lColor
'
'            picGraph.DrawStyle = 5
'            Polygon picGraph.hdc, pts(0), 4
'            picGraph.DrawStyle = 0
            
            Dim X As Long
            Dim Y As Long
'
            UserControl.ForeColor = vbBlack
            For X = -1 To 1
                For Y = -1 To 1
                    If Y <> 0 Or X <> 0 Then
                        UserControl.CurrentX = tmpLeft + X
                        UserControl.CurrentY = tmpTop + Y
                        UserControl.Print tmpLineText;
                    End If
                Next Y
            Next X

            UserControl.ForeColor = Lines(i).lColor

            UserControl.CurrentX = tmpLeft
            UserControl.CurrentY = tmpTop
            UserControl.Print tmpLineText;
                       
            UserControl.ForeColor = vbBlack
            
            tmpTop = tmpTop + tmpLineTextHeight + 5
        End If
    Next i
    
End Sub

Function getShortName(ByVal valueMS As Double) As String
    Dim unitNr As Long
    
    While valueMS > 1000
        valueMS = valueMS / 1000
        unitNr = unitNr + 1
    Wend
    
    getShortName = Round(valueMS, 1) & unitNames(unitNr)
End Function

Sub DrawGrid()
    Dim i As Double
    Dim LineColor As Long
    
    Dim Verschuiving As Double
    Dim Zero As Double
    
    Dim CharLeft As Long
    Dim Char As String
    Dim txtHeight As Long
    Dim tmpLineHeight As Long
    Dim tmpLineDown As Long
    Dim tmpLineUp As Long
    Dim tmpColor As Long
    
    tmpColor = RGB(100, 100, 100) '&H24211E   ' picGraph.ForeColor
    UserControl.ForeColor = vbGreen
    
    'txtHeight = (UserControl.ScaleHeight) - (UserControl.TextHeight("H") * 1.5)
    Verschuiving = (LineEveryX) - ((DragX / ScaleX) Mod (LineEveryX / ScaleX))
    
    For i = (-LineEveryX + (Verschuiving)) To lWidth + LineEveryX / ScaleX Step (LineEveryX / ScaleX)
        If i > 0 Then
            UserControl.Line (i + offsetX, 0)-(i + offsetX, lHeight), tmpColor
        End If
        'Char = CStr((DragX + i * ScaleX) / MessureRate)
        'UserControl.CurrentY = txtHeight
        'UserControl.CurrentX = i - (UserControl.TextWidth(Char) / 2) + picGraph.Left
        'If UserControl.CurrentX > picGraph.Left - (UserControl.TextWidth(Char)) Then
        '    UserControl.Print Char
        'End If
    Next i
    
    CharLeft = offsetX - 10
    'txtHeight = '(UserControl.TextHeight("H") / 2) - (picGraph.Top)
    Zero = (Range - MaxY)
    
    Char = "0"
    
    'picGraph.Line (0, picGraph.ScaleHeight - 1)-(lWidth, picGraph.ScaleHeight - 1), tmpColor
    'picGraph.Line (0, 0)-(lWidth, 0), tmpColor
    
    UserControl.CurrentY = UserControl.ScaleHeight - UserControl.TextHeight(Char) '+ picGraph.Top - UserControl.TextHeight(Char)
    UserControl.CurrentX = 4 'CharLeft - UserControl.TextWidth(Char)
    UserControl.Print Char;
    
    Char = getShortName(MaxY)
    
    UserControl.CurrentY = 0 'picGraph.Height + picGraph.Top - UserControl.TextHeight(Char)
    UserControl.CurrentX = 4 'CharLeft - UserControl.TextWidth(Char)
    UserControl.Print Char;
    
    
'    For i = 0 To Range Step LineEveryY
'
'        If i = 0 Then
'            LineColor = vbBlue
'        Else
'            LineColor = tmpColor
'        End If
'
'        tmpLineHeight = lHeight - (ScaleY * Zero)
'        tmpLineDown = tmpLineHeight + (ScaleY * i)
'        tmpLineUp = tmpLineHeight - (ScaleY * i)
'
'
'
'        If tmpLineDown <= lHeight Then
'            If tmpLineDown < lHeight Then
'                picGraph.Line (0, tmpLineDown)-(lWidth, tmpLineDown), LineColor
'            End If
'
'            'Char = CStr(-i)
'            'If i = 0 Then
'            '    Char = Char & " " & GrafiekEenheid
'            '    UserControl.FontBold = True
'            'Else
'            '    UserControl.FontBold = False
'            'End If
'            'UserControl.CurrentY = tmpLineDown - txtHeight
'            'UserControl.CurrentX = CharLeft - UserControl.TextWidth(Char)
'            'UserControl.Print Char
'            'If i = 0 Then
'            '    UserControl.FontBold = False
'            'End If
'        End If
'
'       ' If tmpLineUp >= 0 Then
'        '    If tmpLineUp > 0 Then
'       '         picGraph.Line (0, tmpLineUp)-(lWidth, tmpLineUp), LineColor
'        '    End If
''            Char = CStr(i)
''            If i = 0 Then
''                Char = Char & " " & GrafiekEenheid
''                UserControl.FontBold = True
''            End If
''            UserControl.CurrentY = tmpLineUp - txtHeight + 5
''            UserControl.CurrentX = CharLeft - UserControl.TextWidth(Char)
''            UserControl.Print Char
''            If i = 0 Then
''                UserControl.FontBold = False
''            End If
'        'End If
'
'        'If i > 0 And i < Range Then
'        '    tmpLineHeight = lHeight - (ScaleY * i) - 1
'        '    picGraph.Line (0, tmpLineHeight)-(lWidth, tmpLineHeight), LineColor
'        'End If
'
'
'    Next i
    
    

    
    
End Sub

Function ReturnCoords(LineNumber As Long) As Double()

    ReturnCoords = Lines(LineNumber).lPoints
End Function

Private Function DrawPoints(ByRef LineDraw As vLine, Optional Test As Boolean = False) As Long
    Dim i As Long
    
    Dim PrevPointX As Double
    Dim PrevPointY As Double
    
    Dim tmpStep As Double
    Dim tmpX As Double
    Dim tmpY As Double
    
    Dim lStart As Long
    Dim lEnd As Long
    
    Dim tmpLineThickness As Long
    
    Dim j As Long
'    Dim GemGestegen As Single
'    Dim GemetenPiek As Single
'    Dim MagGaanMeten As Boolean
'    Dim isOmhoogGeweest As Boolean
    
    tmpLineThickness = UserControl.DrawWidth
    UserControl.DrawWidth = LineDraw.lThickness
    
    lStart = DragX
    lEnd = (lWidth * ScaleX) + lStart
    If lEnd > UBound(LineDraw.lPoints) Then
        lEnd = UBound(LineDraw.lPoints)
    End If
    
    If lStart < 0 Then lStart = 0
    
    PrevPointX = 0
    PrevPointY = 0
    tmpStep = 0 ' * ScaleX
    For i = lStart To lEnd
        If LineDraw.lPoints(i) >= 0 Then
            tmpY = lHeight - ((LineDraw.lPoints(i) + Abs(MinY)) * ScaleY)
        Else
            tmpY = lHeight - ((LineDraw.lPoints(i) + Abs(MinY)) * ScaleY)
        End If
        
        tmpX = tmpStep / ScaleX
        
        If i = lStart Then
            PrevPointX = tmpX
            PrevPointY = tmpY
        End If
        
        If LineDraw.lPoints(i) > DrawPoints Then
             DrawPoints = LineDraw.lPoints(i)
        End If
        
'        If Test = True Then
'            If i > 3 Then
'                GemGestegen = (Lines(0).lPoints(i) + Lines(0).lPoints(i - 1) + Lines(0).lPoints(i - 2) + Lines(0).lPoints(i - 3)) / 3
'                'GemGestegen = (Lines(0).lPoints(i) - Lines(0).lPoints(i - 1)) + (Lines(0).lPoints(i - 1) - Lines(0).lPoints(i - 2))
'
'                If GemGestegen > Lines(0).lPoints(i) And GemGestegen < Lines(0).lPoints(i - 3) Then
'                    LineDraw.lColor = vbCyan 'dalen
'                    GemetenPiek = 0
'                    If isOmhoogGeweest = True Then
'                        MagGaanMeten = True
'                        isOmhoogGeweest = False
'                    End If
'                ElseIf GemGestegen < Lines(0).lPoints(i) And GemGestegen > Lines(0).lPoints(i - 3) Then
'                    LineDraw.lColor = vbCyan 'stijgen
'                    GemetenPiek = 1
'                    isOmhoogGeweest = True
'                    MagGaanMeten = False
'                Else
'                    LineDraw.lColor = vbCyan 'piek
'
'                    If Abs(LineDraw.lPoints(i)) > GemetenPiek And MagGaanMeten = True Then
'                        GemetenPiek = Abs(LineDraw.lPoints(i))
'                        LineDraw.lColor = vbMagenta       'piek
'                    ElseIf MagGaanMeten = True Then
'                        LineDraw.lColor = &H8080FF
'                    End If
'                End If
'
'
'            End If
'
'
'        End If
        
        UserControl.Line (PrevPointX + offsetX, PrevPointY)-(tmpX + offsetX, tmpY), LineDraw.lColor
        
        PrevPointX = tmpX
        PrevPointY = tmpY
        tmpStep = tmpStep + 1
        
    Next i

    UserControl.DrawWidth = tmpLineThickness
    
End Function

Function getPoint(LineNumber As Long, ItemNumber As Long) As Double
    getPoint = Lines(LineNumber).lPoints(ItemNumber)
End Function

Function getUbound(LineNumber As Long) As Long
    getUbound = UBound(Lines(LineNumber).lPoints)
End Function

Private Sub UserControl_Resize()
    On Error Resume Next
    'picGraph.Top = 5
    'picGraph.Height = UserControl.ScaleHeight - picGraph.Top * 2 ' - (Scroll.Height * 2)
    'Scroll.Top = UserControl.ScaleHeight - Scroll.Height - 1
    'picGraph.Width = UserControl.ScaleWidth - picGraph.Left
    'Scroll.Width = picGraph.Width
    'Scroll.Left = picGraph.Left
    
    mouseX = UserControl.ScaleWidth - 1 ' / Screen.TwipsPerPixelX
    
    'lblInfo.Top = picGraph.Top + (picGraph.Height / 2) - (lblInfo.Height / 2)
    
    Redraw
End Sub
