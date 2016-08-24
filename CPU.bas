Attribute VB_Name = "CPU"
''''Sub ReCalculateWords()
''''    Dim WC As Long 'word count
''''    Dim WH As Long 'word height
''''    Dim WW As Long 'word width
''''    Dim WL As Long 'word length
''''
''''    Dim TL As Long 'text length
''''    'Dim MS As String 'mid string
''''
''''    'TLength = Len(m_StrText)
''''
''''    ReDim WordMap(0 To UBound(m_byteText) + 2)
''''
''''    'WordMap(0).S = 0
''''
''''    For TL = 0 To UBound(m_byteText)
''''        'MSL = Asc(Mid$(m_StrText, TL + 1, 1))
''''
''''        If m_byteText(TL) = 32 Or m_byteText(TL) = 13 Or m_byteText(TL) = 10 Or m_byteText(TL) = 45 Then ' a space
''''            If WL > 0 Then
''''                WordMap(WC).H = WH
''''                WordMap(WC).W = WW
''''                WordMap(WC).L = WL
''''                WC = WC + 1
''''                WH = 0
''''                WW = 0
''''                WL = 0
''''
''''                WordMap(WC).S = TL + 1
''''                MarkupS(TL).lPartOfWord = -1
''''            End If
''''        Else
''''            MarkupS(TL).lPartOfWord = WC
''''            If CharMap(TL).H > WH Then
''''                WH = CharMap(TL).H
''''            End If
''''            WW = WW + CharMap(TL).W
''''            WL = WL + 1
''''
''''        End If
''''
''''
''''
''''    Next TL
''''
''''    WordMap(WC).H = WH
''''    WordMap(WC).W = WW
''''    WordMap(WC).L = WL
''''
''''    WordCount = WC
''''
''''End Sub






















''''Sub Redraw()
''''    Dim i As Long
''''    Dim TextOffsetX As Long
''''    Dim TextOffsetY As Long
''''    Dim TW As Long 'text width
''''    Dim TH As Long 'text height
''''    Dim MarkupList() As Current_Style
''''
''''    Dim TL As Long 'text length
''''    Dim MS As String 'mid string
''''    Dim CS As Long 'command style
''''    Dim SS As Long 'seek string
''''    Dim FC As Long 'fore color
''''    Dim MFC As String 'mid fore color
''''    Dim MC As Long 'markup count
''''    Dim MarkingColor As Long
''''    Dim LNW As Long
''''
''''    If m_bRefreshing Then Exit Sub
''''
''''    m_bRefreshing = True
''''
''''    'UserControl.AutoRedraw = False
''''    UserControl.Cls
''''
''''    UserControl.Font = m_StdFont
''''    UserControl.ForeColor = m_OleForeColor
''''
''''    m_StrText = "" 'reset the text
''''
''''    UserControl.BackColor = m_OleBackgroundColor
''''
''''    UserControl.DrawStyle = 0
''''    If m_bBorder Then
''''        UserControl.Line (0, 0)-(0, UserControl.ScaleHeight), m_OleBorderColor
''''        UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), m_OleBorderColor
''''        UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), m_OleBorderColor
''''        UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), m_OleBorderColor
''''
''''        UserControl.Line (UserControl.ScaleWidth, 0)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), m_OleBorderColor
''''    End If
''''
''''
''''    'TextOffsetX = 0
''''    'TextOffsetY = 0
''''
''''    UserControl.ForeColor = m_OleForeColor
''''
''''
''''    MarkingColor = -1
''''
''''    TH = UserControl.TextHeight(MS)
''''
''''    If m_bLineNumbers Then
''''
''''        LNW = 0
''''
''''        For i = 0 To Fix((UserControl.ScaleHeight - 2) / TH)
''''            UserControl.CurrentX = 3
''''            UserControl.CurrentY = TH * i
''''            UserControl.Print CStr(i + 1)
''''            TW = UserControl.TextWidth(i + 1)
''''            If LNW < TW Then
''''                LNW = TW
''''            End If
''''        Next i
''''
''''        LNW = LNW + TH / 2
''''        UserControl.Line (LNW, 0)-(LNW, UserControl.ScaleHeight), m_OleForeColor
''''        LNW = LNW + TH / 2
''''        TextOffsetX = LNW
''''
''''    End If
''''
''''
''''
''''
''''
''''    For TL = 1 To Len(m_StrMarkupText)
''''
''''
''''        CS = 0
''''        MS = Mid$(m_StrMarkupText, TL, 1)
''''
''''        Select Case MS
''''            Case "\" 'an new line maybe?
''''
''''
''''            Case "{" 'something importand starts
''''
''''                'CS = GetMidText(m_StrMarkupText, "{", "}", TL)
''''                SS = InStr(TL + 1, m_StrMarkupText, " ")
''''                If SS > 0 Then
''''                    'Clipboard.Clear
''''                    'Clipboard.SetText m_StrMarkupText
''''
''''                    'Debug.Print Mid$(m_StrMarkupText, TL + 1, 1)
''''                    If Asc(Mid$(m_StrMarkupText, TL + 1, 1)) = 92 Then
''''                        CS = Asc(Mid$(m_StrMarkupText, TL + 2, 1))
''''                    Else
''''
''''                    End If
''''
''''                    Select Case CS
''''                        Case 0
''''                            ReDim Preserve MarkupList(0 To MC)
''''                            MarkupList(MC).lStyle = sNone
''''                            MC = MC + 1
''''                            TL = TL + 3
''''
''''                        Case 98 '"\b"
''''                            ReDim Preserve MarkupList(0 To MC)
''''                            MarkupList(MC).lStyle = sBold
''''                            MarkupList(MC).prev_Value = UserControl.FontBold
''''                            UserControl.FontBold = Not UserControl.FontBold
''''                            MC = MC + 1
''''                            TL = TL + 3
''''                        Case 117 '"\u"
''''                            ReDim Preserve MarkupList(0 To MC)
''''                            MarkupList(MC).lStyle = sUnderline
''''                            MarkupList(MC).prev_Value = UserControl.FontUnderline
''''                            UserControl.FontUnderline = Not UserControl.FontUnderline
''''                            MC = MC + 1
''''                            TL = TL + 3
''''
''''                        Case 105 '"\i"
''''                            ReDim Preserve MarkupList(0 To MC)
''''                            MarkupList(MC).lStyle = sItalic
''''                            MarkupList(MC).prev_Value = UserControl.FontItalic
''''                            UserControl.FontItalic = Not UserControl.FontItalic
''''
''''                            MC = MC + 1
''''                            TL = TL + 3
''''
''''                        Case 99 '"\c"
''''                            FC = InStr(TL + 3, m_StrMarkupText, " ")
''''                            MFC = Mid(m_StrMarkupText, FC + 1, 6)
''''                            FC = CLng("&h" & MFC)
''''                            ReDim Preserve MarkupList(0 To MC)
''''                            MarkupList(MC).lStyle = sForeColor
''''                            MarkupList(MC).prev_Value = UserControl.ForeColor
''''                            UserControl.ForeColor = FC
''''                            MC = MC + 1
''''                            TL = TL + 3 + 6 + 1
''''
''''                        Case 109 '"\m"
''''                            FC = InStr(TL + 3, m_StrMarkupText, " ")
''''                            MFC = Mid(m_StrMarkupText, FC + 1, 6)
''''                            FC = CLng("&h" & MFC)
''''                            ReDim Preserve MarkupList(0 To MC)
''''                            MarkupList(MC).lStyle = sMarking
''''                            MarkupList(MC).prev_Value = MarkingColor
''''                            MarkingColor = FC
''''                            MC = MC + 1
''''                            TL = TL + 3 + 6 + 1
''''
''''                    End Select
''''                End If
''''
''''                GoTo nextChar
''''
''''            Case "}"
''''                If MC > 0 Then
''''                    MC = MC - 1
''''
''''                    Select Case MarkupList(MC).lStyle
''''                        Case sNone
''''
''''                        Case sBold
''''                            UserControl.FontBold = CBool(MarkupList(MC).prev_Value)
''''
''''                        Case sFontName
''''                            UserControl.FontName = CStr(MarkupList(MC).prev_Value)
''''
''''                        Case sUnderline
''''                            UserControl.FontUnderline = CBool(MarkupList(MC).prev_Value)
''''
''''                        Case sItalic
''''                            UserControl.FontItalic = CBool(MarkupList(MC).prev_Value)
''''
''''                        Case sForeColor
''''                            UserControl.ForeColor = CLng(MarkupList(MC).prev_Value)
''''
''''                        Case sMarking
''''                            MarkingColor = CLng(MarkupList(MC).prev_Value)
''''
''''                    End Select
''''
''''                    ReDim Preserve MarkupList(0 To IIf(MC > 0, MC - 1, 0))
''''                End If
''''
''''                GoTo nextChar
''''        End Select
''''
''''
''''        m_StrText = m_StrText & MS
''''
''''        TW = UserControl.TextWidth(MS)
''''
''''        If TextOffsetX + TW > UserControl.ScaleWidth Then
''''            TextOffsetX = IIf(m_bLineNumbers, LNW, 0)
''''
''''            TextOffsetY = TextOffsetY + TH
''''        End If
''''
''''        UserControl.CurrentX = TextOffsetX
''''        UserControl.CurrentY = TextOffsetY
''''
''''
''''        If MarkingColor > -1 Then
''''            ReDim pts(0 To 3)
''''            pts(0).X = TextOffsetX
''''            pts(0).Y = TextOffsetY
''''
''''            pts(1).X = TextOffsetX + TW
''''            pts(1).Y = TextOffsetY
''''
''''            pts(2).X = TextOffsetX + TW
''''            pts(2).Y = TextOffsetY + TH
''''
''''            pts(3).X = TextOffsetX
''''            pts(3).Y = TextOffsetY + TH
''''
''''            UserControl.FillColor = MarkingColor
''''            UserControl.FillStyle = vbFSSolid
''''            UserControl.DrawStyle = 5
''''            Polygon UserControl.hdc, pts(0), 4
''''
''''        Else
''''            UserControl.DrawWidth = 1
''''            UserControl.FillStyle = vbFSTransparent
''''        End If
''''
''''        If MS = " " Then
''''            If TextOffsetX = IIf(m_bLineNumbers, LNW, 0) Then GoTo nextChar
''''        End If
''''
''''        If TextOffsetY < UserControl.ScaleHeight Then
''''            UserControl.Print MS
''''        End If
''''
''''        TextOffsetX = TextOffsetX + TW
''''
''''nextChar:
''''
''''    Next TL
''''
''''doneRefreshing:
''''
''''
''''    m_bRefreshing = False
''''End Sub
