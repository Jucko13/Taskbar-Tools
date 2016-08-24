Attribute VB_Name = "modInternet"
Option Explicit

Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hInternetSession As Long, ByVal lpszVerb As String, ByVal lpszObjectName As String, ByVal lpszVersion As String, ByVal lpszReferer As String, ByVal lpszAcceptTypes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As String, ByVal lOptionalLength As Long) As Boolean
Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInternetHandle As Long) As Boolean
Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal lpszServerName As String, ByVal nProxyPort As Integer, ByVal lpszUsername As String, ByVal lpszPassword As String, ByVal dwService As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszCallerName As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal dwOption As Long, lpBuffer As Any, ByVal dwBufferLength As Long) As Long

Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Global Const IF_FROM_CACHE = &H1000000
Global Const IF_MAKE_PERSISTENT = &H2000000
Global Const IF_NO_CACHE_WRITE = &H4000000
Global Const INTERNET_DEFAULT_HTTP_PORT = 80
Global Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000
Global Const INTERNET_FLAG_RELOAD = &H80000000
Global Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Global Const INTERNET_OPTION_PER_CONNECTION_OPTION As Long = 75
Global Const INTERNET_OPTION_REFRESH As Long = 37
Global Const INTERNET_OPTION_SETTINGS_CHANGED As Long = 39
Global Const INTERNET_PER_CONN_FLAGS As Long = 1
Global Const INTERNET_PER_CONN_PROXY_BYPASS As Long = 3
Global Const INTERNET_PER_CONN_PROXY_SERVER As Long = 2
Global Const INTERNET_SERVICE_HTTP = 3
Global Const BUFFER_LEN = 256

Global Const UserAgentString As String = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; InfoPath.3)"

Private m_SafeChar(0 To 255) As Boolean

Enum Results
    NONE = 0
    OK = 1
    Fatal = 2
    Login = 4
End Enum

Global lastCheckResult As Integer




Function BaseJucko_Encode(str As String) As String
    Dim jucko As String
    jucko = Chr(65) & Chr(66) & Chr(67) & Chr(72) & Chr(73) & Chr(81) & Chr(82) & Chr(114) & Chr(74) & Chr(75) & Chr(118) & Chr(119) & Chr(120) & Chr(121) & Chr(122) & Chr(48) & Chr(49) & Chr(50) & Chr(51) & Chr(52) & Chr(78) & Chr(79) & Chr(80) & Chr(87) & Chr(88) & Chr(89) & Chr(90) & Chr(97) & Chr(98) & Chr(99) & Chr(100) & Chr(68) & Chr(69) & Chr(70) & Chr(71) & Chr(101) & Chr(102) & Chr(103) & Chr(104) & Chr(105) & Chr(106) & Chr(107) & Chr(83) & Chr(84) & Chr(85) & Chr(86) & Chr(108) & Chr(109) & Chr(110) & Chr(111) & Chr(112) & Chr(113) & Chr(115) & Chr(116) & Chr(117) & Chr(53) & Chr(54) & Chr(76) & Chr(77) & Chr(55) & Chr(56) & Chr(57) & Chr(43) & Chr(47)
    
    Dim juckoSplit(0 To 80) As String * 1
    Dim i As Long
    
    For i = 1 To Len(jucko)
        juckoSplit(i - 1) = Mid$(jucko, i, 1)
    Next i
    
    Dim Parts(0 To 7) As Long
    Dim currentStep As Long
    Dim stringAsc(0 To 4) As Long
    
    Dim totalResult As String
    
    Dim j As Long
    
    For i = 0 To Len(str) - 1 Step 5
        
        For j = 0 To 4
            If Len(str) - 1 >= i + j Then
                stringAsc(j) = Asc(Mid(str, i + j + 1, 1))
            Else
                stringAsc(j) = 0 'Len(jucko) - 1
            End If
        Next j

        Parts(0) = (stringAsc(0) And 248) / 2 ^ 3
        Parts(1) = (stringAsc(0) And 7) * 2 ^ 2
        Parts(1) = Parts(1) Or (stringAsc(1) And 192) / 2 ^ 6
        Parts(2) = (stringAsc(1) And 62) / 2 ^ 1
        Parts(3) = (stringAsc(1) And 1)
        Parts(3) = ((stringAsc(2) And 240) / 2 ^ 4) Or Parts(3) * 2 ^ 4
        Parts(4) = ((stringAsc(2) And 15) * 2 ^ 1)
        Parts(4) = Parts(4) Or (stringAsc(3) And 128)
        Parts(5) = (stringAsc(3) And 124) / 2 ^ 2
        Parts(6) = (stringAsc(3) And 3) * 2 ^ 3
        Parts(6) = ((stringAsc(4) And 224) / 2 ^ 5) Or Parts(6)
        Parts(7) = (stringAsc(4) And 31)
        
        For j = 0 To 7
            totalResult = totalResult & juckoSplit(Parts(j))
        Next j

    Next i
    
    
    BaseJucko_Encode = totalResult
End Function


Function BaseJucko_Decode(str As String) As String
    Dim jucko As String
    jucko = Chr(65) & Chr(66) & Chr(67) & Chr(72) & Chr(73) & Chr(81) & Chr(82) & Chr(114) & Chr(74) & Chr(75) & Chr(118) & Chr(119) & Chr(120) & Chr(121) & Chr(122) & Chr(48) & Chr(49) & Chr(50) & Chr(51) & Chr(52) & Chr(78) & Chr(79) & Chr(80) & Chr(87) & Chr(88) & Chr(89) & Chr(90) & Chr(97) & Chr(98) & Chr(99) & Chr(100) & Chr(68) & Chr(69) & Chr(70) & Chr(71) & Chr(101) & Chr(102) & Chr(103) & Chr(104) & Chr(105) & Chr(106) & Chr(107) & Chr(83) & Chr(84) & Chr(85) & Chr(86) & Chr(108) & Chr(109) & Chr(110) & Chr(111) & Chr(112) & Chr(113) & Chr(115) & Chr(116) & Chr(117) & Chr(53) & Chr(54) & Chr(76) & Chr(77) & Chr(55) & Chr(56) & Chr(57) & Chr(43) & Chr(47)
    
    Dim juckoSplit(0 To 63) As String * 1
    Dim i As Long
    
    Dim privatessss As String
    
    For i = 1 To Len(jucko)
        juckoSplit(i - 1) = Mid$(jucko, i, 1)
        privatessss = privatessss & "chr(" & Asc(Mid$(jucko, i, 1)) & ") & "
    Next i
    
    Dim Parts(0 To 7) As Long
    Dim currentStep As Long
    Dim stringAsc(0 To 4) As Long
    Dim stringMid As String
    
    Dim totalResult As String
    
    Dim j As Long
    Dim K As Long
    
    For i = 0 To Len(str) - 1 Step 8
        
        For K = 0 To 7
            stringMid = Mid(str, i + K + 1, 1)
            
            For j = 0 To UBound(juckoSplit)
                If stringMid = juckoSplit(j) Then
                    Parts(K) = j
                    Exit For
                End If
            Next j
        Next K
        
        stringAsc(0) = Parts(0) * 2 ^ 3 Or (Parts(1) And 28) / 2 ^ 2
        
        stringAsc(1) = (Parts(1) And 3) * 2 ^ 6 Or (Parts(2) * 2 ^ 1) Or (Parts(3) And 16) / 2 ^ 4
        
        stringAsc(2) = (Parts(3) And 15) * 2 ^ 4 Or (Parts(4) And 30) / 2 ^ 1
        
        stringAsc(3) = (Parts(4) And 1) * 2 ^ 7 Or (Parts(5) * 2 ^ 2) Or (Parts(6) And 24) / 2 ^ 3
        
        stringAsc(4) = (Parts(6) And 7) * 2 ^ 5 Or Parts(7)
        
        For j = 0 To 4
            If stringAsc(j) <> 0 Then
                totalResult = totalResult & Chr(stringAsc(j))
            End If
        Next j

    Next i
    
    BaseJucko_Decode = totalResult
End Function


Function Encrypt(str As String, Optional Password As String = "None") As String
    Dim tmpPart As String
    Dim i As Long
    Dim countitup As Long

    For i = 1 To Len(str)
        countitup = Asc(Mid$(str, i, 1))
        countitup = countitup + Asc(Mid$(Password, (i Mod Len(Password)) + 1, 1))
        tmpPart = tmpPart & Chr$(countitup And &HFF)
    Next i
    
    Encrypt = tmpPart
End Function



Function Decrypt(str As String, Optional Password As String = "None") As String
    Dim tmpPart As String
    Dim i As Long
    Dim countitup As Long

    For i = 1 To Len(str)
        countitup = Asc(Mid$(str, i, 1))
        countitup = countitup - Asc(Mid$(Password, (i Mod Len(Password)) + 1, 1))
        tmpPart = tmpPart & Chr$(countitup And &HFF)
    Next i
    
    Decrypt = tmpPart

End Function


Public Function GetUrlSource(sURL As String) As String
    Dim sBuffer   As String * BUFFER_LEN, iResult As Integer, sData As String
    Dim hInternet As Long, hSession As Long, lReturn As Long
    Dim tmpTime   As Long

    Const INTERNET_FLAG_ASYNC = &H10000000
    
    ''st sURL
    
    hSession = InternetOpen(UserAgentString, 0, vbNullString, vbNullString, 0)

    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

    If hInternet Then
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sData = Mid(sBuffer, 1, lReturn)

        Do While lReturn <> 0
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sData = sData + Mid(sBuffer, 1, lReturn)
        Loop
    End If
    iResult = InternetCloseHandle(hInternet)

    GetUrlSource = sData
End Function

Sub st(str As String)
    On Error GoTo ErrHandler:
    Clipboard.Clear
    Clipboard.SetText str
    Exit Sub
    
ErrHandler:
    Debug.Print Err.Description
    
End Sub


Public Function PostInfo(webPage As String, PostData As String) As String

    Dim hInternetOpen As Long
    Dim hInternetConnect As Long
    Dim hHttpOpenRequest As Long
    Dim bRet As Boolean
    Dim bDoLoop As Boolean
    Dim sReadBuffer As String * 4096
    Dim lNumberOfBytesRead As Long
    Dim sBuffer As String
    Dim sHeader As String
    Dim lPostDataLen As Long
    Dim Script As String
    Dim Server As String
    Dim tmpTime As Long



    Server = Replace(webPage, "http://", "")
    Script = Mid(Server, InStr(Server, "/"), Len(Server))
    Server = Mid(Server, 1, InStr(Server, "/") - 1)

    hInternetOpen = 0
    hInternetConnect = 0
    hHttpOpenRequest = 0


    hInternetOpen = InternetOpen(UserAgentString, 0, vbNullString, vbNullString, 0)

    If hInternetOpen <> 0 Then

        hInternetConnect = InternetConnect(hInternetOpen, Server, INTERNET_DEFAULT_HTTP_PORT, vbNullString, "HTTP/1.0", INTERNET_SERVICE_HTTP, 0, 0)

        If hInternetConnect <> 0 Then
            hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "POST", Script, "HTTP/1.1", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

            If hHttpOpenRequest <> 0 Then
                sHeader = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
                lPostDataLen = Len(PostData)
                bRet = HttpSendRequest(hHttpOpenRequest, sHeader, Len(sHeader), PostData, lPostDataLen)

                bDoLoop = True
                Do While bDoLoop
                    sReadBuffer = vbNullString
                    bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
                    sBuffer = sBuffer & Left(sReadBuffer, lNumberOfBytesRead)
                    If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
                    DoEvents
                Loop

                PostInfo = sBuffer
                bRet = InternetCloseHandle(hHttpOpenRequest)
            End If
            bRet = InternetCloseHandle(hInternetConnect)
        End If
        bRet = InternetCloseHandle(hInternetOpen)
    End If
End Function




Public Function UploadFile(webPage As String, imageFolder As String, tempImage As String, imagePath As String) As String

    Dim hInternetOpen As Long
    Dim hInternetConnect As Long
    Dim hHttpOpenRequest As Long
    Dim bRet As Boolean
    Dim bDoLoop As Boolean
    Dim sReadBuffer As String * 4096
    Dim lNumberOfBytesRead As Long
    Dim sBuffer As String
    Dim sHeader As String
    Dim lPostDataLen As Long
    Dim Script As String
    Dim Server As String
    Dim tmpTime As Long

    Dim PostData As String
    
    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open tempImage For Binary As fileNum
    lPostDataLen = LOF(fileNum)
    PostData = String(lPostDataLen, " ")
    
    Get fileNum, , PostData
    Close fileNum
    

    Server = Replace(webPage, "http://", "")
    Script = Mid(Server, InStr(Server, "/"), Len(Server))
    Server = Mid(Server, 1, InStr(Server, "/") - 1)
    
    
    Dim filename As String
    
    filename = Right(imagePath, Len(imagePath) - InStrRev(imagePath, "\"))
    
    hInternetOpen = 0
    hInternetConnect = 0
    hHttpOpenRequest = 0


    hInternetOpen = InternetOpen(UserAgentString, 0, vbNullString, vbNullString, 0)

    If hInternetOpen <> 0 Then

        hInternetConnect = InternetConnect(hInternetOpen, Server, INTERNET_DEFAULT_HTTP_PORT, vbNullString, "HTTP/1.0", INTERNET_SERVICE_HTTP, 0, 0)

        If hInternetConnect <> 0 Then
            hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "POST", Script, "HTTP/1.1", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

            If hHttpOpenRequest <> 0 Then
                sHeader = "Content-Type: application/octet-stream" & vbCrLf
                sHeader = sHeader & "Content-Disposition: attachment; filename=""" & filename & """" & vbCrLf
                sHeader = sHeader & "FileName: " & filename & "" & vbCrLf
                sHeader = sHeader & "FolderName: " & imageFolder & "" & vbCrLf
                'MsgBox sHeader
                
                'lPostDataLen = Len(PostData)
                bRet = HttpSendRequest(hHttpOpenRequest, sHeader, Len(sHeader), PostData, lPostDataLen)

                bDoLoop = True
                Do While bDoLoop
                    sReadBuffer = vbNullString
                    bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
                    sBuffer = sBuffer & Left(sReadBuffer, lNumberOfBytesRead)
                    If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
                    DoEvents
                Loop

                'PostInfo = sBuffer
                bRet = InternetCloseHandle(hHttpOpenRequest)
            End If
            bRet = InternetCloseHandle(hInternetConnect)
        End If
        bRet = InternetCloseHandle(hInternetOpen)
    End If
    
    UploadFile = sBuffer
End Function





Function URLEncode(ByVal txt As String) As String
Dim i As Integer
Dim Ch As String
Dim ch_asc As Integer
Dim result As String

    SetSafeChars

    result = ""
    For i = 1 To Len(txt)
        ' Translate the next character.
        Ch = Mid$(txt, i, 1)
        ch_asc = Asc(Ch)
        
        'Debug.Assert ch <> "ä"
        
        
        If ch_asc = vbKeySpace Then
            ' Use a plus.
            result = result & "+"
        ElseIf m_SafeChar(ch_asc) Then
            ' Use the character.
            result = result & Ch
        Else
            ' Convert the character to hex.
            result = result & "%" & Right$("0" & _
                Hex$(ch_asc), 2)
                
            'result = "%" & Right$("0" & Hex(ch_asc), 2)
        End If
    Next i

    URLEncode = result
End Function


' Set m_SafeChar(i) = True for characters that
' do not need protection.
Private Sub SetSafeChars()
    Static done_before As Boolean
    Dim i As Integer

    If done_before Then Exit Sub
    done_before = True

    For i = 0 To 47
        m_SafeChar(i) = False
    Next i
    For i = 48 To 57
        m_SafeChar(i) = True
    Next i
    For i = 58 To 64
        m_SafeChar(i) = False
    Next i
    For i = 65 To 90
        m_SafeChar(i) = True
    Next i
    For i = 91 To 96
        m_SafeChar(i) = False
    Next i
    For i = 97 To 122
        m_SafeChar(i) = True
    Next i
    For i = 123 To 255
        m_SafeChar(i) = False
    Next i
End Sub




Public Function HTMLEntititesDecode(p_strText As String) As String
Dim strTemp As String
strTemp = p_strText
strTemp = Replace(strTemp, "&quot;", """")
strTemp = Replace(strTemp, "&amp;", "&")
strTemp = Replace(strTemp, "&apos;", "'")
strTemp = Replace(strTemp, "&lt;", "<")
strTemp = Replace(strTemp, "&gt;", ">")
strTemp = Replace(strTemp, "&nbsp;", " ")
strTemp = Replace(strTemp, "&iexcl;", "¡")
strTemp = Replace(strTemp, "&cent;", "¢")
strTemp = Replace(strTemp, "&pound;", "£")
strTemp = Replace(strTemp, "&curren;", "¤")
strTemp = Replace(strTemp, "&yen;", "¥")
strTemp = Replace(strTemp, "&brvbar;", "¦")
strTemp = Replace(strTemp, "&sect;", "§")
strTemp = Replace(strTemp, "&uml;", "¨")
strTemp = Replace(strTemp, "&copy;", "©")
strTemp = Replace(strTemp, "&ordf;", "ª")
strTemp = Replace(strTemp, "&laquo;", "«")
strTemp = Replace(strTemp, "&not;", "¬")
strTemp = Replace(strTemp, "*", "")
strTemp = Replace(strTemp, "&reg;", "®")
strTemp = Replace(strTemp, "&macr;", "¯")
strTemp = Replace(strTemp, "&deg;", "°")
strTemp = Replace(strTemp, "&plusmn;", "±")
strTemp = Replace(strTemp, "&sup2;", "²")
strTemp = Replace(strTemp, "&sup3;", "³")
strTemp = Replace(strTemp, "&acute;", "´")
strTemp = Replace(strTemp, "&micro;", "µ")
strTemp = Replace(strTemp, "&para;", "¶")
strTemp = Replace(strTemp, "&middot;", "·")
strTemp = Replace(strTemp, "&cedil;", "¸")
strTemp = Replace(strTemp, "&sup1;", "¹")
strTemp = Replace(strTemp, "&ordm;", "º")
strTemp = Replace(strTemp, "&raquo;", "»")
strTemp = Replace(strTemp, "&frac14;", "¼")
strTemp = Replace(strTemp, "&frac12;", "½")
strTemp = Replace(strTemp, "&frac34;", "¾")
strTemp = Replace(strTemp, "&iquest;", "¿")
strTemp = Replace(strTemp, "&Agrave;", "À")
strTemp = Replace(strTemp, "&Aacute;", "Á")
strTemp = Replace(strTemp, "&Acirc;", "Â")
strTemp = Replace(strTemp, "&Atilde;", "Ã")
strTemp = Replace(strTemp, "&Auml;", "Ä")
strTemp = Replace(strTemp, "&Aring;", "Å")
strTemp = Replace(strTemp, "&AElig;", "Æ")
strTemp = Replace(strTemp, "&Ccedil;", "Ç")
strTemp = Replace(strTemp, "&Egrave;", "È")
strTemp = Replace(strTemp, "&Eacute;", "É")
strTemp = Replace(strTemp, "&Ecirc;", "Ê")
strTemp = Replace(strTemp, "&Euml;", "Ë")
strTemp = Replace(strTemp, "&Igrave;", "Ì")
strTemp = Replace(strTemp, "&Iacute;", "Í")
strTemp = Replace(strTemp, "&Icirc;", "Î")
strTemp = Replace(strTemp, "&Iuml;", "Ï")
strTemp = Replace(strTemp, "&ETH;", "Ð")
strTemp = Replace(strTemp, "&Ntilde;", "Ñ")
strTemp = Replace(strTemp, "&Ograve;", "Ò")
strTemp = Replace(strTemp, "&Oacute;", "Ó")
strTemp = Replace(strTemp, "&Ocirc;", "Ô")
strTemp = Replace(strTemp, "&Otilde;", "Õ")
strTemp = Replace(strTemp, "&Ouml;", "Ö")
strTemp = Replace(strTemp, "&times;", "×")
strTemp = Replace(strTemp, "&Oslash;", "Ø")
strTemp = Replace(strTemp, "&Ugrave;", "Ù")
strTemp = Replace(strTemp, "&Uacute;", "Ú")
strTemp = Replace(strTemp, "&Ucirc;", "Û")
strTemp = Replace(strTemp, "&Uuml;", "Ü")
strTemp = Replace(strTemp, "&Yacute;", "Ý")
strTemp = Replace(strTemp, "&THORN;", "Þ")
strTemp = Replace(strTemp, "&szlig;", "ß")
strTemp = Replace(strTemp, "&agrave;", "à")
strTemp = Replace(strTemp, "&aacute;", "á")
strTemp = Replace(strTemp, "&acirc;", "â")
strTemp = Replace(strTemp, "&atilde;", "ã")
strTemp = Replace(strTemp, "&auml;", "ä")
strTemp = Replace(strTemp, "&aring;", "å")
strTemp = Replace(strTemp, "&aelig;", "æ")
strTemp = Replace(strTemp, "&ccedil;", "ç")
strTemp = Replace(strTemp, "&egrave;", "è")
strTemp = Replace(strTemp, "&eacute;", "é")
strTemp = Replace(strTemp, "&ecirc;", "ê")
strTemp = Replace(strTemp, "&euml;", "ë")
strTemp = Replace(strTemp, "&igrave;", "ì")
strTemp = Replace(strTemp, "&iacute;", "í")
strTemp = Replace(strTemp, "&icirc;", "î")
strTemp = Replace(strTemp, "&iuml;", "ï")
strTemp = Replace(strTemp, "&eth;", "ð")
strTemp = Replace(strTemp, "&ntilde;", "ñ")
strTemp = Replace(strTemp, "&ograve;", "ò")
strTemp = Replace(strTemp, "&oacute;", "ó")
strTemp = Replace(strTemp, "&ocirc;", "ô")
strTemp = Replace(strTemp, "&otilde;", "õ")
strTemp = Replace(strTemp, "&ouml;", "ö")
strTemp = Replace(strTemp, "&divide;", "÷")
strTemp = Replace(strTemp, "&oslash;", "ø")
strTemp = Replace(strTemp, "&ugrave;", "ù")
strTemp = Replace(strTemp, "&uacute;", "ú")
strTemp = Replace(strTemp, "&ucirc;", "û")
strTemp = Replace(strTemp, "&uuml;", "ü")
strTemp = Replace(strTemp, "&yacute;", "ý")
strTemp = Replace(strTemp, "&thorn;", "þ")
strTemp = Replace(strTemp, "&yuml;", "ÿ")
strTemp = Replace(strTemp, "&OElig;", "Œ")
strTemp = Replace(strTemp, "&oelig;", "œ")
strTemp = Replace(strTemp, "&Scaron;", "Š")
strTemp = Replace(strTemp, "&scaron;", "š")
strTemp = Replace(strTemp, "&Yuml;", "Ÿ")
strTemp = Replace(strTemp, "&fnof;", "ƒ")
strTemp = Replace(strTemp, "&circ;", "ˆ")
strTemp = Replace(strTemp, "&tilde;", "˜")
strTemp = Replace(strTemp, "&thinsp;", "")
strTemp = Replace(strTemp, "&zwnj;", "")
strTemp = Replace(strTemp, "&zwj;", "")
strTemp = Replace(strTemp, "&lrm;", "")
strTemp = Replace(strTemp, "&rlm;", "")
strTemp = Replace(strTemp, "&ndash;", "–")
strTemp = Replace(strTemp, "&mdash;", "—")
strTemp = Replace(strTemp, "&lsquo;", "‘")
strTemp = Replace(strTemp, "&rsquo;", "’")
strTemp = Replace(strTemp, "&sbquo;", "‚")
strTemp = Replace(strTemp, "&ldquo;", "“")
strTemp = Replace(strTemp, "&rdquo;", "”")
strTemp = Replace(strTemp, "&bdquo;", "„")
strTemp = Replace(strTemp, "&dagger;", "†")
strTemp = Replace(strTemp, "&Dagger;", "‡")
strTemp = Replace(strTemp, "&bull;", "•")
strTemp = Replace(strTemp, "&hellip;", "…")
strTemp = Replace(strTemp, "&permil;", "‰")
strTemp = Replace(strTemp, "&lsaquo;", "‹")
strTemp = Replace(strTemp, "&rsaquo;", "›")
strTemp = Replace(strTemp, "&euro;", "€")
strTemp = Replace(strTemp, "&trade;", "™")
HTMLEntititesDecode = strTemp
End Function
