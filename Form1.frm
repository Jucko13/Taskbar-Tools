VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00584D43&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10545
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   721
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   703
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.uCheckBox uCheckBox3 
      Height          =   900
      Left            =   1240
      TabIndex        =   9
      Top             =   1020
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   1588
      BackgroundColor =   16744576
      BorderColor     =   16761024
      BorderThickness =   2
      Caption         =   "Een Checkbox"
      CaptionBorder   =   -1  'True
      CaptionBorderColor=   16761024
      CaptionOffsetLeft=   10
      CaptionOffsetTop=   1
      CheckBackgroundColor=   16744576
      CheckBorderColor=   16761024
      CheckBorderThickness=   2
      CheckSelectionColor=   12582912
      CheckSize       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12582912
      Value           =   1
   End
   Begin Project1.uLoadBar uLoadCPU 
      Height          =   210
      Left            =   4845
      TabIndex        =   12
      Top             =   15
      Width           =   585
      _ExtentX        =   1058
      _ExtentY        =   370
      BarType         =   0
      BarWidth        =   0
      Caption         =   ""
      CaptionBorder   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   14780160
      Value           =   100
   End
   Begin VB.Timer tmrCheckHotKey 
      Interval        =   1
      Left            =   300
      Top             =   1010
   End
   Begin VB.Timer tmrComports 
      Interval        =   1000
      Left            =   495
      Top             =   3090
   End
   Begin MSCommLib.MSComm comm 
      Left            =   645
      Top             =   4875
      _ExtentX        =   953
      _ExtentY        =   953
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin Project1.uLoadBar uLoadBar1 
      Height          =   360
      Left            =   1730
      TabIndex        =   7
      Top             =   4640
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   635
      BackgroundColor =   0
      BarColor        =   65280
      BarType         =   0
      BarWidth        =   0
      Border          =   0   'False
      BorderColor     =   0
      Caption         =   "Loading..."
      CaptionBorder   =   -1  'True
      CaptionBorderColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65280
   End
   Begin Project1.uDropDown uCommPort 
      Height          =   240
      Left            =   4520
      TabIndex        =   6
      Top             =   240
      Width           =   1600
      _ExtentX        =   2831
      _ExtentY        =   423
      BackgroundColor =   14780160
      ForeColor       =   8388608
      SelectionBackgroundColor=   14780160
      SelectionBorderColor=   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      ItemHeight      =   40
      VisibleItems    =   15
      ScrollBarWidth  =   35
   End
   Begin Project1.uCheckBox uCheckBox1 
      Height          =   880
      Left            =   1220
      TabIndex        =   5
      Top             =   3900
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   1535
      BackgroundColor =   8421631
      BorderColor     =   12632319
      BorderThickness =   2
      Caption         =   "Een Checkbox"
      CaptionBorder   =   -1  'True
      CaptionBorderColor=   12632319
      CaptionOffsetLeft=   10
      CaptionOffsetTop=   1
      CheckBackgroundColor=   8421631
      CheckBorderColor=   12632319
      CheckBorderThickness=   2
      CheckSelectionColor=   128
      CheckSize       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   128
      Value           =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   375
      Top             =   1875
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   20
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Top             =   20
      Width           =   4320
   End
   Begin Project1.uButton cmdSettings 
      Height          =   240
      Left            =   1140
      TabIndex        =   1
      Top             =   240
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   423
      MouseOverBackgroundColor=   8421504
      CaptionBorderColor=   16711680
      Caption         =   "Google"
      BorderAnimation =   4
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdGames 
      Height          =   240
      Left            =   2265
      TabIndex        =   2
      Top             =   240
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   423
      MouseOverBackgroundColor=   8421504
      Caption         =   "Games"
      BorderAnimation =   4
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdPrograms 
      Height          =   240
      Left            =   3390
      TabIndex        =   3
      Top             =   240
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   423
      MouseOverBackgroundColor=   8421504
      Caption         =   "Programs"
      BorderAnimation =   4
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uButton cmdExtra 
      Height          =   240
      Left            =   15
      TabIndex        =   4
      Top             =   240
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   423
      MouseOverBackgroundColor=   8421504
      Caption         =   "X"
      BorderAnimation =   4
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uCheckBox uCheckBox2 
      Height          =   900
      Left            =   1215
      TabIndex        =   8
      Top             =   1950
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   1588
      BackgroundColor =   8454016
      BorderColor     =   12648384
      BorderThickness =   2
      Caption         =   "Een Checkbox"
      CaptionBorder   =   -1  'True
      CaptionBorderColor=   12648384
      CaptionOffsetLeft=   10
      CaptionOffsetTop=   1
      CheckBackgroundColor=   8454016
      CheckBorderColor=   12648384
      CheckBorderThickness=   2
      CheckSelectionColor=   49152
      CheckSize       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   49152
      Value           =   1
   End
   Begin Project1.uCheckBox uCheckBox4 
      Height          =   900
      Left            =   1215
      TabIndex        =   10
      Top             =   2925
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   1588
      BackgroundColor =   8454143
      BorderColor     =   12648447
      BorderThickness =   2
      Caption         =   "Een Checkbox"
      CaptionBorder   =   -1  'True
      CaptionBorderColor=   12648447
      CaptionOffsetLeft=   10
      CaptionOffsetTop=   1
      CheckBackgroundColor=   8454143
      CheckBorderColor=   12648447
      CheckBorderThickness=   2
      CheckSelectionColor=   49344
      CheckSize       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   49344
      Value           =   1
   End
   Begin Project1.uCheckBox uCheckBox5 
      Height          =   360
      Left            =   945
      TabIndex        =   11
      Top             =   5775
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   635
      BackgroundColor =   5786947
      Caption         =   "Een Checkbox"
      CaptionBorder   =   -1  'True
      CaptionBorderColor=   0
      CaptionOffsetLeft=   10
      CaptionOffsetTop=   1
      CheckBackgroundColor=   5786947
      CheckBorderColor=   16777215
      CheckSelectionColor=   16777215
      CheckSize       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      Value           =   1
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   420
      Left            =   1170
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6450
      Width           =   675
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetProcessDPIAware Lib "user32.dll" () As Long


Private DragX As Long
Private Dragging As Boolean
Private ProgramTop As Long
Private ProgramLeft As Long
Private ProgramLeft2 As Long
Private ProgramParentLeft As Long

Private ShowMode As Boolean


Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Boolean
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function GetKeyboardState Lib "user32.dll" (pbKeyState As Byte) As Long

Private Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private KeyFired(0 To 255) As Byte
Private KeyStates(0 To 255) As Byte
Private Const KEY_CONTROL_F12 As Long = 1
Private Const KEY_CONTROL_F9 As Long = 2

Private CpuUsage As clsCPUUsageNT

'Private Sub cmdClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Unload frmGames
'    Unload frmSettings
'
'    Unload Me
'End Sub



Public Sub Wait(ByVal dblMilliseconds As Double)
    Dim dblStart As Double
    Dim dblEnd As Double
    Dim dblTickCount As Double

    dblTickCount = GetTickCount()
    dblStart = GetTickCount()
    dblEnd = GetTickCount + dblMilliseconds

    Do
        DoEvents
        dblTickCount = GetTickCount()
    Loop Until dblTickCount > dblEnd Or dblTickCount < dblStart
End Sub


Private Sub cmdPrograms_Click(Button As Integer, x As Single, y As Single)
    CloseWindows

    If allSettings(ActiveSetting).sButtonText(2) = "" Then
        frmProgramTools.Visible = False
        frmProgramTools.Form_Resize
        frmProgramTools.Left = Me.Left

        If Me.Top - frmProgramTools.Height + Screen.TwipsPerPixelY < 0 Then
            frmProgramTools.Top = Me.Top + Me.Height - Screen.TwipsPerPixelY
        Else
            frmProgramTools.Top = Me.Top - frmProgramTools.Height + Screen.TwipsPerPixelY
        End If
        frmProgramTools.Width = 7000 'Me.Width

        frmProgramTools.Visible = True
        frmProgramTools.tmrTopMost_Timer
    Else
        ExecuteButton 2
    End If
End Sub

Private Sub cmdGames_Click(Button As Integer, x As Single, y As Single)
    CloseWindows
    If allSettings(ActiveSetting).sButtonText(1) = "" Then
        frmGames.Visible = False
        frmGames.Left = Me.Left

        If Me.Top - frmGames.Height + Screen.TwipsPerPixelY < 0 Then
            frmGames.Top = Me.Top + Me.Height - Screen.TwipsPerPixelY
        Else
            frmGames.Top = Me.Top - frmGames.Height + Screen.TwipsPerPixelY
        End If
        frmGames.Width = 6000 'Me.Width

        frmGames.Visible = True
        frmGames.tmrTopMost_Timer
    Else
        ExecuteButton 1
    End If

End Sub

Private Sub cmdSettings_Click(Button As Integer, x As Single, y As Single)
    CloseWindows

    If allSettings(ActiveSetting).sButtonText(0) = "" Then
        frmSettings.Form_Load
        frmSettings.Visible = False
        frmSettings.Left = Me.Left
        If Me.Top - frmSettings.Height + Screen.TwipsPerPixelY < 0 Then
            frmSettings.Top = Me.Top + Me.Height    '+ Screen.TwipsPerPixelY
        Else
            frmSettings.Top = Me.Top - frmSettings.Height + Screen.TwipsPerPixelY
        End If
        
        frmSettings.Width = 6000 'Me.Width

        frmSettings.Visible = True
        SetTopMostWindow frmSettings.hWnd, True
    Else
        ExecuteButton 0


    End If
End Sub

Sub ExecuteButton(lIndex As Long)
    Dim lPath As String
    Dim lFolder As String
    Dim lParameters As String

    If InStr(1, allSettings(ActiveSetting).sActionPath(lIndex), "{UPLOAD}") Or _
       InStr(1, allSettings(ActiveSetting).sActionParameters(lIndex), "{UPLOAD}") Or _
       InStr(1, allSettings(ActiveSetting).sActionFolder(lIndex), "{UPLOAD}") Then

        On Error GoTo ErrHandler

        ' -p m2560 -D -c stk500v2 -P com9 -b 115200 -F -V -U flash:w:main.hex

        comm.CommPort = uCommPort.ListIndex + 1
        comm.Settings = "19200,n,8,1"
        comm.PortOpen = True

        Wait 10
        If comm.PortOpen = True Then
            comm.DTREnable = False
            Wait 20
            comm.DTREnable = True
            Wait 20
            comm.PortOpen = False

            lPath = Replace(allSettings(ActiveSetting).sActionPath(lIndex), "{UPLOAD}", otherSettings.sUploadArguments)
            lFolder = Replace(allSettings(ActiveSetting).sActionFolder(lIndex), "{UPLOAD}", otherSettings.sUploadArguments)
            lParameters = Replace(allSettings(ActiveSetting).sActionParameters(lIndex), "{UPLOAD}", otherSettings.sUploadArguments)

            lPath = Replace(lPath, "{COM}", CStr(uCommPort.ListIndex + 1))
            lFolder = Replace(lFolder, "{COM}", CStr(uCommPort.ListIndex + 1))
            lParameters = Replace(lParameters, "{COM}", CStr(uCommPort.ListIndex + 1))

            ShellExecute 0, "open", lPath, lParameters, lFolder, vbNormalFocus

        End If
    Else
        lPath = Replace(allSettings(ActiveSetting).sActionPath(lIndex), "{TEXT}", txtSearch.Text)
        lFolder = Replace(allSettings(ActiveSetting).sActionFolder(lIndex), "{TEXT}", txtSearch.Text)
        lParameters = Replace(allSettings(ActiveSetting).sActionParameters(lIndex), "{TEXT}", txtSearch.Text)

        ShellExecute Me.hWnd, "OPEN", lPath, lParameters, lFolder, vbNormalFocus
    End If


    Exit Sub
ErrHandler:
    MsgBox "Could not connect to Com" & uCommPort.ListIndex + 1
End Sub


Sub CloseWindows()
    frmGames.Visible = False
    frmSettings.Visible = False
    frmProgramTools.Visible = False
End Sub



Sub LoadDefaultSettings()

    settingCount = 7
    ReDim allSettings(0 To settingCount - 1)

    allSettings(0).sButtonColor = 31    '&H584D43
    allSettings(0).sButtonFontColor = 35    'vbWhite
    allSettings(0).sButtonText(0) = ""
    allSettings(0).sActionParameters(0) = ""
    allSettings(0).sActionPath(0) = "{TEXT}"
    allSettings(0).sActivationText = ""


    allSettings(1).sButtonColor = 50    '&HE18700
    allSettings(1).sButtonFontColor = 35    'vbWhite
    allSettings(1).sButtonText(0) = "Google"
    allSettings(1).sActionParameters(0) = "https://www.google.nl/search?q=" & "{TEXT}"
    allSettings(1).sActionPath(0) = "chrome.exe"
    allSettings(1).sActivationText = "g"


    allSettings(2).sButtonColor = 8    '&H1EAC0B
    allSettings(2).sButtonFontColor = 35    'vbWhite
    allSettings(2).sButtonText(0) = "Vb6"
    allSettings(2).sActionParameters(0) = "https://www.google.nl/search?q=vb6+" & "{TEXT}" & "+-VB.Net+-.Net"
    allSettings(2).sActionPath(0) = "chrome.exe"
    allSettings(2).sActivationText = "v"


    allSettings(3).sButtonColor = 19    '&H1B9DC8
    allSettings(3).sButtonFontColor = 35    'vbWhite
    allSettings(3).sButtonText(0) = "VindictusDB"
    allSettings(3).sActionParameters(0) = "http://vindictusdb.com/search?sw=" & "{TEXT}"
    allSettings(3).sActionPath(0) = "chrome.exe"
    allSettings(3).sActivationText = "d"


    allSettings(4).sButtonColor = 5    '&H2121EA
    allSettings(4).sButtonFontColor = 35    'vbWhite
    allSettings(4).sButtonText(0) = "YouTube"
    allSettings(4).sActionParameters(0) = "https://www.youtube.com/results?search_query=" & "{TEXT}"
    allSettings(4).sActionPath(0) = "chrome.exe"
    allSettings(4).sActivationText = "y"


    allSettings(5).sButtonColor = 15    '&HC67200
    allSettings(5).sButtonFontColor = 35    'vbWhite
    allSettings(5).sButtonText(0) = "Hotmail"
    allSettings(5).sActionParameters(0) = """https://dub109.mail.live.com/default.aspx#fid=flsearch&srch=1&skws=" & "{TEXT}" & "&sdr=4&satt=0"""
    allSettings(5).sActionPath(0) = "chrome.exe"
    allSettings(5).sActivationText = "h"


    allSettings(6).sButtonColor = 27    '&H9DA934
    allSettings(6).sButtonFontColor = 35    'vbWhite
    allSettings(6).sButtonText(0) = "Upload"
    allSettings(6).sActionParameters(0) = "/c {UPLOAD}"
    allSettings(6).sActionPath(0) = "cmd.exe"
    allSettings(6).sActionFolder(0) = "D:\Github\swagbot\software\arduino\src\"

    allSettings(6).sButtonText(1) = "Make AT2560"
    allSettings(6).sActionParameters(1) = "/c make.exe"
    allSettings(6).sActionPath(1) = "cmd.exe"
    allSettings(6).sActionFolder(1) = "D:\Github\swagbot\software\arduino\src\"

    allSettings(6).sButtonText(2) = "Make AT32"
    allSettings(6).sActionParameters(2) = "/c make.exe"
    allSettings(6).sActionPath(2) = "cmd.exe"
    allSettings(6).sActionFolder(2) = "D:\Github\swagbot\software\rp6\"
    allSettings(6).sActivationText = "a"



    otherSettings.sUploadArguments = "avrdude.exe -p m2560 -D -c stk500v2 -P com{COM} -b 115200 -F -V -U flash:w:main.hex"

    ActiveSetting = 0

    SaveAllSettings
End Sub

Sub DeleteAllSettings()
    Dim i As Long
    Dim j As Long
    Dim currentSaveCount As Long

    'normal settings
    currentSaveCount = CLng(GetSetting(AppName, "Settings", "howmany", 0))
    On Error Resume Next

    For i = 1 To currentSaveCount
        DeleteSetting AppName, "Settings", "sButtonColor(" & i & ")"
        DeleteSetting AppName, "Settings", "sButtonFontColor(" & i & ")"
        For j = 0 To 2
            DeleteSetting AppName, "Settings", "sButtonText(" & i & ")(" & j & ")"
            DeleteSetting AppName, "Settings", "sActionPath(" & i & ")(" & j & ")"
            DeleteSetting AppName, "Settings", "sActionFolder(" & i & ")(" & j & ")"
            DeleteSetting AppName, "Settings", "sActionParameters(" & i & ")(" & j & ")"
        Next j
        DeleteSetting AppName, "Settings", "sActivationText(" & i & ")"
    Next i

    DeleteSetting AppName, "Settings", "howmany"

    'other settings
    DeleteSetting AppName, "Others", "sUploadArguments"


End Sub

Sub SaveAllSettings()
    Dim i As Long
    Dim j As Long

    'normal settings
    SaveSetting AppName, "Settings", "howmany", settingCount

    For i = 1 To settingCount
        SaveSetting AppName, "Settings", "sButtonColor(" & i & ")", allSettings(i - 1).sButtonColor
        SaveSetting AppName, "Settings", "sButtonFontColor(" & i & ")", allSettings(i - 1).sButtonFontColor

        For j = 0 To 2
            SaveSetting AppName, "Settings", "sButtonText(" & i & ")(" & j & ")", allSettings(i - 1).sButtonText(j)
            SaveSetting AppName, "Settings", "sActionPath(" & i & ")(" & j & ")", allSettings(i - 1).sActionPath(j)
            SaveSetting AppName, "Settings", "sActionFolder(" & i & ")(" & j & ")", allSettings(i - 1).sActionFolder(j)
            SaveSetting AppName, "Settings", "sActionParameters(" & i & ")(" & j & ")", allSettings(i - 1).sActionParameters(j)
        Next j

        SaveSetting AppName, "Settings", "sActivationText(" & i & ")", allSettings(i - 1).sActivationText
    Next i

    'other settings
    SaveSetting AppName, "Others", "sUploadArguments", otherSettings.sUploadArguments

End Sub

Sub LoadAllSettings()
    Dim i As Long
    Dim j As Long

    'normal settings
    settingCount = CLng(GetSetting(AppName, "Settings", "howmany", 0))
    If settingCount <= 0 Then
        ActiveSetting = -1
        Exit Sub
    End If

    ReDim allSettings(0 To settingCount - 1)

    For i = 1 To settingCount
        allSettings(i - 1).sButtonColor = CLng(GetSetting(AppName, "Settings", "sButtonColor(" & i & ")", 0))
        allSettings(i - 1).sButtonFontColor = CLng(GetSetting(AppName, "Settings", "sButtonFontColor(" & i & ")", 0))
        For j = 0 To 2
            allSettings(i - 1).sButtonText(j) = CStr(GetSetting(AppName, "Settings", "sButtonText(" & i & ")(" & j & ")", ""))
            allSettings(i - 1).sActionPath(j) = CStr(GetSetting(AppName, "Settings", "sActionPath(" & i & ")(" & j & ")", ""))
            allSettings(i - 1).sActionFolder(j) = CStr(GetSetting(AppName, "Settings", "sActionFolder(" & i & ")(" & j & ")", ""))
            allSettings(i - 1).sActionParameters(j) = CStr(GetSetting(AppName, "Settings", "sActionParameters(" & i & ")(" & j & ")", ""))
        Next j

        allSettings(i - 1).sActivationText = CStr(GetSetting(AppName, "Settings", "sActivationText(" & i & ")", ""))
    Next i

    'other settings
    otherSettings.sUploadArguments = CStr(GetSetting(AppName, "Others", "sUploadArguments", ""))

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CpuUsage.Terminate

    tmrCheckHotKey.Enabled = False
    tmrComports.Enabled = False

    Unload frmGames
    Unload frmSettings
    Unload frmProgramTools

    Unload Me

End Sub

Private Sub tmrCheckHotKey_Timer()
    GetKeyboardState KeyStates(0)

    Dim i As Long



    If (KeyStates(vbKeyControl) And &H80) = &H80 Then
        If (KeyStates(vbKeyF12) And &H80) = &H80 And KeyFired(KEY_CONTROL_F12) = False Then
            KeyFired(KEY_CONTROL_F12) = True

            ShowMode = Not ShowMode
            CloseWindows

            If ShowMode Then
                Me.Hide
                SetParent Me.hWnd, ParentOrigional

                SetWindowPos Me.hWnd, -1, (ProgramLeft2 / Screen.TwipsPerPixelX), (ProgramTop / Screen.TwipsPerPixelY), 0, 0, SWP_NOSIZE Or SWP_SHOWWINDOW
                'ShowWindow Me.hwnd, 1
                'Me.SetFocus
                'txtSearch.SetFocus
            Else
                SetParent Me.hWnd, ParentTaskBar

                SetWindowPos Me.hWnd, 0, (ProgramLeft / Screen.TwipsPerPixelX), (ProgramTop / Screen.TwipsPerPixelY), 0, 0, SWP_NOSIZE Or SWP_SHOWWINDOW
            End If
        ElseIf (KeyStates(vbKeyF12) And &H80) = 0 Then
            KeyFired(KEY_CONTROL_F12) = False
        End If

        If (KeyStates(vbKeyF9) And &H80) = &H80 And KeyFired(KEY_CONTROL_F9) = False Then
            Dim MyAppHWnd As Long
            Dim CurrentForegroundThreadID As Long
            Dim NewForegroundThreadID As Long
            Dim lngRetVal As Long

            CurrentForegroundThreadID = GetWindowThreadProcessId(GetForegroundWindow(), ByVal 0&)
            NewForegroundThreadID = GetWindowThreadProcessId(Me.hWnd, ByVal 0&)

            'AttachThreadInput is used to ensure SetForegroundWindow will work
            'even if our application isn't currently the foreground window
            '(e.g. an automated app running in the background)
            Call AttachThreadInput(CurrentForegroundThreadID, NewForegroundThreadID, True)
            lngRetVal = SetForegroundWindow(MyAppHWnd)
            Call AttachThreadInput(CurrentForegroundThreadID, NewForegroundThreadID, False)

            SetForegroundWindow Me.hWnd
        ElseIf (KeyStates(vbKeyF9) And &H80) = 0 Then
            KeyFired(KEY_CONTROL_F9) = False
        End If
    Else
        KeyFired(KEY_CONTROL_F12) = False
        KeyFired(KEY_CONTROL_F9) = False
    End If

End Sub


Private Sub Form_Load()
    Dim i As Long
    Dim j As RECT
    Dim Ret As Long
    Dim lParam As Long

    'Set wmi = GetObject("winmgmts:\\.\root\CIMV2")

    SetProcessDPIAware

    'ret = RegisterHotKey(Me.hWnd, 40000, MOD_CONTROL, vbKeyF)

    'DeleteAllSettings

    LoadAllSettings

    'uTextBox1.SelBold = True

    If settingCount = 0 Then
        LoadDefaultSettings
    End If

    ButtonColors(0) = RGB(0, 0, 0)

    For i = 1 To 5
        ButtonColors(i) = RGB(i * 51, 0, 0)
        ButtonColors(i + 5) = RGB(0, i * 51, 0)
        ButtonColors(i + 10) = RGB(0, 0, i * 51)
        ButtonColors(i + 15) = RGB(i * 51, i * 51, 0)
        ButtonColors(i + 20) = RGB(i * 51, 0, i * 51)
        ButtonColors(i + 25) = RGB(0, i * 51, i * 51)
        ButtonColors(i + 30) = RGB(i * 51, i * 51, i * 51)

        ButtonColors(i + 35) = RGB(i * 25, i * 51, 0)
        ButtonColors(i + 40) = RGB(i * 25, 0, i * 51)
        ButtonColors(i + 45) = RGB(0, i * 25, i * 51)
        ButtonColors(i + 50) = RGB(i * 51, i * 25, 0)
        ButtonColors(i + 55) = RGB(0, i * 51, i * 25)
        ButtonColors(i + 60) = RGB(i * 51, 0, i * 25)
    Next i

    For i = 0 To 18
        ButtonColors(i + 61) = GetSysColor(i)
    Next i


    'EnumChildWindows ParentTaskBar, AddressOf EnumChildProc, lParam
    i = 0

    Do
        ParentOrigional = GetParent(Me.hWnd)
        ParentTaskBar = FindWindow("Shell_TrayWnd", vbNullString)
        EnumChildWindows ParentTaskBar, AddressOf EnumChildProc, lParam

        ParentTaskBar = MSTaskListWClass
        If i > 0 And ParentTaskBar = 0 Then
            If i > 6 Then
                MsgBox "Could not find the taskbar! Please restart the program", vbCritical
                Unload Me
                Exit Sub
            End If
            Wait 1000
        End If
        i = i + 1
    Loop While ParentTaskBar = 0

    GetWindowRect ParentTaskBar, j

    Me.Height = (j.Bottom - j.Top - 6) * Screen.TwipsPerPixelY
    Me.Width = 4000 '6050    '301 * Screen.TwipsPerPixelX

    ProgramTop = 4 * Screen.TwipsPerPixelY    '((j.Bottom - j.Top) / 2 - (Me.Height / 2)) ' - 10 * Screen.TwipsPerPixelY ' ((j.Bottom - j.Top) / 2 - 30 / 2) * 15
    ProgramLeft = (j.Right - j.Left) * Screen.TwipsPerPixelX - Me.Width    '1715 * Screen.TwipsPerPixelX - Me.Width
    ProgramLeft2 = ((Screen.Width / 2 - Me.Width / 2))

    ProgramParentLeft = j.Left

    Me.Top = ProgramTop
    Me.Left = ProgramLeft
    SetParent Me.hWnd, ParentTaskBar

    Set CpuUsage = New clsCPUUsageNT
    CpuUsage.Initialize

    Randomize
    For i = 1 To 30
        uCommPort.AddItem "Com" & i, , , -1, vbCenter
    Next i

    SetSearchingMode

    DoEvents
    
    tmrComports.Enabled = True
End Sub

Private Sub cmdExtra_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Exit Sub
    'If X <= 5 Then
    '    Dragging = True
    '    DragX = X    ' + (x + cmdExtra.Left) * Screen.TwipsPerPixelX
    'End If
End Sub

Private Sub cmdExtra_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Exit Sub
    If Dragging Then
        SetWindowPos Me.hWnd, 0, Me.Left / Screen.TwipsPerPixelX - ProgramParentLeft + (x - DragX), ProgramTop / Screen.TwipsPerPixelY, 0, 0, &H20 Or &H1 Or &H40

        If ShowMode Then
            ProgramLeft2 = Me.Left
        Else
            ProgramLeft = Me.Left
        End If
        'DoEvents
        Me.Refresh

        cmdExtra.MousePointer = 9
        If frmGames.Visible Then
            frmGames.Left = Me.Left
        End If
        If frmSettings.Visible Then
            frmSettings.Left = Me.Left
        End If
        If frmProgramTools.Visible Then
            frmProgramTools.Left = Me.Left
        End If
    Else
        If x <= 5 Then
            cmdExtra.MousePointer = 9
        Else
            cmdExtra.MousePointer = 0
        End If
    End If
End Sub

Private Sub cmdExtra_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Exit Sub
    'cmdExtra.MousePointer = 0
    'If Dragging = True Then
    '    Dragging = False
    'Else
        Unload Me
    'End If
End Sub

Private Sub Form_Resize()
    If Not Me.Visible Or Me.WindowState <> 0 Then Exit Sub
    On Error Resume Next

    txtSearch.Left = 1
    txtSearch.Width = Me.ScaleWidth - 1 - uLoadCPU.Width
    txtSearch.Top = 1
    txtSearch.Height = Me.ScaleHeight / 2 - 2

    uLoadCPU.Left = txtSearch.Width + txtSearch.Left + 1
    uLoadCPU.Width = Me.ScaleWidth - uLoadCPU.Left - 1
    uLoadCPU.Height = txtSearch.Height
    uLoadCPU.Top = 1
    uCommPort.ScrollBarWidth = uLoadCPU.Width + 1
    
    
    uCommPort.Width = 60 + uLoadCPU.Width + 1
    uCommPort.Left = Me.ScaleWidth - uCommPort.Width - 1
    uCommPort.Top = txtSearch.Top + txtSearch.Height + 1
    uCommPort.Height = Me.ScaleHeight - uCommPort.Top - 1
    
    
    
    
    
    cmdExtra.Top = uCommPort.Top
    cmdExtra.Left = txtSearch.Left
    cmdExtra.Height = Me.ScaleHeight - cmdExtra.Top - 1
    cmdExtra.Width = uCommPort.Height
    
    cmdSettings.Top = cmdExtra.Top
    cmdSettings.Height = cmdExtra.Height

    cmdSettings.Width = Fix((uCommPort.Left - cmdExtra.Width - 3) / 3)
    cmdSettings.Left = cmdExtra.Left + cmdExtra.Width + 1


    cmdGames.Top = cmdExtra.Top
    cmdGames.Height = cmdExtra.Height

    cmdGames.Width = Fix((uCommPort.Left - cmdExtra.Width - 3) / 3)
    cmdGames.Left = cmdSettings.Left + cmdSettings.Width + 1


    cmdPrograms.Top = cmdExtra.Top
    cmdPrograms.Height = cmdExtra.Height

    cmdPrograms.Width = uCommPort.Left - (cmdGames.Left + cmdGames.Width + 2)
    cmdPrograms.Left = cmdGames.Left + cmdGames.Width + 1


    

    


End Sub

Sub SetSearchingMode()

    If ActiveSetting > -1 Then

        If allSettings(ActiveSetting).sButtonText(0) <> "" Then
            cmdSettings.Caption = allSettings(ActiveSetting).sButtonText(0)
        Else
            cmdSettings.Caption = "Settings"
        End If

        If allSettings(ActiveSetting).sButtonText(1) <> "" Then
            cmdGames.Caption = allSettings(ActiveSetting).sButtonText(1)
        Else
            cmdGames.Caption = "Games"
        End If

        If allSettings(ActiveSetting).sButtonText(2) <> "" Then
            cmdPrograms.Caption = allSettings(ActiveSetting).sButtonText(2)
        Else
            cmdPrograms.Caption = "Programs"
        End If

        If allSettings(ActiveSetting).sButtonColor = -1 Then allSettings(ActiveSetting).sButtonColor = 0
        SetProgramColor ButtonColors(allSettings(ActiveSetting).sButtonColor), ButtonColors(allSettings(ActiveSetting).sButtonFontColor)
    End If
End Sub

Sub SetProgramColor(lBackColor As Long, lFontColor As Long)
    Dim i As Long

    Me.BackColor = lBackColor
    txtSearch.ForeColor = lBackColor

    cmdSettings.BackgroundColor = lBackColor
    cmdSettings.ForeColor = lFontColor

    cmdGames.BackgroundColor = lBackColor
    cmdGames.ForeColor = lFontColor

    cmdPrograms.BackgroundColor = lBackColor
    cmdPrograms.ForeColor = lFontColor

    cmdExtra.BackgroundColor = lBackColor
    cmdExtra.ForeColor = lFontColor

    frmSettings.BackColor = lBackColor
    frmGames.BackColor = lBackColor
    frmProgramTools.BackColor = lBackColor

    frmGames.cmdClose.BackgroundColor = lBackColor
    frmGames.cmdClose.ForeColor = lFontColor

    uCommPort.BackgroundColor = lBackColor
    uCommPort.ForeColor = lFontColor
    uCommPort.SelectionBackgroundColor = lBackColor

    'uLoadBar1.BackgroundColor = &H584D43
    'uLoadBar1.BarColor = lBackColor

    '    For i = 0 To uLoadCPU.Count - 1
    '        uLoadCPU(i).BarColor = lBackColor
    '        uLoadCPU(i).BackgroundColor = lFontColor
    '        uLoadCPU(i).ForeColor = lFontColor
    '        uLoadCPU(i).CaptionBorderColor = lBackColor
    '    Next i

    uLoadCPU.BarColor = lBackColor
    uLoadCPU.BackgroundColor = lFontColor
    uLoadCPU.ForeColor = lFontColor
    uLoadCPU.CaptionBorderColor = lBackColor

    For i = 0 To 3
        frmGames.cmdStart(i).BackgroundColor = lBackColor
        frmGames.cmdStart(i).ForeColor = lFontColor
    Next i


    Dim L As Control

    For Each L In frmProgramTools.Controls
        If Left$(L.Name, 2) = "No" Then GoTo Next_frmProgramTools:
        Select Case TypeName(L)
            Case "uFrame", "uDropDown"
                L.BackgroundColor = lBackColor
                L.ForeColor = lFontColor
                If TypeName(L) = "uDropDown" Then
                    L.SelectionBackgroundColor = lBackColor
                End If
            Case "uCheckBox", "uOptionBox"
                L.BackgroundColor = lBackColor
                L.ForeColor = lFontColor
                L.CheckBackgroundColor = lBackColor
                L.CheckBorderColor = lFontColor
                L.CheckSelectionColor = lFontColor
            
            Case "uLoadBar"
                L.BackgroundColor = lBackColor
                
            Case "Label"
                L.BackColor = lBackColor
                L.ForeColor = lFontColor

            Case "Line"
                L.BorderColor = lFontColor

            Case "uButton"
                L.BackgroundColor = lBackColor
                L.ForeColor = lFontColor
                L.FocusColor = lBackColor

            Case "PictureBox"
                L.BackColor = lBackColor

        End Select
Next_frmProgramTools:
    Next


    For Each L In frmSettings.Controls
        Select Case TypeName(L)
            Case "uFrame", "uDropDown"
                L.BackgroundColor = lBackColor
                L.ForeColor = lFontColor
                If TypeName(L) = "uDropDown" Then
                    L.SelectionBackgroundColor = lBackColor
                End If
            Case "Label"
                L.BackColor = lBackColor
                L.ForeColor = lFontColor
            Case "uButton"
                L.BackgroundColor = lBackColor
                L.ForeColor = lFontColor
        End Select
    Next

    tmrComports_Timer

End Sub


Private Sub tmrComports_Timer()
    Dim tmpstr() As Long
    On Error Resume Next

    If frmProgramTools.uCommRefresh.Value = u_Checked Then
        tmpstr = EnumSerialPorts

        Dim i As Long

        uCommPort.RedrawPause

        For i = 0 To uCommPort.ListCount - 1
            uCommPort.ItemColor(i) = -1 'ButtonColors(allSettings(ActiveSetting).sButtonColor)
        Next i

        For i = 0 To UBound(tmpstr)
            If tmpstr(i) > 0 Then
                uCommPort.ItemColor(tmpstr(i) - 1) = vbGreen
            End If
        Next i

        uCommPort.RedrawResume
    End If

    If ShowMode = False Then
        Dim j As RECT
        GetWindowRect ParentTaskBar, j


        ProgramLeft = (j.Right - j.Left) * Screen.TwipsPerPixelX - Me.Width    '- 100

        SetWindowPos Me.hWnd, 0, (ProgramLeft / Screen.TwipsPerPixelX), (ProgramTop / Screen.TwipsPerPixelY), 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE
    End If

    uLoadCPU.Value = CpuUsage.Query

End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then
        If Len(txtSearch.Text) = 0 Then
            ActiveSetting = 0
            SetSearchingMode
        End If
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = 0
        cmdSettings_Click 0, 0, 0
        txtSearch.Text = ""
        KeyCode = 0
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim sSearch As String


    If KeyCode = vbKeySpace Then
        If Len(txtSearch.Text) = 2 Then
            If settingCount = 0 Then Exit Sub
            sSearch = LCase(Left$(txtSearch.Text, 1))
            ActiveSetting = 0

            For i = 0 To UBound(allSettings)
                If sSearch = LCase(allSettings(i).sActivationText) Then
                    ActiveSetting = i
                    txtSearch.Text = ""
                    Exit For
                End If

            Next i

            SetSearchingMode
        End If
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = 0
    End If
End Sub

Function ExecuteSearch() As Boolean
    Dim lPath As String
    Dim lFolder As String
    Dim lParameters As String

    ExecuteSearch = False





    txtSearch.Text = ""
    ActiveSetting = 0
    SetSearchingMode


    Exit Function

    '            On Error GoTo errhandler
    '
    '            ' -p m2560 -D -c stk500v2 -P com9 -b 115200 -F -V -U flash:w:main.hex
    '
    '            comm.CommPort = uCommPort.ListIndex + 1
    '            comm.Settings = "19200,n,8,1"
    '            comm.PortOpen = True
    '
    '
    '            Wait 10
    '            If comm.PortOpen = True Then
    '                comm.DTREnable = False
    '                Wait 20
    '                comm.DTREnable = True
    '                Wait 20
    '                comm.PortOpen = False
    '
    '                ShellExecute 0, "open", "cmd.exe", "/c avrdude.exe  -p m2560 -D -c stk500v2 -P com" & (uCommPort.ListIndex + 1) & " -b 115200 -F -V -U flash:w:main.hex", "C:\Users\ricardo\Documents\GitHub\swagbot\software" & "\arduino\src\", vbNormalFocus
    '
    '            End If
    '
    '            Exit Function
    '
    '        Case Else
    '            ExecuteSearch = True
    '            Exit Function
    '    End Select



    Exit Function

ErrHandler:
    'MsgBox "Could not connect to Com" & uCommPort.ListIndex + 1

End Function


Private Sub uCheckBox1_ActivateNextState(u_Cancel As Boolean, u_NewState As uCheckboxConstants)
    u_Cancel = True

    If u_NewState = u_Checked Then
        u_NewState = u_UnChecked
    ElseIf u_NewState = u_UnChecked Then
        u_NewState = u_Cross
    ElseIf u_NewState = u_Cross Then
        u_NewState = u_PartialChecked
    Else
        u_NewState = u_Checked
    End If
End Sub

