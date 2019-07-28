VERSION 5.00
Begin VB.UserControl uToolTip 
   BackColor       =   &H001AA6FA&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H001AA6FA&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picTooltip 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H001AA6FA&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1110
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   1920
      Begin VB.Label lblTooltip 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0024211E&
         Caption         =   "Hi this is a help caption"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001AA6FA&
         Height          =   195
         Left            =   45
         TabIndex        =   1
         Top             =   15
         Width           =   1830
      End
   End
   Begin VB.Label lblIcon 
      Alignment       =   2  'Center
      BackColor       =   &H0024211E&
      Caption         =   " TT "
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001AA6FA&
      Height          =   390
      Left            =   15
      TabIndex        =   2
      Top             =   15
      Width           =   750
   End
End
Attribute VB_Name = "uToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private m_poiMousePos As POINTAPI
Private m_poiPreviousMousePos As POINTAPI

Private WithEvents m_tmrCheckControls As Timer
Attribute m_tmrCheckControls.VB_VarHelpID = -1

Private m_colData As New Collection
Private m_lonCountHWND As Long
Private m_lonFormHWND As Long
'Private m_lonPreviousHWND As Long

Public Sub setForm(frm As Form)
    m_lonFormHWND = frm.hWnd
End Sub

Public Sub Add(hWnd As Long, txt As String)
    'Debug.Print "added:" & hWnd
    m_colData.Add txt, CStr(hWnd)
End Sub

Private Sub m_tmrCheckControls_Timer()
On Error GoTo ExistsNonObjectErrorHandler

    Dim currentHWND As Long
    Dim whatToShow As String
    
    GetCursorPos m_poiMousePos
    
    currentHWND = WindowFromPoint(m_poiMousePos.X, m_poiMousePos.Y)
    whatToShow = m_colData(CStr(currentHWND))
    
    If m_poiMousePos.X = m_poiPreviousMousePos.X And m_poiMousePos.Y = m_poiPreviousMousePos.Y Then
        m_lonCountHWND = m_lonCountHWND + 1
        
        If m_lonCountHWND = 2 Then
            Dim r As RECT
            Dim fr As RECT
            
            GetWindowRect currentHWND, r
            GetWindowRect m_lonFormHWND, fr
            
            lblTooltip.Caption = whatToShow
            lblTooltip.Width = lblTooltip.Width + 20
            lblTooltip.Height = lblTooltip.Height + 1
            lblTooltip.Left = 1
            lblTooltip.Top = 1
            picTooltip.Width = (lblTooltip.Width + 2) * Screen.TwipsPerPixelX
            picTooltip.Height = (lblTooltip.Height + 2) * Screen.TwipsPerPixelY
            picTooltip.Left = m_poiMousePos.X * Screen.TwipsPerPixelX ' - (r.Left - fr.Left)
            picTooltip.Top = (m_poiMousePos.Y + 20) * Screen.TwipsPerPixelY '- (r.Top - fr.Top)
            'Extender.Left = m_poiMousePos.X - fr.Left
            'Extender.Top = m_poiMousePos.Y - fr.Top
            
            picTooltip.Visible = True
        End If
    Else
        GoTo ExistsNonObjectErrorHandler
    End If
    
    Exit Sub
    
ExistsNonObjectErrorHandler:
    m_lonCountHWND = 0
    picTooltip.Visible = False
    lblTooltip.Caption = ""
    m_poiPreviousMousePos = m_poiMousePos
End Sub

Private Sub UserControl_Initialize()
    Set m_tmrCheckControls = UserControl.Controls.Add("VB.Timer", "m_tmrMouseOver")
    m_tmrCheckControls.Interval = 200
    m_tmrCheckControls.Enabled = False
    
    picTooltip.Visible = False
    SetParent picTooltip.hWnd, GetParent(0)
    SetTopMostWindow picTooltip.hWnd, True
    
    SetWindowLong picTooltip.hWnd, -20, GetWindowLong(picTooltip.hWnd, -20) Or &H80&
    
    'Debug.Print Screen.Width; Screen.Height
End Sub

Sub StartTimer()
    m_tmrCheckControls.Enabled = True

End Sub

Sub Redraw()

End Sub

Private Sub UserControl_Resize()
    Extender.Width = lblIcon.Width + 2
    Extender.Height = lblIcon.Height + 2
End Sub

Private Sub UserControl_Terminate()
    SetParent picTooltip.hWnd, GetParent(UserControl.hWnd)
End Sub
