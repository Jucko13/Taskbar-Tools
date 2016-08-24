VERSION 5.00
Begin VB.Form frmGames 
   BackColor       =   &H00584D43&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   84
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   627
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTopMost 
      Interval        =   100
      Left            =   1800
      Top             =   270
   End
   Begin Project1.uButton cmdStart 
      Height          =   810
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   1429
      Caption         =   "Vindictus"
      BorderAnimation =   4
      sPicture        =   "Vindictus.ico"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionOffsetTop=   2
      MouseOverBackgroundColor=   8421504
   End
   Begin Project1.uButton cmdStart 
      Height          =   810
      Index           =   1
      Left            =   1140
      TabIndex        =   1
      Top             =   15
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   1429
      Caption         =   "Battlefield 3"
      BorderAnimation =   4
      sPicture        =   "Bf3.ico"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionOffsetTop=   2
      MouseOverBackgroundColor=   8421504
   End
   Begin Project1.uButton cmdClose 
      Height          =   810
      Left            =   4515
      TabIndex        =   2
      Top             =   15
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   1429
      Caption         =   "Close"
      BorderAnimation =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseOverBackgroundColor=   8421504
   End
   Begin Project1.uButton cmdStart 
      Height          =   810
      Index           =   3
      Left            =   3390
      TabIndex        =   3
      Top             =   15
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   1429
      Caption         =   "Call of Duty 2"
      BorderAnimation =   4
      sPicture        =   "cod2.ico"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionOffsetTop=   2
      MouseOverBackgroundColor=   8421504
   End
   Begin Project1.uButton cmdStart 
      Height          =   810
      Index           =   2
      Left            =   2265
      TabIndex        =   4
      Top             =   15
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   1429
      Caption         =   "LoL"
      BorderAnimation =   4
      sPicture        =   "lol.ico"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionOffsetTop=   2
      MouseOverBackgroundColor=   8421504
   End
End
Attribute VB_Name = "frmGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdClose_Click(Button As Integer, X As Single, Y As Single)
    Me.Hide
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Dim i As Long

    For i = 0 To cmdStart.Count - 1
        cmdStart(i).Width = Fix(Me.ScaleWidth / 5)
        cmdStart(i).Left = (cmdStart(i).Width + 1) * i + 1
    Next i

    cmdClose.Left = (cmdStart(cmdStart.Count - 1).Width + 1) * i + 1
    cmdClose.Width = Me.ScaleWidth - cmdClose.Left - 1
End Sub

Sub tmrTopMost_Timer()
    If Me.Visible Then SetTopMostWindow frmGames.hwnd, True
End Sub

Private Sub cmdStart_Click(Index As Integer, Button As Integer, X As Single, Y As Single)
    Dim tmpFolderPath As String
    Dim tmpExecutablePath As String
    Dim tmpParameters As String

    Select Case Index

        Case 0
            tmpFolderPath = "D:\Program Files\Vindictus\en-EU\"
            tmpExecutablePath = "D:\Program Files\Vindictus\en-EU\Vindictus.exe"
            tmpParameters = ""

        Case 1
            tmpFolderPath = ""
            tmpExecutablePath = "chrome.exe"
            tmpParameters = """http://battlelog.battlefield.com/bf3/servers/"""

        Case 2
            tmpFolderPath = "C:\Program Files\Riot Games\League of Legends\"
            tmpExecutablePath = "C:\Program Files\Riot Games\League of Legends\lol.launcher.admin.exe"
            tmpParameters = ""

        Case 3
            tmpFolderPath = "D:\Program Files\Activision\Call of Duty 2\"
            tmpExecutablePath = "D:\Program Files\Activision\Call of Duty 2\CoD2MP_s.exe"
            tmpParameters = ""

    End Select

    ShellExecute Me.hwnd, "OPEN", tmpExecutablePath, tmpParameters, tmpFolderPath, vbNormalFocus
    Me.Hide
End Sub
