VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmProgramTools 
   AutoRedraw      =   -1  'True
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   11355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   757
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1386
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3285
      Index           =   1
      Left            =   12795
      ScaleHeight     =   3285
      ScaleWidth      =   7005
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   7005
      Begin Project1.uListBox uProcess 
         Height          =   2880
         Left            =   30
         TabIndex        =   76
         Top             =   30
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   5080
         BackgroundColor =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   65280
         Text            =   ""
         SelectionBorderColor=   32768
         SelectionForeColor=   65280
         ItemHeight      =   19
         VisibleItems    =   10
      End
      Begin Project1.uButton uRefreshProcess 
         Height          =   315
         Left            =   30
         TabIndex        =   77
         Top             =   2940
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         FocusColor      =   0
         BackgroundColorDisabled=   0
         BorderColorDisabled=   0
         ForeColorDisabled=   0
         CaptionBorderColorDisabled=   0
         FocusColorDisabled=   0
         Caption         =   "Refresh"
         BorderAnimation =   0
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
   End
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3410
      Index           =   4
      Left            =   9780
      ScaleHeight     =   227
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   467
      TabIndex        =   66
      Top             =   4830
      Visible         =   0   'False
      Width           =   7005
      Begin Project1.uLoadBar loadScores 
         Height          =   1680
         Left            =   615
         TabIndex        =   69
         Top             =   780
         Visible         =   0   'False
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   2963
         BackgroundColor =   5786947
         BarColor        =   65280
         BarWidth        =   10
         BorderColor     =   0
         Caption         =   "Loading Scores"
         CaptionBorder   =   -1  'True
         CaptionBorderColor=   0
         CaptionType     =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   65280
      End
      Begin Project1.uButton cmdSort 
         Height          =   200
         Index           =   0
         Left            =   60
         TabIndex        =   71
         Top             =   3140
         Width           =   1400
         _ExtentX        =   2461
         _ExtentY        =   344
         FocusColor      =   0
         BackgroundColorDisabled=   0
         BorderColorDisabled=   0
         ForeColorDisabled=   0
         CaptionBorderColorDisabled=   0
         FocusColorDisabled=   0
         Caption         =   "Sort By Date"
         BorderAnimation =   0
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
      Begin VB.VScrollBar scrollScores 
         Height          =   3290
         LargeChange     =   10
         Left            =   6765
         Max             =   10
         TabIndex        =   70
         Top             =   60
         Width           =   180
      End
      Begin Project1.uButton cmdSort 
         Height          =   200
         Index           =   1
         Left            =   1530
         TabIndex        =   72
         Top             =   3140
         Width           =   1400
         _ExtentX        =   2461
         _ExtentY        =   344
         FocusColor      =   0
         BackgroundColorDisabled=   0
         BorderColorDisabled=   0
         ForeColorDisabled=   0
         CaptionBorderColorDisabled=   0
         FocusColorDisabled=   0
         Caption         =   "Sort By Score"
         BorderAnimation =   0
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
      Begin Project1.uButton cmdRefreshScore 
         Height          =   200
         Left            =   3000
         TabIndex        =   73
         Top             =   3140
         Width           =   1400
         _ExtentX        =   2461
         _ExtentY        =   344
         FocusColor      =   0
         BackgroundColorDisabled=   0
         BorderColorDisabled=   0
         ForeColorDisabled=   0
         CaptionBorderColorDisabled=   0
         FocusColorDisabled=   0
         Caption         =   "Refresh"
         BorderAnimation =   0
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
   End
   Begin VB.Timer tmrAdapterRefresh 
      Interval        =   1000
      Left            =   4725
      Top             =   2880
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00584D43&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Index           =   3
      Left            =   6765
      ScaleHeight     =   3120
      ScaleWidth      =   5625
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   825
      Visible         =   0   'False
      Width           =   5625
      Begin VB.TextBox txtPenKeyStroke 
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   1
         Left            =   1590
         TabIndex        =   61
         Text            =   "{left}"
         Top             =   1935
         Width           =   3945
      End
      Begin VB.TextBox txtPenKeyStroke 
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   0
         Left            =   1590
         TabIndex        =   60
         Text            =   "{right}"
         Top             =   1710
         Width           =   3945
      End
      Begin VB.TextBox txtPenProgram 
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   1
         Left            =   1575
         TabIndex        =   59
         Text            =   "mspaint.exe"
         Top             =   1020
         Width           =   3945
      End
      Begin VB.TextBox txtPenProgram 
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   0
         Left            =   1575
         TabIndex        =   58
         Text            =   "notepad.exe"
         Top             =   795
         Width           =   3945
      End
      Begin Project1.uOptionBox uPenOption 
         Height          =   285
         Index           =   0
         Left            =   330
         TabIndex        =   53
         Top             =   465
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   503
         BackgroundColor =   5786947
         Caption         =   "Start Programs:"
         CaptionOffsetTop=   1
         CheckBackgroundColor=   5786947
         CheckBorderColor=   16777215
         CheckSelectionColor=   16777215
         CheckSize       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin Project1.uCheckBox uMicrosoftPen 
         Height          =   315
         Left            =   105
         TabIndex        =   52
         Top             =   90
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   556
         BackgroundColor =   5786947
         Caption         =   "Enable Microsoft Pen Function Overriding"
         CaptionOffsetLeft=   10
         CaptionOffsetTop=   1
         CheckBackgroundColor=   5786947
         CheckBorderColor=   16777215
         CheckSelectionColor=   16777215
         CheckSize       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin Project1.uOptionBox uPenOption 
         Height          =   285
         Index           =   1
         Left            =   330
         TabIndex        =   54
         Top             =   1380
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   503
         BackgroundColor =   5786947
         Caption         =   "Send Keystrokes"
         CaptionOffsetTop=   1
         CheckBackgroundColor=   5786947
         CheckBorderColor=   16777215
         CheckSelectionColor=   16777215
         CheckSize       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   435
         X2              =   240
         Y1              =   1515
         Y2              =   1515
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   435
         X2              =   240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   255
         X2              =   255
         Y1              =   390
         Y2              =   1530
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Double Press:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   495
         TabIndex        =   57
         Top             =   1920
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single Press:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   495
         TabIndex        =   56
         Top             =   1695
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Double Press:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   495
         TabIndex        =   55
         Top             =   1005
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Single Press:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   495
         TabIndex        =   50
         Top             =   780
         Width           =   1080
      End
   End
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   3300
      Index           =   2
      Left            =   255
      ScaleHeight     =   220
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   467
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   7275
      Visible         =   0   'False
      Width           =   7005
      Begin Project1.uButton uSetIp 
         Height          =   240
         Left            =   5250
         TabIndex        =   84
         Top             =   855
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   423
         BackgroundColor =   5786947
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8421504
         Caption         =   "Set IP"
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
      Begin Project1.uButton uSetDNS 
         Height          =   240
         Left            =   5250
         TabIndex        =   83
         Top             =   1155
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   423
         BackgroundColor =   5786947
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8421504
         Caption         =   "Set DNS"
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
      Begin VB.TextBox txtSetIP 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   75
         TabIndex        =   82
         Top             =   1155
         Width           =   1650
      End
      Begin VB.TextBox txtSetIP 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   4
         Left            =   1785
         TabIndex        =   81
         Top             =   1155
         Width           =   1665
      End
      Begin Project1.uButton uRenew 
         Height          =   270
         Left            =   2655
         TabIndex        =   80
         Top             =   2970
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   476
         BackgroundColor =   5786947
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8421504
         Caption         =   "Renew"
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
      Begin Project1.uButton uCloseStats 
         Height          =   270
         Left            =   6420
         TabIndex        =   79
         Top             =   45
         Visible         =   0   'False
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   476
         BackgroundColor =   5786947
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8421504
         Caption         =   "X"
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
      Begin VB.TextBox txtPingResult 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   810
         Left            =   195
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   1580
         Visible         =   0   'False
         Width           =   6525
      End
      Begin Project1.uListBox lstIP 
         Height          =   705
         Index           =   0
         Left            =   75
         TabIndex        =   63
         Top             =   90
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1244
         BorderColor     =   12632256
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
         Border          =   0   'False
         SelectionBackgroundColor=   16764768
         SelectionBorderColor=   16764768
         ItemHeight      =   14
         VisibleItems    =   3
      End
      Begin VB.TextBox txtSetIP 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   3525
         TabIndex        =   46
         Top             =   855
         Width           =   1665
      End
      Begin VB.TextBox txtSetIP 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   1785
         TabIndex        =   45
         Top             =   855
         Width           =   1665
      End
      Begin VB.TextBox txtSetIP 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   44
         Top             =   855
         Width           =   1650
      End
      Begin Project1.uButton uEnableDisable 
         Height          =   270
         Index           =   0
         Left            =   4815
         TabIndex        =   36
         Top             =   2970
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   476
         BackgroundColor =   5786947
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8421504
         Caption         =   "Disable"
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
      Begin Project1.uButton uRefreshIp 
         Height          =   270
         Left            =   45
         TabIndex        =   38
         Top             =   2970
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   476
         BackgroundColor =   5786947
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8421504
         Caption         =   "Refresh List"
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
      Begin Project1.uButton uChangeIP 
         Height          =   270
         Left            =   1350
         TabIndex        =   39
         Top             =   2970
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   476
         BackgroundColor =   5786947
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8421504
         Caption         =   "Set Automatic"
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
      Begin Project1.uButton uPing 
         Height          =   240
         Left            =   3510
         TabIndex        =   47
         Top             =   1155
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   423
         BackgroundColor =   5786947
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8421504
         Caption         =   "Ping"
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
      Begin Project1.uListBox lstAdapter 
         Height          =   1080
         Left            =   75
         TabIndex        =   64
         Top             =   1455
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   1905
         BorderColor     =   0
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
         Border          =   0   'False
         SelectionBackgroundColor=   16764768
         ItemHeight      =   14
      End
      Begin Project1.uButton uEnableDisable 
         Height          =   270
         Index           =   1
         Left            =   5895
         TabIndex        =   65
         Top             =   2970
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   476
         BackgroundColor =   5786947
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8421504
         Caption         =   "Enable"
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
      Begin Project1.uButton uStats 
         Height          =   270
         Left            =   3735
         TabIndex        =   74
         Top             =   2970
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   476
         BackgroundColor =   5786947
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8421504
         Caption         =   "Stats"
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
      Begin VB.TextBox txtStats 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   3195
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   78
         Top             =   45
         Visible         =   0   'False
         Width           =   6915
      End
      Begin VB.Label lblIPSettings 
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   4
         Left            =   1005
         TabIndex        =   43
         Top             =   2550
         Width           =   1020
      End
      Begin VB.Label lblIPSettings 
         BackStyle       =   0  'Transparent
         Caption         =   "DefaultGateway:"
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Index           =   3
         Left            =   5565
         TabIndex        =   42
         Top             =   2550
         Width           =   1320
      End
      Begin VB.Label lblIPSettings 
         BackStyle       =   0  'Transparent
         Caption         =   "SubnetMask:"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   2
         Left            =   3690
         TabIndex        =   41
         Top             =   2550
         Width           =   1395
      End
      Begin VB.Label lblIPSettings 
         BackStyle       =   0  'Transparent
         Caption         =   "IP:"
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   1
         Left            =   2085
         TabIndex        =   40
         Top             =   2550
         Width           =   1200
      End
      Begin VB.Label lblIPSettings 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected IP:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   37
         Top             =   2550
         Width           =   870
      End
   End
   Begin Project1.uButton uClose 
      Height          =   255
      Left            =   6090
      TabIndex        =   14
      Top             =   45
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   450
      BackgroundColor =   5786947
      ForeColor       =   16777215
      Caption         =   "X"
      BorderAnimation =   0
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
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00584D43&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2880
      Index           =   0
      Left            =   120
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   466
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4305
      Visible         =   0   'False
      Width           =   6990
      Begin Project1.uCheckBox uComAscii 
         Height          =   195
         Left            =   3180
         TabIndex        =   29
         Top             =   585
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   344
         BackgroundColor =   5786947
         Caption         =   "Show Ascii Charvalues"
         CaptionOffsetTop=   1
         CheckSize       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.TextBox txtComSplit 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   2235
         TabIndex        =   28
         Text            =   "FF"
         Top             =   45
         Width           =   300
      End
      Begin VB.TextBox txtComSplit 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   1920
         TabIndex        =   27
         Text            =   "FF"
         Top             =   45
         Width           =   300
      End
      Begin VB.TextBox txtComSplit 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   1605
         TabIndex        =   26
         Text            =   "FF"
         Top             =   45
         Width           =   300
      End
      Begin VB.TextBox txtComSplit 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   1290
         TabIndex        =   25
         Text            =   "FF"
         Top             =   45
         Width           =   300
      End
      Begin VB.TextBox txtComSplit 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   975
         TabIndex        =   24
         Text            =   "FF"
         Top             =   45
         Width           =   300
      End
      Begin VB.TextBox txtComSplit 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   660
         TabIndex        =   23
         Text            =   "FF"
         Top             =   45
         Width           =   300
      End
      Begin VB.TextBox txtComSplit 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   345
         TabIndex        =   22
         Text            =   "FF"
         Top             =   45
         Width           =   300
      End
      Begin VB.TextBox txtComSplit 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   30
         TabIndex        =   21
         Text            =   "FF"
         Top             =   45
         Width           =   300
      End
      Begin Project1.uButton uConnect 
         Height          =   270
         Left            =   3570
         TabIndex        =   19
         Top             =   1995
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   476
         BackgroundColor =   5786947
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8421504
         Caption         =   "Connect"
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
      Begin Project1.uDropDown uDropBaud 
         Height          =   225
         Left            =   4530
         TabIndex        =   18
         Top             =   300
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   397
         BackgroundColor =   5786947
         SelectionBackgroundColor=   5786947
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
         ItemHeight      =   19
      End
      Begin VB.TextBox txtChar 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4530
         TabIndex        =   16
         Top             =   45
         Width           =   1410
      End
      Begin VB.TextBox txtCom 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2580
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "Form4.frx":0000
         Top             =   285
         Width           =   2835
      End
      Begin Project1.uOptionBox uOptionBox1 
         Height          =   195
         Index           =   0
         Left            =   3180
         TabIndex        =   30
         Top             =   840
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   344
         BackgroundColor =   5786947
         CheckBackgroundColor=   5786947
         CheckBorderColor=   16777215
         CheckSelectionColor=   16777215
         CheckSize       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin Project1.uOptionBox uOptionBox1 
         Height          =   195
         Index           =   1
         Left            =   3180
         TabIndex        =   31
         Top             =   1095
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   344
         BackgroundColor =   5786947
         CheckBackgroundColor=   5786947
         CheckBorderColor=   16777215
         CheckSelectionColor=   16777215
         CheckSize       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin VB.PictureBox NopicBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   0
         ScaleHeight     =   150
         ScaleWidth      =   2565
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   75
         Width           =   2565
      End
      Begin VB.PictureBox NopicBackText 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2580
         Left            =   0
         ScaleHeight     =   2580
         ScaleWidth      =   240
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
      End
      Begin Project1.uCheckBox uCommRefresh 
         Height          =   200
         Left            =   3180
         TabIndex        =   62
         Top             =   1340
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   344
         BackgroundColor =   5786947
         Caption         =   "Refresh commports"
         CaptionOffsetTop=   1
         CheckSize       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin VB.Label lblCom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baud Rate:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   3210
         TabIndex        =   20
         Top             =   285
         Width           =   1410
      End
      Begin VB.Label lblCom 
         BackStyle       =   0  'Transparent
         Caption         =   "Split Char (ASCII):"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   3195
         TabIndex        =   17
         Top             =   30
         Width           =   1410
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00584D43&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Index           =   10
      Left            =   570
      ScaleHeight     =   3120
      ScaleWidth      =   5625
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   810
      Visible         =   0   'False
      Width           =   5625
      Begin VB.Timer tmrTopMost 
         Interval        =   100
         Left            =   3735
         Top             =   2070
      End
      Begin VB.TextBox txtMasterIP 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   345
         TabIndex        =   2
         Text            =   "192.168.0.6"
         Top             =   1005
         Width           =   4200
      End
      Begin VB.Timer tmrMouse 
         Enabled         =   0   'False
         Interval        =   3
         Left            =   3315
         Top             =   2070
      End
      Begin VB.HScrollBar hSpeed 
         Height          =   270
         LargeChange     =   100
         Left            =   345
         Max             =   500
         TabIndex        =   1
         Top             =   1485
         Value           =   100
         Width           =   3090
      End
      Begin VB.Timer tmrConnection 
         Interval        =   200
         Left            =   2895
         Top             =   2070
      End
      Begin MSWinsockLib.Winsock socket 
         Index           =   0
         Left            =   1635
         Top             =   2070
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin Project1.uCheckBox uMouseEnabled 
         Height          =   315
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   556
         BackgroundColor =   5786947
         Caption         =   "Enable Mouse Capture over Ethernet"
         CaptionOffsetLeft=   10
         CaptionOffsetTop=   1
         CheckBackgroundColor=   5786947
         CheckBorderColor=   16777215
         CheckSelectionColor=   16777215
         CheckSize       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin Project1.uCheckBox uSlave 
         Height          =   315
         Left            =   345
         TabIndex        =   4
         Top             =   675
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   556
         BackgroundColor =   5786947
         Caption         =   "This PC is the slave"
         CaptionOffsetLeft=   10
         CaptionOffsetTop=   1
         CheckBackgroundColor=   5786947
         CheckBorderColor=   16777215
         CheckSelectionColor=   16777215
         CheckSize       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin Project1.uCheckBox uMaster 
         Height          =   315
         Left            =   345
         TabIndex        =   5
         Top             =   345
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   556
         BackgroundColor =   5786947
         Caption         =   "This PC is the Master"
         CaptionOffsetLeft=   10
         CaptionOffsetTop=   1
         CheckBackgroundColor=   5786947
         CheckBorderColor=   16777215
         CheckSelectionColor=   16777215
         CheckSize       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin Project1.uButton cmdConnect 
         Height          =   240
         Left            =   345
         TabIndex        =   6
         Top             =   1230
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   423
         MouseOverBackgroundColor=   8421504
         CaptionBorderColor=   16711680
         Caption         =   "Connect"
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
      Begin MSWinsockLib.Winsock socket 
         Index           =   1
         Left            =   2055
         Top             =   2070
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock socket 
         Index           =   2
         Left            =   2475
         Top             =   2070
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   165
         X2              =   165
         Y1              =   270
         Y2              =   1350
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   360
         X2              =   165
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   360
         X2              =   165
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   345
         X2              =   150
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   345
         X2              =   150
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed: 1x"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3540
         TabIndex        =   10
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Socket1: Closed"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   9
         Top             =   2460
         Width           =   2265
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Socket2: Closed"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   8
         Top             =   2640
         Width           =   2265
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Socket3: Closed"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   30
         TabIndex        =   7
         Top             =   2820
         Width           =   2265
      End
   End
   Begin Project1.uButton uMenu 
      Height          =   360
      Index           =   1
      Left            =   45
      TabIndex        =   12
      Top             =   45
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   635
      BackgroundColor =   5786947
      ForeColor       =   16777215
      MouseOverBackgroundColor=   8421504
      Caption         =   "Process Explorer"
      BorderAnimation =   0
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
   Begin Project1.uButton uMenu 
      Height          =   285
      Index           =   0
      Left            =   1365
      TabIndex        =   13
      Top             =   45
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   503
      BackgroundColor =   5786947
      ForeColor       =   16777215
      MouseOverBackgroundColor=   8421504
      Caption         =   "COMM"
      BorderAnimation =   0
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
   Begin Project1.uButton uMenu 
      Height          =   285
      Index           =   2
      Left            =   2685
      TabIndex        =   35
      Top             =   45
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   503
      BackgroundColor =   5786947
      ForeColor       =   16777215
      MouseOverBackgroundColor=   8421504
      Caption         =   "Change IP"
      BorderAnimation =   0
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
   Begin Project1.uButton uMenu 
      Height          =   290
      Index           =   3
      Left            =   4010
      TabIndex        =   51
      Top             =   50
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   503
      BackgroundColor =   5786947
      ForeColor       =   5786947
      MouseOverBackgroundColor=   8421504
      Caption         =   "Microsoft Pen"
      BorderAnimation =   0
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
   Begin Project1.uButton uMenu 
      Height          =   285
      Index           =   4
      Left            =   5325
      TabIndex        =   67
      Top             =   45
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   503
      BackgroundColor =   5786947
      ForeColor       =   16777215
      MouseOverBackgroundColor=   8421504
      Caption         =   "Scores"
      BorderAnimation =   0
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
   Begin SHDocVwCtl.WebBrowser wb1 
      Height          =   2710
      Left            =   310
      TabIndex        =   68
      Top             =   590
      Width           =   5710
      ExtentX         =   10072
      ExtentY         =   4780
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Line LineBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   386
      X2              =   1
      Y1              =   21
      Y2              =   21
   End
   Begin VB.Line LineBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   396
      X2              =   11
      Y1              =   279
      Y2              =   279
   End
   Begin VB.Line LineBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   1
      X2              =   1
      Y1              =   1
      Y2              =   114
   End
   Begin VB.Line LineBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   386
      X2              =   1
      Y1              =   1
      Y2              =   1
   End
   Begin VB.Line LineBorder 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   386
      X2              =   386
      Y1              =   19
      Y2              =   132
   End
End
Attribute VB_Name = "frmProgramTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'use the cls from this website: http://vb.mvps.org/samples/SysInfo/
'the cls has a ton of functions to prevent standby, sleep, usb removal ect.

Dim objWMIService As Object

Dim PreventNext As Boolean
Dim DataBuffer(0 To 3, 0 To 100) As Byte
Dim DataLength(0 To 3) As Long

Dim SocketState(0 To 15) As String
Dim DataSend(0 To 3) As Boolean


Dim PrevMouse As POINTAPI
Dim MouseInit As Boolean

Private Type networkSettings
    nIndex As Long
    nDHCP As Boolean
    nEnabled As Boolean
    nSubnet As String
    nIP As String
    nGateway As String
    nDescription As String
    nMAC As String
    nCardEnabled As Boolean
    nStatus As Long
End Type

Dim networkStatusString() As String
Private Const networkStatusSum As String = "Disconnected,Connecting,Connected,Disconnecting,Hardware Not Present,Hardware Disabled,Hardware Malfunction,Media Disconnected,Authenticating,Authentication Succeeded,Authentication Failed,Invalid address,Credentials Required,Other"

Dim networkStatusColorBack() As Long
Private Const networkStatusColorBackSum As String = "&H0,&H0080C0FF&,&H0080FF80&,&H0080C0FF&,&H008080FF&,&H00808080&,&H00FF80FF&,&H00FFFF80&,&hffffff,&hffffff,&hffffff,&H008080FF&,&hffffff,&hffffff"

Dim networkStatusColorFore() As Long
Private Const networkStatusColorforeSum As String = "&hffffff,&h0,&h0,&h0,&h0,&h0,&h0,&h0,&h0,&h0,&h0,&h0,&h0,&h0"



Dim NetworkAdapters() As networkSettings


Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long








Dim strPage As String
Dim LoggedIn As Boolean
Dim ScorePage As String

Private Const UserNameField As String = "ctl00$ContentPlaceHolder1$UsernameTextBox"
Private Const PasswordField As String = "ctl00$ContentPlaceHolder1$PasswordTextBox"
Private Const SubmitButton As String = "ctl00$ContentPlaceHolder1$SubmitButton"

Private Type Score
    VakCode As String
    VakNaam As String
    VakType As String
    Weging As String
    Datum As String
    Cijfer As String
    StudiePunten As String
    HonorPunten As String
    Categorie As String
End Type

Dim Scores() As Score

Private Type lProcess
    Name As String
    PercentProcessorTime As Double
    IDProcess As Long
    
End Type

Dim Processes() As lProcess


Private Sub cmdConnect_Click(Button As Integer, x As Single, y As Single)
    Dim i As Long

    If uMaster.Value = u_Checked Then


        For i = 0 To socket.Count - 1
            If socket(i).State <> sckClosed Then
                socket(i).Close
            End If

            DoEvents

            socket(i).Bind 1250 + i

            socket(i).Listen
        Next i


    ElseIf uSlave.Value = u_Checked Then


        For i = 0 To socket.Count - 1
            If socket(i).State <> sckClosed Then
                socket(i).Close
            End If
            DoEvents

            socket(i).Connect txtMasterIP.Text, 1250 + i
            DoEvents
        Next i

        'tmrMouse.Interval = 3
        'tmrMouse.Enabled = True

    End If

End Sub

Private Sub cmdRefreshScore_Click(Button As Integer, x As Single, y As Single)

    loadScores.Visible = True
    loadScores.Loading = True
    If LoggedIn Then
        wb1.Navigate "https://student.osiris.hhs.nl/osiris_student_hhsprd/Dossier.do"
    Else
        wb1.Navigate "https://student.osiris.hhs.nl/"
    End If

    wb1.Silent = True

End Sub

Private Sub cmdSort_Click(Index As Integer, Button As Integer, x As Single, y As Single)
    
    SortBy Index
End Sub

Private Sub Form_Load()

    SocketState(0) = "Closed"
    SocketState(1) = "Open"
    SocketState(2) = "Listening"
    SocketState(3) = "Connection pending"
    SocketState(4) = "Resolving host"
    SocketState(5) = "Host resolved"
    SocketState(6) = "Connecting"
    SocketState(7) = "Connected"
    SocketState(8) = "Connection closing..."
    SocketState(9) = "Error!"

    uDropBaud.AddItem 4800, 4800
    uDropBaud.AddItem 9600, 9600
    uDropBaud.AddItem 14400, 14400
    uDropBaud.AddItem 19200, 19200
    uDropBaud.AddItem 28800, 28800
    uDropBaud.AddItem 38400, 38400
    uDropBaud.AddItem 56000, 56000
    uDropBaud.AddItem 57600, 57600
    uDropBaud.AddItem 115200, 115200
    uDropBaud.ListIndex = 1
    
    
    Dim i As Long
    txtCom.Text = ""
    For i = 0 To 255
        txtCom.Text = txtCom.Text & Fmat(Hex(i), 2) & " "
    Next i
    
    lstIP(0).Font.Size = 6
    lstIP(0).setTabStop 0, 3
    lstIP(0).setTabStop 1, 100
    lstIP(0).setTabStop 2, 200
    lstIP(0).setTabStop 3, 300
    lstIP(0).setTabStop 4, 400

'    lstIP(0).AddItem "10.1.99." & vbTab & "255.255.255.0" & vbTab & ""
'    lstIP(0).AddItem "10.1.0." & vbTab & "255.255.255.0" & vbTab & ""
'    lstIP(0).AddItem "10.1.1." & vbTab & "255.255.255.0" & vbTab & ""
'    lstIP(0).AddItem "10.1.2." & vbTab & "255.255.255.0" & vbTab & ""
'    lstIP(0).AddItem "10.255.0." & vbTab & "255.255.255.248" & vbTab & ""
'
'    lstIP(1).AddItem "10.0.99." & vbTab & "255.255.255.0" & vbTab & ""
'    lstIP(1).AddItem "10.0.0." & vbTab & "255.255.255.0" & vbTab & ""
'    lstIP(1).AddItem "10.0.1." & vbTab & "255.255.255.0" & vbTab & ""
'    lstIP(1).AddItem "10.0.2." & vbTab & "255.255.255.0" & vbTab & ""
'    lstIP(1).AddItem "10.0.3." & vbTab & "255.255.255.0" & vbTab & ""
    
    lstIP(0).AddItem "192.168.0.6" & vbTab & "255.255.255.0" & vbTab & "192.168.0.1" & vbTab & "8.8.8.8" & vbTab & "8.8.4.4"
    lstIP(0).AddItem "192.168.0.22" & vbTab & "255.255.255.0" & vbTab & "192.168.0.1" & vbTab & "8.8.8.8" & vbTab & "8.8.4.4"
    lstIP(0).AddItem "172.16.0.10" & vbTab & "255.255.255.0" & vbTab & "172.16.0.1" & vbTab & "8.8.8.8" & vbTab & "8.8.4.4"
    
    'For i = 0 To 1
        lstIP(0).ListIndex = 0
    'Next i
    
    lstAdapter.Font.Size = 6

    lstAdapter.setTabStop 0, 20, vbCenter
    lstAdapter.setTabStop 1, 40
    lstAdapter.setTabStop 2, lstAdapter.Width - 20, vbRightJustify


    uMenu_Click 1, 0, 0, 0

    tmrTopMost.Enabled = True

    networkStatusString = Split(networkStatusSum, ",")
    
    Dim tmpstr() As String
    tmpstr = Split(networkStatusColorforeSum, ",")
    ReDim networkStatusColorFore(0 To UBound(tmpstr))
    For i = 0 To UBound(tmpstr)
        networkStatusColorFore(i) = Val(tmpstr(i))
    Next i
    
    tmpstr = Split(networkStatusColorBackSum, ",")
    ReDim networkStatusColorBack(0 To UBound(tmpstr))
    For i = 0 To UBound(tmpstr)
        networkStatusColorBack(i) = Val(tmpstr(i))
    Next i
    
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    
    
    RefreshNetworkAdapters
End Sub

Sub PenButtonPress_Single()
    If uPenOption(0).Value = u_Checked Then    'program
        ShellExecute Me.hWnd, "open", txtPenProgram(0).Text, "", "", vbNormalFocus
    Else    'keypress
        VbSendKeys txtPenKeyStroke(0).Text
    End If

End Sub

Sub PenButtonPress_Double()
    If uPenOption(0).Value = u_Checked Then    'program
        ShellExecute Me.hWnd, "open", txtPenProgram(1).Text, "", "", vbNormalFocus
    Else    'keypress
        VbSendKeys txtPenKeyStroke(1).Text
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    RemoveKeyboardHook
End Sub



Sub RefreshNetworkAdapters()
    Dim colNetAdapters, objNetAdapter
    Dim i As Long
    
    Static Refreshing As Boolean
    
    If Refreshing = True Then Exit Sub
    Refreshing = True
    

    Set colNetAdapters = objWMIService.ExecQuery _
                         ("Select * from Win32_NetworkAdapterConfiguration ")  '& "where IPEnabled=TRUE"

    Dim previousSelected As Long
    previousSelected = lstAdapter.ListIndex
    
    lstAdapter.RedrawPause
    lstAdapter.Clear


    DoEvents
    Debug.Print colNetAdapters.Count
    If colNetAdapters.Count > 0 Then
        ReDim NetworkAdapters(0 To 300) As networkSettings
    Else
        Exit Sub
    End If


    For Each objNetAdapter In colNetAdapters
        i = CLng(objNetAdapter.Index)
        
        NetworkAdapters(i).nDHCP = CBool(objNetAdapter.DHCPEnabled)
        NetworkAdapters(i).nDescription = CStr(objNetAdapter.Description)
        NetworkAdapters(i).nIP = CStr(objNetAdapter.IPAddress(0))
        NetworkAdapters(i).nSubnet = CStr(objNetAdapter.IPSubnet(0))
        NetworkAdapters(i).nEnabled = CBool(objNetAdapter.IPEnabled)    'Or CBool(objNetAdapter.IPXEnabled)
        'MsgBox TypeName(objNetAdapter.MACAddress)
        If TypeName(objNetAdapter.SettingID) = "Null" Then
            NetworkAdapters(i).nMAC = ""
        Else
            NetworkAdapters(i).nMAC = CStr(objNetAdapter.SettingID)
        End If

        If TypeName(objNetAdapter.DefaultIPGateway) = "Variant()" Then
            NetworkAdapters(i).nGateway = CStr(objNetAdapter.DefaultIPGateway(0))
        Else
            NetworkAdapters(i).nGateway = ""
        End If

        DoEvents
        'lol = lol & textstst(objNetAdapter)



    Next


    Set colNetAdapters = objWMIService.ExecQuery _
                         ("Select * from Win32_NetworkAdapter where NetConnectionStatus >= 0")  '& "where IPEnabled=TRUE"
    DoEvents
    
    For Each objNetAdapter In colNetAdapters
        i = CLng(objNetAdapter.Index)
        If TypeName(objNetAdapter.NetConnectionStatus) = "Null" Then
            NetworkAdapters(i).nCardEnabled = False
            NetworkAdapters(i).nStatus = -1
        Else
            'MsgBox objNetAdapter.NetConnectionStatus
            NetworkAdapters(i).nCardEnabled = CBool(objNetAdapter.NetEnabled)
            NetworkAdapters(i).nStatus = objNetAdapter.NetConnectionStatus
        End If
        'Debug.Print i & " NetConnectionStatus: " & objNetAdapter.NetConnectionStatus

        'If TypeName(objNetAdapter.NetConnectionStatus) <> "Null" Then
        'If CLng(objNetAdapter.NetConnectionStatus) = 2 Then objNetAdapter.Disable

        'End If

        lstAdapter.AddItem i & vbTab & NetworkAdapters(i).nDescription & vbTab & networkStatusString(NetworkAdapters(i).nStatus), i, , networkStatusColorBack(NetworkAdapters(i).nStatus), networkStatusColorFore(NetworkAdapters(i).nStatus)
        DoEvents
    Next

    If lstAdapter.ListCount >= 5 Then
        lstAdapter.ItemsVisible = 5
    Else
        lstAdapter.ItemsVisible = lstAdapter.ListCount
    End If


    'For i = 0 To colNetAdapters.Count - 1
    '    lstAdapter.AddItem i & vbTab & NetworkAdapters(i).nDescription, , , IIf(NetworkAdapters(i).nCardEnabled, vbGreen, vbRed)
    'Next i

    'Clipboard.Clear
    'Clipboard.SetText lol

    If previousSelected = -1 Then
        If lstAdapter.ListCount > 0 Then
            previousSelected = 0
        End If
    End If

    lstAdapter.ListIndex = previousSelected
    
    lstAdapter.RedrawResume
    
    
    Refreshing = False
End Sub

Function textstst(objItem As Variant) As String
    Dim tmpstr As String
    tmpstr = ""
    
    tmpstr = tmpstr & "Description                 : " & vbTab & objItem.Description & vbCrLf
    tmpstr = tmpstr & "IPAddress                   : " & vbTab & GetMultiString_FromArray(objItem.IPAddress, ", ") & vbCrLf
    tmpstr = tmpstr & "IPSubnet                    : " & vbTab & GetMultiString_FromArray(objItem.IPSubnet, ", ") & vbCrLf
    tmpstr = tmpstr & "DefaultIPGateway            : " & vbTab & GetMultiString_FromArray(objItem.DefaultIPGateway, ", ") & vbCrLf & vbCrLf
    tmpstr = tmpstr & "DNSServerSearchOrder        : " & vbTab & GetMultiString_FromArray(objItem.DNSServerSearchOrder, ", ") & vbCrLf & vbCrLf & vbCrLf
    
    
    tmpstr = tmpstr & "ArpAlwaysSourceRoute        : " & vbTab & objItem.ArpAlwaysSourceRoute & vbCrLf
    tmpstr = tmpstr & "ArpUseEtherSNAP             : " & vbTab & objItem.ArpUseEtherSNAP & vbCrLf
    tmpstr = tmpstr & "DHCPEnabled                 : " & vbTab & objItem.DHCPEnabled & vbCrLf
    tmpstr = tmpstr & "DHCPLeaseExpires            : " & vbTab & ConvertFromWMIDateTime(objItem.DHCPLeaseExpires) & vbCrLf
    tmpstr = tmpstr & "DHCPLeaseObtained           : " & vbTab & ConvertFromWMIDateTime(objItem.DHCPLeaseObtained) & vbCrLf
    tmpstr = tmpstr & "DHCPServer                  : " & vbTab & objItem.DHCPServer & vbCrLf
    tmpstr = tmpstr & "DNSDomain                   : " & vbTab & objItem.DNSDomain & vbCrLf
    tmpstr = tmpstr & "DNSDomainSuffixSearchOrder  : " & vbTab & GetMultiString_FromArray(objItem.DNSDomainSuffixSearchOrder, ", ") & vbCrLf
    
    tmpstr = tmpstr & "DNSEnabledForWINSResolution : " & vbTab & objItem.DNSEnabledForWINSResolution & vbCrLf
    tmpstr = tmpstr & "DNSHostName                 : " & vbTab & objItem.DNSHostName & vbCrLf
    
    tmpstr = tmpstr & "DatabasePath                : " & vbTab & objItem.DatabasePath & vbCrLf
    tmpstr = tmpstr & "DeadGWDetectEnabled         : " & vbTab & objItem.DeadGWDetectEnabled & vbCrLf
    
    tmpstr = tmpstr & "DefaultTOS                  : " & vbTab & objItem.DefaultTOS & vbCrLf
    tmpstr = tmpstr & "DefaultTTL                  : " & vbTab & objItem.DefaultTTL & vbCrLf
    
    tmpstr = tmpstr & "DomainDNSRegistrationEnabled: " & vbTab & objItem.DomainDNSRegistrationEnabled & vbCrLf
    tmpstr = tmpstr & "ForwardBufferMemory         : " & vbTab & objItem.ForwardBufferMemory & vbCrLf
    tmpstr = tmpstr & "FullDNSRegistrationEnabled  : " & vbTab & objItem.FullDNSRegistrationEnabled & vbCrLf
    tmpstr = tmpstr & "GatewayCostMetric           : " & vbTab & GetMultiString_FromArray(objItem.GatewayCostMetric, ", ") & vbCrLf
    tmpstr = tmpstr & "IGMPLevel                   : " & vbTab & objItem.IGMPLevel & vbCrLf
    
    tmpstr = tmpstr & "IPConnectionMetric          : " & vbTab & objItem.IPConnectionMetric & vbCrLf
    tmpstr = tmpstr & "IPEnabled                   : " & vbTab & objItem.IPEnabled & vbCrLf
    tmpstr = tmpstr & "IPFilterSecurityEnabled     : " & vbTab & objItem.IPFilterSecurityEnabled & vbCrLf
    tmpstr = tmpstr & "IPPortSecurityEnabled       : " & vbTab & objItem.IPPortSecurityEnabled & vbCrLf
    tmpstr = tmpstr & "IPSecPermitIPProtocols      : " & vbTab & GetMultiString_FromArray(objItem.IPSecPermitIPProtocols, ", ") & vbCrLf
    tmpstr = tmpstr & "IPSecPermitTCPPorts         : " & vbTab & GetMultiString_FromArray(objItem.IPSecPermitTCPPorts, ", ") & vbCrLf
    tmpstr = tmpstr & "IPSecPermitUDPPorts         : " & vbTab & GetMultiString_FromArray(objItem.IPSecPermitUDPPorts, ", ") & vbCrLf
    
    tmpstr = tmpstr & "IPUseZeroBroadcast          : " & vbTab & objItem.IPUseZeroBroadcast & vbCrLf
    tmpstr = tmpstr & "IPXAddress                  : " & vbTab & objItem.IPXAddress & vbCrLf
    tmpstr = tmpstr & "IPXEnabled                  : " & vbTab & objItem.IPXEnabled & vbCrLf
    tmpstr = tmpstr & "IPXFrameType                : " & vbTab & GetMultiString_FromArray(objItem.IPXFrameType, ", ") & vbCrLf
    tmpstr = tmpstr & "IPXNetworkNumber            : " & vbTab & GetMultiString_FromArray(objItem.IPXNetworkNumber, ", ") & vbCrLf
    tmpstr = tmpstr & "IPXVirtualNetNumber         : " & vbTab & objItem.IPXVirtualNetNumber & vbCrLf
    tmpstr = tmpstr & "Index                       : " & vbTab & objItem.Index & vbCrLf
    tmpstr = tmpstr & "KeepAliveInterval           : " & vbTab & objItem.KeepAliveInterval & vbCrLf
    tmpstr = tmpstr & "KeepAliveTime               : " & vbTab & objItem.KeepAliveTime & vbCrLf
    tmpstr = tmpstr & "MACAddress                  : " & vbTab & objItem.MACAddress & vbCrLf
    tmpstr = tmpstr & "MTU                         : " & vbTab & objItem.MTU & vbCrLf
    tmpstr = tmpstr & "NumForwardPackets           : " & vbTab & objItem.NumForwardPackets & vbCrLf
    tmpstr = tmpstr & "PMTUBHDetectEnabled         : " & vbTab & objItem.PMTUBHDetectEnabled & vbCrLf
    tmpstr = tmpstr & "PMTUDiscoveryEnabled        : " & vbTab & objItem.PMTUDiscoveryEnabled & vbCrLf
    tmpstr = tmpstr & "ServiceName                 : " & vbTab & objItem.ServiceName & vbCrLf
    tmpstr = tmpstr & "SettingID                   : " & vbTab & objItem.SettingID & vbCrLf
    tmpstr = tmpstr & "TcpMaxConnectRetransmissions: " & vbTab & objItem.TcpMaxConnectRetransmissions & vbCrLf
    tmpstr = tmpstr & "TcpMaxDataRetransmissions   : " & vbTab & objItem.TcpMaxDataRetransmissions & vbCrLf
    tmpstr = tmpstr & "TcpNumConnections           : " & vbTab & objItem.TcpNumConnections & vbCrLf
    tmpstr = tmpstr & "TcpUseRFC1122UrgentPointer  : " & vbTab & objItem.TcpUseRFC1122UrgentPointer & vbCrLf
    tmpstr = tmpstr & "TcpWindowSize               : " & vbTab & objItem.TcpWindowSize & vbCrLf
    tmpstr = tmpstr & "TcpipNetbiosOptions         : " & vbTab & objItem.TcpipNetbiosOptions & vbCrLf
    tmpstr = tmpstr & "WINSEnableLMHostsLookup     : " & vbTab & objItem.WINSEnableLMHostsLookup & vbCrLf
    tmpstr = tmpstr & "WINSHostLookupFile          : " & vbTab & objItem.WINSHostLookupFile & vbCrLf
    tmpstr = tmpstr & "WINSPrimaryServer           : " & vbTab & objItem.WINSPrimaryServer & vbCrLf
    tmpstr = tmpstr & "WINSScopeID                 : " & vbTab & objItem.WINSScopeID & vbCrLf
    tmpstr = tmpstr & "WINSSecondaryServer         : " & vbTab & objItem.WINSSecondaryServer & vbCrLf

    tmpstr = tmpstr & vbCrLf & vbCrLf


    textstst = tmpstr

End Function

Function ConvertFromWMIDateTime(dDateTime)
    On Error Resume Next
    Dim oDateTime
    Set oDateTime = CreateObject("WbemScripting.SWbemDateTime")
    oDateTime.Value = dDateTime
    ConvertFromWMIDateTime = oDateTime.GetVarDate
    Set oDateTime = Nothing
End Function

Function GetMultiString_FromArray(ArrayString, Seprator)
    Dim StrMultiArray
    If IsNull(ArrayString) Then
        StrMultiArray = ArrayString
    Else
        StrMultiArray = Join(ArrayString, Seprator)
    End If
    GetMultiString_FromArray = StrMultiArray
End Function


Sub setNetworkCardState(EnableCard As Boolean)
    Dim colNetAdapters, objNetAdapter
    Dim strIPAddress, strSubnetMask, strGateway, strGatewaymetric
    Dim errEnable, errGateways
    Dim i As Long

    i = lstAdapter.ListIndex
    If i = -1 Then Exit Sub

    i = lstAdapter.ItemData(i)

    Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter " & "where Index=" & i)

    

    For Each objNetAdapter In colNetAdapters
        If EnableCard Then
            objNetAdapter.Enable
        Else
            objNetAdapter.Disable
        End If
    Next

End Sub



Sub ToggleNetworkType(Automatic As Boolean)
    Dim colNetAdapters, objNetAdapter
    Dim strIPAddress, strSubnetMask, strGateway, strGatewaymetric
    Dim errEnable, errGateways

    If Automatic = False And lstAdapter.ListIndex = -1 Then Exit Sub


    If Automatic = False Then
        SetStaticIp
    Else

        Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration " & "where Index=" & lstAdapter.ItemData(lstAdapter.ListIndex))
        For Each objNetAdapter In colNetAdapters
            objNetAdapter.EnableDHCP
        Next
    End If



End Sub


Private Sub lstAdapter_ItemChange(ItemIndex As Long)
    Dim i As Long
    i = lstAdapter.ListIndex
    If i = -1 Then Exit Sub

    i = lstAdapter.ItemData(i)

    If NetworkAdapters(i).nDHCP Then
        lblIPSettings(4).Caption = "Type:" & vbCrLf & "Automatic "
        uChangeIP.Caption = "Set Static"
    Else
        lblIPSettings(4).Caption = "Type:" & vbCrLf & "Static "
        uChangeIP.Caption = "Set Automatic"
    End If

    lblIPSettings(1).Caption = "IP:" & vbCrLf & NetworkAdapters(i).nIP

    lblIPSettings(2).Caption = "SubnetMask:" & vbCrLf & NetworkAdapters(i).nSubnet

    lblIPSettings(3).Caption = "DefaultGateway:" & vbCrLf & NetworkAdapters(i).nGateway
End Sub


Function Fmat(str As String, Length As Long) As String
    Dim strLength As Long
    strLength = Len(str)

    If strLength > Length Then
        Fmat = String(Length, "X")
    ElseIf strLength < Length Then
        Fmat = String(Length - strLength, "0") & str
    Else
        Fmat = str
    End If

End Function



Sub Form_Resize()
    Dim i As Long

    Me.Height = 3705
    Me.Width = 7060 'frmMain.Width

    LineBorder(1).X1 = 1
    LineBorder(1).Y1 = 1
    LineBorder(1).X2 = Me.ScaleWidth - 1
    LineBorder(1).Y2 = 1

    LineBorder(2).X1 = 1
    LineBorder(2).Y1 = 1
    LineBorder(2).X2 = 1
    LineBorder(2).Y2 = Me.ScaleHeight - 1

    LineBorder(0).X1 = Me.ScaleWidth - 2
    LineBorder(0).Y1 = 1
    LineBorder(0).X2 = Me.ScaleWidth - 2
    LineBorder(0).Y2 = Me.ScaleHeight - 1

    LineBorder(3).X1 = 1
    LineBorder(3).Y1 = Me.ScaleHeight - 2
    LineBorder(3).X2 = Me.ScaleWidth - 1
    LineBorder(3).Y2 = Me.ScaleHeight - 2


    LineBorder(4).X1 = 1
    LineBorder(4).Y1 = 23 + uMenu(1).Top - 1
    LineBorder(4).X2 = Me.ScaleWidth - 1
    LineBorder(4).Y2 = 23 + uMenu(1).Top - 1


    For i = 0 To picTab.Count - 1
        On Error Resume Next
        picTab(i).Left = 2
        picTab(i).Width = Me.ScaleWidth - 4
        picTab(i).Top = (23 + uMenu(1).Top)
        picTab(i).Height = Me.ScaleHeight - 3 - (23 + uMenu(1).Top)
    Next i
    
    loadScores.Left = 0
    loadScores.Top = 0
    loadScores.Width = picTab(4).Width
    loadScores.Height = picTab(4).Height
    
    uClose.Top = 3
    uClose.Left = Me.ScaleWidth - 3 - uClose.Width


End Sub

Private Sub hSpeed_Change()
    lblSpeed.Caption = "Speed: " & (hSpeed.Value / 100) & "x"
End Sub





Private Sub lstIP_ItemChange(Index As Integer, ItemIndex As Long)
    If ItemIndex = -1 Then
        Exit Sub
    End If

    Dim tmpSplit() As String
    tmpSplit = Split(lstIP(Index).List(ItemIndex), vbTab)
    
    Dim i As Long
    
    For i = 0 To UBound(tmpSplit)
        txtSetIP(i).Text = tmpSplit(i)
    Next i
    
End Sub

Private Sub scrollScores_Change()
    DrawScores
End Sub

Private Sub scrollScores_Scroll()
    scrollScores_Change
End Sub

Private Sub socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If socket(Index).State <> sckClosed Then
        socket(Index).Close
    End If


    socket(Index).Accept requestID
End Sub


Private Sub socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim bytes() As Byte
    Dim m_CursorPos As POINTAPI
    Dim i As Long

    socket(Index).GetData bytes, vbByte

    'GetAsyncKeyState (vbKeyControl)

    GetCursorPos m_CursorPos

    For i = 0 To UBound(bytes)
        'DataBuffer(Index, DataLength + i) = bytes(i)

        Select Case Index
            Case 0
                m_CursorPos.x = m_CursorPos.x + (bytes(i) - 128)
                SetCursorPos m_CursorPos.x, m_CursorPos.y

            Case 1
                m_CursorPos.y = m_CursorPos.y + (bytes(i) - 128)
                SetCursorPos m_CursorPos.x, m_CursorPos.y

            Case 2
                'to build yet

        End Select
        DoEvents

    Next i

    'DataLength(Index) = DataLength(Index) + UBound(bytes)
End Sub


Private Sub socket_SendComplete(Index As Integer)
    DataSend(Index) = False
End Sub

Private Sub tmrAdapterRefresh_Timer()
    If Me.Visible = True And picTab(2).Visible = True Then
        RefreshNetworkAdapters
    End If

End Sub

Private Sub tmrConnection_Timer()
    Dim i As Long
    Dim L As Long
    Dim j As Long

    For i = 0 To socket.Count - 1
        lblState(i).Caption = "Socket " & (i + 1) & ": " & SocketState(socket(i).State)
        If socket(i).State = 7 Then L = L + 1
        If socket(i).State = 8 Then j = j + 1
    Next i

    If L = socket.Count And tmrMouse.Enabled = False And uMaster.Value = u_Checked Then
        tmrMouse.Interval = 30
        tmrMouse.Enabled = True
    ElseIf j = socket.Count Then
        tmrMouse.Enabled = False

        For i = 0 To socket.Count - 1
            socket(i).Close
        Next i
    End If
End Sub

Private Sub tmrMouse_Timer()
    Dim currchar As String
    Dim Axis As Long

    Dim strSend As String

    Dim Xaxis As Long
    Dim Yaxis As Long

    Dim m_CursorPos As POINTAPI

    Dim CenterX As Long
    Dim CenterY As Long
    Static ControlDown As Boolean
    Static WhatIndex As Long

    CenterX = Screen.Width / 2 / Screen.TwipsPerPixelX
    CenterY = Screen.Height / 2 / Screen.TwipsPerPixelY



    If GetAsyncKeyState(vbKeyControl) <> 0 Then
        ControlDown = True
        Exit Sub
    Else
        If ControlDown = True Then
            ControlDown = False
            SetCursorPos CenterX, CenterY
        End If
    End If

    GetCursorPos m_CursorPos


    If WhatIndex = 0 Then
        WhatIndex = 1

        SetCursorPos CenterX, m_CursorPos.y

        If MouseInit Then
            Xaxis = (m_CursorPos.x - PrevMouse.x) + 128
            PrevMouse.x = CenterX

            If Xaxis < 0 Then Xaxis = 0
            If Xaxis > 255 Then Xaxis = 255

            DoEvents
            If DataSend(0) = False Then
                DataSend(0) = True
                socket(0).SendData Chr(Xaxis)
            End If

        Else
            PrevMouse.x = m_CursorPos.x
            MouseInit = True
        End If
    ElseIf WhatIndex = 1 Then
        WhatIndex = 0

        SetCursorPos m_CursorPos.x, CenterY

        If MouseInit Then
            Yaxis = (m_CursorPos.y - PrevMouse.y) + 128
            PrevMouse.y = CenterY

            If Yaxis < 0 Then Yaxis = 0
            If Yaxis > 255 Then Yaxis = 255

            DoEvents
            If DataSend(1) = False Then
                DataSend(1) = True
                socket(1).SendData Chr(Yaxis)
            End If
        Else
            PrevMouse.y = m_CursorPos.y
            MouseInit = True
        End If
    End If

    '
    '    If uMouseEnabled.Value = u_UnChecked Then
    '        If socket.State <> sckClosed Then
    '            socket.Close
    '        End If
    '        tmrMouse.Enabled = False
    '        Exit Sub
    '
    '    End If
    '
    '
    '
    '    If uMaster.Value = u_Checked Then
    '        CenterX = Screen.Width / 2 / Screen.TwipsPerPixelX
    '        CenterY = Screen.Height / 2 / Screen.TwipsPerPixelY
    '
    '        If GetAsyncKeyState(vbKeyControl) <> 0 Then
    '            ControlDown = True
    '            Exit Sub
    '        Else
    '            If ControlDown = True Then
    '                ControlDown = False
    '                SetCursorPos CenterX, CenterY
    '            End If
    '        End If
    '
    '        GetCursorPos m_CursorPos
    '        SetCursorPos CenterX, CenterY
    '
    '        If MouseInit Then
    '            Xaxis = (m_CursorPos.x - PrevMouse.x) + 64 + 128
    '            Yaxis = (m_CursorPos.y - PrevMouse.y) + 64
    '
    '            PrevMouse.x = CenterX
    '            PrevMouse.y = CenterY
    '
    '            If Xaxis < 0 Or Xaxis > 255 Or _
                 '               Yaxis < 0 Or Yaxis > 127 Then
    '               Exit Sub
    '            End If
    '
    '            socket.SendData Chr(Xaxis) & Chr(Yaxis)
    '        Else
    '            PrevMouse = m_CursorPos
    '            MouseInit = True
    '        End If
    '
    '    Else
    '
    '        If Len(DataBuffer) > 1 Then
    '
    '
    '            Dim i As Long
    '
    '            For i = 0 To 1
    '
    '                currchar = Left$(DataBuffer, 1)
    '                DataBuffer = Right$(DataBuffer, Len(DataBuffer) - 1)
    '
    '                Axis = Asc(currchar)
    '                If Axis > 128 Then
    '                    Xaxis = (Axis And 127) - 64
    '                Else
    '                    Yaxis = Axis - 64
    '                End If
    '            Next i
    '
    '            GetCursorPos m_CursorPos
    '
    '            SetCursorPos m_CursorPos.x + (Xaxis * (hSpeed.Value / 100)), m_CursorPos.y + (Yaxis * (hSpeed.Value / 100))
    '
    '        End If
    '    End If

End Sub


Sub tmrTopMost_Timer()
    If Me.Visible Then
        SetTopMostWindow Me.hWnd, True
        tmrTopMost.Enabled = False
    End If
End Sub


Private Sub txtCom_Click()
    Dim lStart As Long
    Dim i As Long

    Dim lStr() As String


    lStart = txtCom.SelStart / 3
    lStart = lStart - (lStart Mod 8)

    lStr = Split(Mid$(txtCom.Text, lStart * 3 + 1, 8 * 3), " ")


    For i = 0 To UBound(lStr) - 1
        txtComSplit(i).Text = Val("&h" & lStr(i))
    Next i
End Sub


Private Sub txtPenKeyStroke_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    Dim tmpTotal As String
'
'    If (Shift And 1) Then
'        tmpTotal = tmpTotal & "+"
'    End If
'
'    If (Shift And 2) Then
'        tmpTotal = tmpTotal & "^"
'    End If
'
'    If (Shift And 4) Then
'        tmpTotal = tmpTotal & "%"
'    End If
'
'    If KeyCode >= vbKeyA And KeyCode <= vbKeyZ Then
'        tmpTotal = tmpTotal & Chr(KeyCode)
'    End If
'
'    txtPenKeyStroke(Index).Text = tmpTotal
End Sub

Private Sub txtPingResult_Click()
    txtPingResult.Visible = False
End Sub

Private Sub uChangeIP_Click(Button As Integer, x As Single, y As Single)
    Dim i As Long
    i = lstAdapter.ListIndex
    If i = -1 Then Exit Sub

    ToggleNetworkType Not NetworkAdapters(lstAdapter.ItemData(i)).nDHCP

    RefreshNetworkAdapters
End Sub

Private Sub uClose_Click(Button As Integer, x As Single, y As Single)
    Me.Hide
End Sub

Private Sub uCloseStats_Click(Button As Integer, x As Single, y As Single)
    txtStats.Visible = False
    uCloseStats.Visible = False
End Sub

Private Sub uEnableDisable_Click(Index As Integer, Button As Integer, x As Single, y As Single)
    Dim i As Long
    i = lstAdapter.ListIndex
    If i = -1 Then Exit Sub

    setNetworkCardState CBool(Index)

    RefreshNetworkAdapters
End Sub

Private Sub uMenu_Click(Index As Integer, Button As Integer, x As Single, y As Single)
    Dim i As Long


    For i = 0 To uMenu.Count - 1
        If i <> Index Then
            uMenu(i).Height = 23
            picTab(i).Visible = False
        Else
            uMenu(Index).Height = 24
            picTab(Index).Visible = True
        End If
    Next i
End Sub

Private Sub uMicrosoftPen_ActivateNextState(u_Cancel As Boolean, u_NewState As uCheckboxConstants)
    If u_NewState = u_UnChecked Then
        u_NewState = u_Checked
        SetKeyboardHook
    ElseIf u_NewState = u_Checked Then
        u_NewState = u_UnChecked
        RemoveKeyboardHook
    End If

End Sub

Private Sub uMouseEnabled_ActivateNextState(u_Cancel As Boolean, u_NewState As uCheckboxConstants)
    PreventNext = True

    uMaster.Value = u_UnChecked
    uSlave.Value = u_UnChecked

    PreventNext = False
End Sub


Private Sub uMaster_ActivateNextState(u_Cancel As Boolean, u_NewState As uCheckboxConstants)
    If PreventNext Then Exit Sub
    PreventNext = True

    If uMouseEnabled.Value = u_UnChecked Then
        u_Cancel = True
        u_NewState = u_UnChecked
        PreventNext = False
        Exit Sub
    End If

    If uMaster.Value = u_UnChecked Then
        uSlave.Value = u_UnChecked
    Else
        uSlave.Value = u_Checked
    End If

    PreventNext = False
End Sub


Private Sub uPing_Click(Button As Integer, x As Single, y As Single)
    Dim Reply As ICMP_ECHO_REPLY
    Dim lngSuccess As Long
    Dim strIPAddress As String
    Dim tmpResult As String

    txtPingResult.Visible = True

    'Get the sockets ready.
    If SocketsInitialize() Then


        'Address to ping
        strIPAddress = txtSetIP(0).Text
        If strIPAddress = "" Then
            txtPingResult.Text = "no ip!"
            GoTo cleanupshit
        End If

        txtPingResult.Text = "Sending Ping: " & strIPAddress & " ..." & vbCrLf
        DoEvents

        'Ping the IP that is passing the address and get a reply.
        lngSuccess = ping(strIPAddress, Reply)

        txtPingResult.Text = txtPingResult.Text & "ICMP code   : " & lngSuccess & vbCrLf
        txtPingResult.Text = txtPingResult.Text & "Message     : " & EvaluatePingResponse(lngSuccess) & vbCrLf
        txtPingResult.Text = txtPingResult.Text & "Time        : " & Reply.RoundTripTime & " ms" & vbCrLf
cleanupshit:
        SocketsCleanup
    Else

        'Winsock error failure, initializing the sockets.
        txtPingResult.Text = "Error while creating WinSock Sockets"

    End If
    DoEvents
End Sub

Private Sub uRefreshIp_Click(Button As Integer, x As Single, y As Single)
    RefreshNetworkAdapters
End Sub

Private Sub uRefreshProcess_Click(Button As Integer, x As Single, y As Single)
    Dim colProcesses, process
    
    Static prevProcessTime As Double
    Dim totalProcessTime As Double
    Dim idleTime As Double
    Static prevIdleTime As Double
    
    uProcess.RedrawPause
    uProcess.Clear
    
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfProc_Process") 'Win32_PerfFormattedData_PerfProc_Process
    
    ReDim Processes(0 To colProcesses.Count - 1)
    
    Dim i As Long
    
    For Each process In colProcesses
        
        With Processes(i)
            .IDProcess = process.IDProcess
            .Name = process.Name
            .PercentProcessorTime = process.PercentUserTime
            
            i = i + 1
        End With
        
    Next
    
    Dim p() As lProcess
    MergeSort1 Processes, p
    
    Debug.Print idleTime - prevIdleTime
    
    prevIdleTime = idleTime
    
    For i = 0 To UBound(Processes)
        uProcess.AddItem Processes(i).Name & vbTab & Processes(i).PercentProcessorTime
        
    Next i
    
    '
    '
    '
    '
    '
    'ProcessID
    'ThreadCount
    
    uProcess.RedrawResume
    
End Sub



Private Sub MergeSort1(ByRef pvarArray() As lProcess, pvarMirror() As lProcess, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngMid As Long
    Dim L As Long
    Dim R As Long
    Dim O As Long
    Dim varSwap As lProcess
 
    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
        ReDim pvarMirror(plngLeft To plngRight)
    End If
    lngMid = plngRight - plngLeft
    Select Case lngMid
        Case 0
        Case 1
            If pvarArray(plngLeft).PercentProcessorTime < pvarArray(plngRight).PercentProcessorTime Then
                varSwap = pvarArray(plngLeft)
                pvarArray(plngLeft) = pvarArray(plngRight)
                pvarArray(plngRight) = varSwap
            End If
        Case Else
            lngMid = lngMid \ 2 + plngLeft
            MergeSort1 pvarArray, pvarMirror, plngLeft, lngMid
            MergeSort1 pvarArray, pvarMirror, lngMid + 1, plngRight
            ' Merge the resulting halves
            L = plngLeft ' start of first (left) half
            R = lngMid + 1 ' start of second (right) half
            O = plngLeft ' start of output (mirror array)
            Do
                If pvarArray(R).PercentProcessorTime > pvarArray(L).PercentProcessorTime Then
                    pvarMirror(O) = pvarArray(R)
                    R = R + 1
                    If R > plngRight Then
                        For L = L To lngMid
                            O = O + 1
                            pvarMirror(O) = pvarArray(L)
                        Next
                        Exit Do
                    End If
                Else
                    pvarMirror(O) = pvarArray(L)
                    L = L + 1
                    If L > lngMid Then
                        For R = R To plngRight
                            O = O + 1
                            pvarMirror(O) = pvarArray(R)
                        Next
                        Exit Do
                    End If
                End If
                O = O + 1
            Loop
            For O = plngLeft To plngRight
                pvarArray(O) = pvarMirror(O)
            Next
    End Select
End Sub



Private Sub uRenew_Click(Button As Integer, x As Single, y As Single)
    Dim colNetAdapters, objNetAdapter
    Dim strIPAddress, strSubnetMask, strGateway, strGatewaymetric
    Dim errEnable, errGateways
    Dim i As Long

    i = lstAdapter.ListIndex
    If i = -1 Then Exit Sub

    i = lstAdapter.ItemData(i)

    Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration " & "where Index=" & i)

    For Each objNetAdapter In colNetAdapters
        objNetAdapter.ReleaseDHCPLease
        DoEvents
        objNetAdapter.RenewDHCPLease
    Next
End Sub

Sub setNetworkDNS()
    Dim colNetAdapters, objNetAdapter
    Dim strDNS
    Dim errDNS As Long, errGateways

    If lstAdapter.ListIndex = -1 Then Exit Sub

    Set colNetAdapters = objWMIService.ExecQuery _
                         ("Select * from Win32_NetworkAdapterConfiguration " & "where Index=" & lstAdapter.ItemData(lstAdapter.ListIndex))

    strDNS = Array(txtSetIP(3).Text, txtSetIP(4).Text)

    For Each objNetAdapter In colNetAdapters
        errDNS = objNetAdapter.SetDNSServerSearchOrder(strDNS)
    Next
    
    txtPingResult.Text = "DNS:     " & GetSetIpError(errDNS)

    txtPingResult.Visible = True
End Sub

Private Sub uSetDNS_Click(Button As Integer, x As Single, y As Single)
    setNetworkDNS
    
    RefreshNetworkAdapters
End Sub

Private Sub uSetIp_Click(Button As Integer, x As Single, y As Single)
    SetStaticIp
    
    RefreshNetworkAdapters
End Sub

Sub SetStaticIp()
    Dim colNetAdapters, objNetAdapter
    Dim strIPAddress, strSubnetMask, strGateway, strGatewaymetric
    Dim errEnable As Long, errGateways As Long

    If lstAdapter.ListIndex = -1 Then Exit Sub

    If txtSetIP(0).Text = "" Or txtSetIP(1).Text = "" Then Exit Sub

    Set colNetAdapters = objWMIService.ExecQuery _
                         ("Select * from Win32_NetworkAdapterConfiguration " & "where Index=" & lstAdapter.ItemData(lstAdapter.ListIndex))

    strIPAddress = Array(txtSetIP(0).Text)
    strSubnetMask = Array(txtSetIP(1).Text)
    If txtSetIP(2).Text = "" Then
        strGateway = Array(txtSetIP(0).Text)
    Else
        strGateway = Array(txtSetIP(2).Text)
    End If

    strGatewaymetric = Array(1)

    For Each objNetAdapter In colNetAdapters
        errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask)
        errGateways = objNetAdapter.SetGateways(strGateway, strGatewaymetric)
    Next
    
    
    txtPingResult.Text = "IP:      " & GetSetIpError(errEnable) & vbCrLf & _
                         "Gateway: " & GetSetIpError(errGateways) & vbCrLf

    txtPingResult.Visible = True
End Sub

Private Function GetSetIpError(errnum As Long) As String
    Dim msg As String

    Select Case errnum

        Case 0:
            msg = "Successful completion, no reboot required"
    
        Case 1:
            msg = "Successful completion, reboot required"
    
        Case 64:
            msg = "Method not supported on this platform"
    
        Case 65:
            msg = "Unknown failure"
    
        Case 66:
            msg = "Invalid subnet mask"
    
        Case 67:
            msg = "An error occurred while processing an Instance that was returned"
    
        Case 68:
            msg = "Invalid input parameter"
    
        Case 69:
            msg = "More than 5 gateways specified"
    
        Case 70:
            msg = "Invalid IP  address"
    
        Case 71:
            msg = "Invalid gateway IP address"
    
        Case 72:
            msg = "An error occurred while accessing the Registry for the requested information"
    
        Case 73:
            msg = "Invalid domain name"
    
        Case 74:
            msg = "Invalid host name"
    
        Case 75:
            msg = "No primary/secondary WINS server defined"
    
        Case 76:
            msg = "Invalid file"
    
        Case 77:
            msg = "Invalid system path"
    
        Case 78:
            msg = "File copy failed"
    
        Case 79:
            msg = "Invalid security parameter"
    
        Case 80:
            msg = "Unable to configure TCP/IP service"
    
        Case 81:
            msg = "Unable to configure DHCP service"
    
        Case 82:
            msg = "Unable to renew DHCP lease"
    
        Case 83:
            msg = "Unable to release DHCP lease"
    
        Case 84:
            msg = "IP not enabled on adapter"
    
        Case 85:
            msg = "IPX not enabled on adapter"
    
        Case 86:
            msg = "Frame/network number bounds error"
    
        Case 87:
            msg = "Invalid frame type"
    
        Case 88:
            msg = "Invalid network number"
    
        Case 89:
            msg = "Duplicate network number"
    
        Case 90:
            msg = "Parameter out of bounds"
    
        Case 91:
            msg = "Access denied"
    
        Case 92:
            msg = "Out of memory"
    
        Case 93:
            msg = "Already exists"
    
        Case 94:
            msg = "Path, file or object not found"
    
        Case 95:
            msg = "Unable to notify service"
    
        Case 96:
            msg = "Unable to notify DNS service"
    
        Case 97:
            msg = "Interface not configurable"
    
        Case 98:
            msg = "Not all DHCP leases could be released/renewed"
    
        Case 100:
            msg = "DHCP not enabled on adapter"
    
        Case 2147786788#:
            msg = "WRITELOCK?!"
    
        Case Else:
            msg = "Other"

    End Select
    
    GetSetIpError = msg
End Function

Private Sub uStats_Click(Button As Integer, x As Single, y As Single)
    
    Dim colNetAdapters, objNetAdapter
    Dim strIPAddress, strSubnetMask, strGateway, strGatewaymetric
    Dim errEnable, errGateways
    Dim i As Long

    i = lstAdapter.ListIndex
    If i = -1 Then Exit Sub

    i = lstAdapter.ItemData(i)

    Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration " & "where Index=" & i)

    For Each objNetAdapter In colNetAdapters
        txtStats.Text = textstst(objNetAdapter)
    Next

    txtStats.Visible = True
    uCloseStats.Visible = True
    
    txtStats.ZOrder 0
    uCloseStats.ZOrder 0
    
End Sub

Private Sub uSlave_ActivateNextState(u_Cancel As Boolean, u_NewState As uCheckboxConstants)
    If PreventNext Then Exit Sub
    PreventNext = True
    If uMouseEnabled.Value = u_UnChecked Then
        u_Cancel = True
        u_NewState = u_UnChecked
        PreventNext = False
        Exit Sub
    End If

    If uSlave.Value = u_UnChecked Then
        uMaster.Value = u_UnChecked
    Else
        uMaster.Value = u_Checked
    End If

    PreventNext = False
End Sub





















Sub PairScores()
    Dim i As Long
    
    Dim strSplit() As String
    Dim StartCollecting As Boolean
    Dim CurrentCount As Long
    Dim Length As Long
    Dim currentSpan As String
    Dim RowCount As Long
    Dim previousSpan As String
    
    Dim EmptyScore As Score
    
    ReDim Scores(0)
    RowCount = 0
    
    i = 0
    'i = InStr(1, ScorePage, "Resultaten Hoofdfase bachelor")
    i = InStr(i + 1, ScorePage, """><b>Categorie</b></span>")
    
AllOverAgain:
    CurrentCount = 0
    previousSpan = "lolz"
    
    st ScorePage
    
    Do
        i = InStr(i + 1, ScorePage, "<p class=")
        If i = 0 Or InStr(i + 1, ScorePage, "</table>") < i + 60 Then GoTo EndOfCalculation
        
        currentSpan = GetMidText(ScorePage, "<p class=", "</p>", i)
        currentSpan = FilterShit(currentSpan)
        
        
        CurrentCount = CurrentCount + 1
        
        Dim kj As Long
        kj = InStr(i + 1, ScorePage, "</tr>") - i
        
        If currentSpan = "" And kj < 80 Then
            CurrentCount = 0
            If Scores(RowCount).VakCode <> "" Or Scores(RowCount).VakNaam <> "" Or Scores(RowCount).VakType <> "" Then
                RowCount = RowCount + 1
            Else
                Scores(RowCount) = EmptyScore
            End If
            ReDim Preserve Scores(0 To RowCount)
        End If
        
        'If currentSpan = "" And CurrentCount = 1 Then
        '    CurrentCount = 0
        'End If
        
        
        
        Select Case CurrentCount
            Case 1:
                'If Scores(RowCount).VakCode = "" Then CurrentCount = CurrentCount - 1
                    
                Scores(RowCount).VakCode = currentSpan
            Case 2:
                If Scores(RowCount).VakCode <> "" Then
                    Scores(RowCount).VakNaam = currentSpan
                    CurrentCount = CurrentCount + 1
                Else
                    Scores(RowCount).VakType = currentSpan
                End If
            Case 3:
                Scores(RowCount).VakNaam = currentSpan
            Case 4:
                If isDatum(currentSpan) Then
                    CurrentCount = CurrentCount + 1
                    Scores(RowCount).Datum = currentSpan
                Else
                    Scores(RowCount).Weging = currentSpan
                End If
            Case 5
                Scores(RowCount).Datum = currentSpan

            Case 6:
                Scores(RowCount).Cijfer = currentSpan
            Case 7:
                Scores(RowCount).StudiePunten = currentSpan
            Case 8:
                Scores(RowCount).HonorPunten = currentSpan
            Case 9:
                Scores(RowCount).Categorie = currentSpan
        End Select
        
        previousSpan = currentSpan
        
    Loop
    
EndOfCalculation:
    i = InStr(i + 1, ScorePage, """><b>Categorie</b></span>")
    If i <> 0 Then GoTo AllOverAgain
    
    For i = 0 To RowCount
        Debug.Print Scores(i).VakCode & " " & Scores(i).VakNaam & " " & Scores(i).Datum & " " & Scores(i).Cijfer & " " & Scores(i).StudiePunten & " " & Scores(RowCount).HonorPunten & " " & Scores(RowCount).Categorie
    Next i
    
    DrawScores
    'MsgBox "Done!"
End Sub

Function isDatum(strInput As String) As Boolean
    If strInput = "" Then
        isDatum = False
        Exit Function
    End If
    
    Dim tmpSplit() As String
    tmpSplit = Split(strInput, "-")
    isDatum = (UBound(tmpSplit) = 2)
End Function

Sub SortBy(intType As Integer)
    Dim i As Long
    Dim j As Long
    Dim n As Long
    
    Dim CountK As Long
    
    Dim K() As Score
    
    Dim hasChanged As Boolean
    Dim isBigger As Boolean
    
    ReDim K(0 To UBound(Scores))
    
    CountK = 0
    
    For i = 0 To UBound(Scores)
        
        For j = 0 To CountK
            If CountK = 0 Then
                K(j) = Scores(i)
                CountK = 1
                GoTo FindNextOne
            End If
            
            isBigger = False
            
            Select Case intType
                Case 0
                    isBigger = DateToNumber(Scores(i).Datum) > DateToNumber(K(j).Datum)
                Case 1
                    If Scores(i).Cijfer = "" Then
                        isBigger = False
                        GoTo InsertIt:
                    End If
                    
                    If K(j).Cijfer = "" Then isBigger = True
                    
                    'Debug.Assert Scores(i).Cijfer <> "V"
                    
                    If Scores(i).Cijfer = "V" And (K(j).Cijfer = "NVD" Or K(j).Cijfer = "O" Or getNumber(K(j).Cijfer) < 5.5) Then isBigger = True
                    
                    If getNumber(Scores(i).Cijfer) > getNumber(K(j).Cijfer) Then isBigger = True
                    
                    If K(j).Cijfer = "V" And getNumber(Scores(i).Cijfer) < 5.5 Then isBigger = False
                    
                    If K(j).Weging <> "" And Scores(i).Weging = "" And isBigger = False Then
                        isBigger = True
                    End If
                    
            End Select
            
InsertIt:
            
            If isBigger Then 'insert
                
                For n = CountK To j + 1 Step -1
                    K(n) = K(n - 1)
                Next n
                K(j) = Scores(i)
                CountK = CountK + 1
                GoTo FindNextOne:
            ElseIf j = CountK Then
                K(j) = Scores(i)
                CountK = CountK + 1
                GoTo FindNextOne:
            End If
            
            
        Next j
FindNextOne:
        
    Next i
    
    Scores = K
    
    DrawScores
End Sub

Function DateToNumber(strInput As String) As Long
    Dim strSplit() As String
    
    
    If strInput = "" Then
        DateToNumber = 0
        Exit Function
    End If
    
    strSplit = Split(strInput, "-")
    If UBound(strSplit) <> 2 Then
        DateToNumber = 0
        Exit Function
    End If
    
    DateToNumber = Val(strSplit(2) & strSplit(1) & strSplit(0))
End Function

Sub DrawScores()
    Const Headers As String = "Naam,Weging,Datum,Resultaat,Punten" 'Cursus,Honoraire punten,Categorie,Type
    Dim splitHeader() As String
    Dim i As Long
    Dim RowHeight As Long
    Dim headerOffset() As Long
    Dim headerWidth() As Long
    Dim strPrint As String
    Dim j As Long
    Dim OffsetY As Long
    Dim tmpHeight As Long
    
    loadScores.Visible = False
    loadScores.Loading = False
    
    picTab(4).Cls
    picTab(4).Picture = LoadPicture
    picTab(4).ForeColor = vbWhite
    picTab(4).FontName = "MS Sans Serif"
    picTab(4).FontSize = 8
    
    splitHeader = Split(Headers, ",")
    
    ReDim headerOffset(0 To UBound(splitHeader))
    ReDim headerWidth(0 To UBound(splitHeader))
    
    headerOffset(0) = 4
    RowHeight = picTab(4).TextHeight("WQpE") + 2
    
    'If scrollScores.Value > UBound(Scores) Then scrollScores.Value = UBound(Scores)
    
    For i = 0 To UBound(splitHeader)
        If i > 0 Then
            headerOffset(i) = headerOffset(i - 1) + headerWidth(i - 1) + 10
        End If
        
        picTab(4).FontBold = True
        picTab(4).FontItalic = True
        If picTab(4).TextWidth(splitHeader(i)) > headerWidth(i) Then
            headerWidth(i) = picTab(4).TextWidth(splitHeader(i))
        End If
        
        picTab(4).CurrentX = headerOffset(i)
        picTab(4).CurrentY = 3
        
        picTab(4).Print splitHeader(i)
        picTab(4).FontBold = False
        picTab(4).FontItalic = False
        OffsetY = 0
        
        For j = 0 To UBound(Scores)
            If Scores(j).Weging = "" And j > 0 Then OffsetY = OffsetY + RowHeight
            
            tmpHeight = (j + 1 - scrollScores.Value) * RowHeight + 3 + OffsetY
            'If Scores(j).Weging = "" Then picTab(4).Line (0, tmpHeight + RowHeight - 1)-(picTab(4).Width, tmpHeight + RowHeight - 1), vbWhite
            
            picTab(4).CurrentX = headerOffset(i)
            picTab(4).CurrentY = tmpHeight
            
            picTab(4).FontBold = (Scores(j).Weging = "")
            
            Select Case i
                'Case 0
                '     strPrint = Scores(j).VakCode
                Case 0
                    strPrint = Scores(j).VakNaam
                'Case 2
                '    strPrint = Scores(j).VakType
                Case 1
                    strPrint = Scores(j).Weging
                Case 2
                    strPrint = Scores(j).Datum
                Case 3
                    strPrint = Scores(j).Cijfer
                
                    If strPrint = "V" Or getNumber(strPrint) >= 5.5 Then
                        picTab(4).ForeColor = vbGreen
                    Else
                        picTab(4).ForeColor = vbRed
                    End If
                Case 4
                    strPrint = Scores(j).StudiePunten
'                Case 7
'                    strPrint = Scores(j).HonorPunten
'                Case 8
'                    strPrint = Scores(j).Categorie
            End Select
            
            If picTab(4).TextWidth(strPrint) > headerWidth(i) Then
                headerWidth(i) = picTab(4).TextWidth(strPrint)
            End If
            
            If tmpHeight >= RowHeight - 3 Then
                picTab(4).Print strPrint
            End If
            
            picTab(4).ForeColor = vbWhite
        Next j
    Next i
    
    scrollScores.Max = Round((((UBound(Scores) + 3) * RowHeight + 3 + OffsetY) - picTab(4).Height) / RowHeight)
    If scrollScores.Max > 0 Then
        scrollScores.LargeChange = scrollScores.Max / 4
    End If
    
    
    DoEvents
End Sub

Function getNumber(ByVal strInput As String) As Double
    On Error GoTo NotANumber
    
    getNumber = CDbl(strInput)
    If CStr(getNumber) <> strInput Then
        If InStr(1, strInput, ".") > 0 Then
            strInput = Replace(strInput, ".", ",")
        ElseIf InStr(1, strInput, ",") > 0 Then
            strInput = Replace(strInput, ",", ".")
        End If
        getNumber = CDbl(strInput)
    End If
NotANumber:
    
End Function

Function FilterShit(strInput As String) As String
    strInput = Replace(strInput, "<i>", "")
    strInput = Replace(strInput, "</i>", "")
    If InStr(1, strInput, "</span>") > 0 Then
        FilterShit = Replace(strInput, "</span>", "")
        FilterShit = Right(FilterShit, Len(FilterShit) - InStrRev(FilterShit, ">"))
    ElseIf InStr(1, strInput, "<br/>") > 0 Then
        FilterShit = ""
    Else
        FilterShit = strInput
    End If
    
    'FilterShit = Right(strInput, Len(strInput) - InStrRev(strInput, ">"))
End Function






Private Sub wb1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    
    
    If Progress = 0 And ProgressMax = 0 Then ' done
        'On Error Resume Next
        
        DoEvents
        'If isDocLoaded = False Then Exit Sub
        If MustLoginFirst Then
            FillFormAndSubmit
            Exit Sub
        End If
        
        If MustSubmitFirst Then
            SubmitSite
            Exit Sub
        End If
        
        If MustKlikUitvoeren Then
            SubmitUitvoeren
            Exit Sub
        End If
        
        If MustLoadScores Then
            ScorePage = GetUrlSource("https://student.osiris.hhs.nl/osiris_student_hhsprd/EmbeddedReport.do")
            PairScores
            'st ScorePage
        End If
        
        If InStr(1, wb1.LocationName, "https://student.osiris.hhs.nl/") > 0 Then
            If LoggedIn = False Then
                wb1.Navigate "https://student.osiris.hhs.nl/osiris_student_hhsprd/Dossier.do"
            End If
            
            LoggedIn = True
        End If
        
    End If
End Sub

Function isDocLoaded() As Boolean

    isDocLoaded = TypeName(wb1.Document.getElementById("body")) <> "Nothing"
End Function

Function MustLoadScores() As Boolean
    If wb1.Document Is Nothing Then Exit Function
    
    MustLoadScores = (wb1.LocationURL = "https://student.osiris.hhs.nl/osiris_student_hhsprd/ToonDossier.do")
        
End Function

Sub SubmitUitvoeren()
    If wb1.Document Is Nothing Then Exit Sub
    
    Dim links As Object
    Dim j As Object
    
    Set links = wb1.Document.getElementsByClassName("psbButtonLink")
    
    wb1.Document.getElementById("M__Id").Click
    
    DoEvents
    
    For Each j In links
        If InStr(1, j.onclick, "dossier") > 0 Then
            j.Click
            Exit Sub
        End If
    Next j
    
End Sub

Function MustKlikUitvoeren() As Boolean
    MustKlikUitvoeren = (wb1.LocationURL = "https://student.osiris.hhs.nl/osiris_student_hhsprd/Dossier.do")
End Function


Sub FillFormAndSubmit()
    If wb1.Document Is Nothing Then Exit Sub
    
    wb1.Document.getElementById("ContentPlaceHolder1_UsernameTextBox").Value = "13057499"
    wb1.Document.getElementById("ContentPlaceHolder1_PasswordTextBox").Value = "q01101981Q"
    wb1.Document.getElementById("ContentPlaceHolder1_SubmitButton").Click
End Sub

Sub SubmitSite()
    If wb1.Document Is Nothing Then Exit Sub
    
    wb1.Document.getElementById("ContentPlaceHolder1_PassiveSignInButton").Click
End Sub


Function MustSubmitFirst() As Boolean
    On Error GoTo endoffunction
    If wb1.Document Is Nothing Then Exit Function
    
    If LenB(wb1.Document.getElementById("ContentPlaceHolder1_PassiveSignInButton")) > 0 Then
        LoggedIn = False
        MustSubmitFirst = True
    End If
endoffunction:
End Function

Function MustLoginFirst() As Boolean
    On Error GoTo endoffunction
    
    If wb1.Document Is Nothing Then Exit Function
    If wb1.Document.ReadyState <> "complete" Then Exit Function
    
    'Debug.Print wb1.Document.getElementById("ContentPlaceHolder1_UsernameTextBox")
    If LenB(wb1.Document.getElementById("ContentPlaceHolder1_UsernameTextBox")) > 0 Then
        LoggedIn = False
        MustLoginFirst = True
    End If
    
endoffunction:
    
End Function


Function GetMidText(zTxt As String, zFind1 As String, zFind2 As String, Optional zStart As Long = 1) As String
    On Error GoTo ErrHandler:
    Dim totalStr1 As Long
    Dim totalStr2 As Long
    Dim totalStr3 As Long
    'On Error Resume Next
10  If zTxt = "" Or zFind1 = "" Or zFind2 = "" Then Exit Function

20  totalStr1 = InStr(zStart, zTxt, zFind1)
30  If totalStr1 = 0 Then Exit Function
40  totalStr1 = totalStr1 + Len(zFind1)

50  totalStr2 = InStr(totalStr1, zTxt, zFind2)
60  If totalStr2 = 0 Then Exit Function


70  totalStr3 = totalStr2 - totalStr1
80  GetMidText = Mid(zTxt, totalStr1, totalStr3)
90  Exit Function
ErrHandler:
    'If chkError.Value = vbChecked Then Resume Next
    'ErrLogger "getMidText()", Err.Number, Err.Description, False, Erl()
End Function
