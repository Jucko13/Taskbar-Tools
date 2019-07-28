VERSION 5.00
Begin VB.Form frmTempStart 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11445
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
   ScaleHeight     =   8715
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRedrawTextbox 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   780
      Top             =   5610
   End
   Begin VB.Timer tmrAddChar 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   300
      Top             =   5535
   End
   Begin Project1.uTextBox uTextBox1 
      Height          =   7980
      Left            =   2040
      TabIndex        =   0
      Top             =   585
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   14076
      BackgroundColor =   0
      BorderColor     =   0
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
      RowLineColor    =   14737632
      RowNumberOnEveryLine=   -1  'True
      WordWrap        =   -1  'True
      MultiLine       =   -1  'True
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "clear"
      Height          =   360
      Left            =   1800
      TabIndex        =   3
      Top             =   4800
      Width           =   990
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   100
      Left            =   240
      Max             =   100
      SmallChange     =   10
      TabIndex        =   2
      Top             =   4200
      Width           =   5055
   End
   Begin Project1.uListBox uListBox 
      Height          =   2055
      Left            =   7320
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   3625
      BackgroundColor =   8421631
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
      Border          =   0   'False
      SelectionBackgroundColor=   12632319
      SelectionBorderColor=   255
      ItemHeight      =   27
   End
   Begin Project1.uFrame UFrame 
      Height          =   3780
      Left            =   150
      TabIndex        =   4
      Top             =   60
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6668
      BorderColor     =   10197915
      Caption         =   "Frame with guide dots"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Project1.uDropDown UDrUFrame 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   3165
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BackgroundColor =   14474460
         BorderColor     =   10197915
         SelectionBackgroundColor=   14474460
         SelectionBorderColor=   10197915
         SelectionBackgroundColorDisabled=   16777215
         SelectionBorderColorDisabled=   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "uFrame"
         ItemHeight      =   20
      End
      Begin Project1.uListBox ULiUFrame 
         Height          =   1455
         Left            =   120
         TabIndex        =   10
         Top             =   1605
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2566
         BackgroundColor =   14474460
         BorderColor     =   10197915
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "uFrame"
         SelectionBackgroundColor=   10197915
         SelectionBorderColor=   10197915
         SelectionForeColor=   16777215
         ItemHeight      =   19
      End
      Begin Project1.uCheckBox UChUFrame 
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   8
         Top             =   2685
         Width           =   1575
         _ExtentX        =   2381
         _ExtentY        =   661
         BackgroundColor =   14474460
         BorderColor     =   10197915
         Caption         =   "Checkbox"
         CheckBorderColor=   10197915
         CheckSelectionColor=   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   3
      End
      Begin Project1.uOptionBox UOpUFrame 
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         Top             =   2685
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackgroundColor =   14474460
         BorderColor     =   10197915
         Caption         =   "Optionbox"
         CheckBorderColor=   10197915
         CheckSelectionColor=   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
      End
      Begin Project1.uLoadBar uLoadBar1 
         Height          =   975
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         Top             =   1605
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         BackgroundColor =   14474460
         BarColor        =   8421504
         BarWidth        =   10
         BorderColor     =   10197915
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Value           =   30
      End
      Begin Project1.uButton UBuButton 
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   5
         Top             =   3165
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BackgroundColor =   14474460
         BorderColor     =   10197915
         ForeColor       =   0
         FocusColor      =   0
         BackgroundColorDisabled=   0
         BorderColorDisabled=   0
         ForeColorDisabled=   0
         CaptionBorderColorDisabled=   0
         FocusColorDisabled=   0
         BorderAnimation =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionOffsetTop=   -1
      End
      Begin Project1.uCheckBox UChUFrame 
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   9
         Top             =   2685
         Width           =   1575
         _ExtentX        =   2381
         _ExtentY        =   661
         BackgroundColor =   14474460
         BorderColor     =   10197915
         Caption         =   "Checkbox"
         CheckBorderColor=   10197915
         CheckSelectionColor=   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
      End
      Begin Project1.uLoadBar uLoadBar1 
         Height          =   975
         Index           =   1
         Left            =   3840
         TabIndex        =   11
         Top             =   1605
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1720
         BackgroundColor =   14474460
         BarColor        =   8421504
         BarType         =   1
         BarWidth        =   0
         BorderColor     =   10197915
         Caption         =   "30"
         CaptionType     =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Value           =   30
      End
      Begin Project1.uLoadBar uLoadBar1 
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   12
         Top             =   1605
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BackgroundColor =   14474460
         BarColor        =   8421504
         BarType         =   0
         BarWidth        =   0
         BorderColor     =   10197915
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Value           =   30
      End
      Begin Project1.uLoadBar uLoadBar1 
         Height          =   975
         Index           =   3
         Left            =   2760
         TabIndex        =   13
         Top             =   1605
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         BackgroundColor =   14474460
         BarColor        =   8421504
         BarWidth        =   10
         BorderColor     =   10197915
         Caption         =   "Loa.."
         CaptionType     =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Loading         =   -1  'True
         LoadingSpeed    =   1
         Value           =   30
      End
      Begin Project1.uLoadBar uLoadBar1 
         Height          =   495
         Index           =   4
         Left            =   4440
         TabIndex        =   15
         Top             =   2085
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         BackgroundColor =   14474460
         BarColor        =   8421504
         BarType         =   0
         BarWidth        =   5
         BorderColor     =   10197915
         Caption         =   ""
         CaptionType     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Value           =   45
      End
      Begin Project1.uButton UBuButton 
         Height          =   375
         Index           =   1
         Left            =   5040
         TabIndex        =   16
         Top             =   3165
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackgroundColor =   14474460
         BorderColor     =   10197915
         ForeColor       =   0
         FocusColor      =   0
         BackgroundColorDisabled=   0
         BorderColorDisabled=   0
         ForeColorDisabled=   0
         CaptionBorderColorDisabled=   0
         FocusColorDisabled=   0
         BorderAnimation =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionOffsetTop=   -1
      End
   End
End
Attribute VB_Name = "frmTempStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    uTextBox1.Clear
    
End Sub

Private Sub Form_Load()
    uTextBox1_Changed
    
    ULiUFrame.AddItem "listitem 1"
    ULiUFrame.AddItem "item 2"
    ULiUFrame.AddItem "apples"
    ULiUFrame.AddItem "banana"
    ULiUFrame.AddItem "coconut"
    ULiUFrame.AddItem "item 6"
    
    
    UDrUFrame.AddItem "listitem 1"
    UDrUFrame.AddItem "item 2"
    UDrUFrame.AddItem "apples"
    UDrUFrame.AddItem "banana"
    UDrUFrame.AddItem "coconut"
    UDrUFrame.AddItem "item 6"
End Sub

Private Sub Form_Resize()
    Dim i As Long
    
    
    uTextBox1.RedrawPause

    'uTextBox1.Left = 0
    'uTextBox1.Top = 0
    'uTextBox1.Height = Me.ScaleHeight
    'uTextBox1.Width = Me.ScaleWidth

    uTextBox1.RedrawResume
    i = IIf(1, 1, 1)

    For i = 0 To 100
        uListBox.AddItem "item nr" & i, , , RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    Next i
    
    'HScroll1.Max = uTextBox1.m_lScrollLeftMax
    
End Sub

Private Sub HScroll1_Change()
    HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
    uTextBox1.m_lScrollLeft = HScroll1.Value
    uTextBox1.Redraw
    
End Sub

Private Sub tmrAddChar_Timer()
    uTextBox1.RedrawPause
    uTextBox1.AddCharAtCursor Chr(Rnd * 26 + 65), False
End Sub

Private Sub tmrRedrawTextbox_Timer()
    uTextBox1.RedrawResume
End Sub

Private Sub uTextBox1_Changed()
    
'    Dim i As Long
'    Dim s As String
'    Dim t As String
'
'
'    s = uTextBox1.Text
'    uTextBox1.RedrawPause
'
'    For i = 1 To Len(s) - 1
'        t = Mid$(s, i, 1)
'        uTextBox1.setCharForeColor i - 1, -1
'        uTextBox1.setCharBold i - 1, False
'
'        Select Case t
'
'
'            Case "(", "e"
'                uTextBox1.setCharForeColor i - 1, vbRed
'                uTextBox1.setCharItallic i - 1, True
'
'            Case "d"
'                uTextBox1.setCharBold i - 1, True
'
'        End Select
'
'    Next i
'
'
'    uTextBox1.RedrawResume
'
'    HScroll1.Max = uTextBox1.m_lScrollLeftMax
End Sub

