VERSION 5.00
Begin VB.Form frmTempStart 
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6570
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
   ScaleHeight     =   5355
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
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
   Begin Project1.uTextBox uTextBox1 
      Height          =   5070
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   8943
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
      RowLines        =   -1  'True
      RowLineColor    =   8421504
      RowNumberOnEveryLine=   -1  'True
   End
End
Attribute VB_Name = "frmTempStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    uTextBox1.clear
    
End Sub

Private Sub Form_Load()
    uTextBox1_Changed
End Sub

Private Sub Form_Resize()
    uTextBox1.RedrawPause

    uTextBox1.Left = 0
    uTextBox1.Top = 0
    uTextBox1.Height = Me.ScaleHeight
    uTextBox1.Width = Me.ScaleWidth

    uTextBox1.RedrawResume

    Dim i As Long
    For i = 0 To 100
        uListBox.AddItem "item nr" & i, , , RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    Next i
    
    HScroll1.Max = uTextBox1.m_lScrollLeftMax
    
End Sub

Private Sub HScroll1_Change()
    HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
    uTextBox1.m_lScrollLeft = HScroll1.Value
    uTextBox1.Redraw
    
End Sub

Private Sub uTextBox1_Changed()
    Dim i As Long
    Dim s As String
    Dim t As String
    
    
    s = uTextBox1.Text
    uTextBox1.RedrawPause
    
    For i = 1 To Len(s) - 1
        t = Mid$(s, i, 1)
        uTextBox1.setCharForeColor i - 1, -1
        uTextBox1.setCharBold i - 1, False
            
        Select Case t
            
                
            Case "(", "e"
                uTextBox1.setCharForeColor i - 1, vbRed
                
            Case "d"
                uTextBox1.setCharBold i - 1, True
                
        End Select
        
    Next i
    
    
    uTextBox1.RedrawResume
    
    HScroll1.Max = uTextBox1.m_lScrollLeftMax
End Sub

