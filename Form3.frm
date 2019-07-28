VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H00584D43&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   LinkTopic       =   "Form3"
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   319
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.uFrame frInfo 
      Height          =   900
      Left            =   15
      TabIndex        =   35
      Top             =   6120
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   1588
      BackgroundColor =   5786947
      ForeColor       =   16777215
      Caption         =   "Text Options"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label lblSettings 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"Form3.frx":0000
         ForeColor       =   &H00FFFFFF&
         Height          =   585
         Index           =   15
         Left            =   90
         TabIndex        =   36
         Top             =   225
         Width           =   3105
      End
   End
   Begin Project1.uButton uSave 
      Height          =   315
      Left            =   3150
      TabIndex        =   5
      Top             =   7035
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      Caption         =   "Save"
      BorderAnimation =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseOverBackgroundColor=   8421504
   End
   Begin Project1.uButton uClose 
      Height          =   315
      Left            =   15
      TabIndex        =   4
      Top             =   7035
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      Caption         =   "Close"
      BorderAnimation =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseOverBackgroundColor=   8421504
   End
   Begin VB.Timer tmrTopMost 
      Interval        =   100
      Left            =   3825
      Top             =   3390
   End
   Begin Project1.uFrame frSettings 
      Height          =   5475
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   9657
      BackgroundColor =   5786947
      ForeColor       =   16777215
      Caption         =   "Settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Project1.uFrame frUploadSettings 
         Height          =   465
         Left            =   165
         TabIndex        =   14
         Top             =   4830
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   820
         BackgroundColor =   5786947
         ForeColor       =   16777215
         Caption         =   "Upload Arguments"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtUpload2560 
            BorderStyle     =   0  'None
            Height          =   210
            Left            =   1740
            TabIndex        =   15
            Top             =   180
            Width           =   2625
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Arduino Mega 2560:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   16
            Top             =   180
            Width           =   1440
         End
      End
      Begin Project1.uFrame frSearchShortcuts 
         Height          =   4560
         Left            =   165
         TabIndex        =   1
         Top             =   195
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   8043
         BackgroundColor =   5786947
         ForeColor       =   16777215
         Caption         =   "Search Shortcuts"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtActivationText 
            BorderStyle     =   0  'None
            Height          =   210
            Left            =   1725
            TabIndex        =   37
            Top             =   540
            Width           =   2625
         End
         Begin VB.TextBox txtButtonText 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   2
            Left            =   1725
            TabIndex        =   33
            Top             =   3450
            Width           =   2625
         End
         Begin VB.TextBox txtButtonText 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   1
            Left            =   1725
            TabIndex        =   31
            Top             =   2310
            Width           =   2625
         End
         Begin VB.TextBox txtButtonText 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   0
            Left            =   1725
            TabIndex        =   29
            Top             =   1170
            Width           =   2625
         End
         Begin VB.TextBox txtPath 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   2
            Left            =   1725
            TabIndex        =   25
            Top             =   3720
            Width           =   2625
         End
         Begin VB.TextBox txtParameters 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   2
            Left            =   1725
            TabIndex        =   24
            Top             =   3990
            Width           =   2625
         End
         Begin VB.TextBox txtFolder 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   2
            Left            =   1725
            TabIndex        =   23
            Top             =   4260
            Width           =   2625
         End
         Begin VB.TextBox txtPath 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   1
            Left            =   1725
            TabIndex        =   19
            Top             =   2580
            Width           =   2625
         End
         Begin VB.TextBox txtParameters 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   1
            Left            =   1725
            TabIndex        =   18
            Top             =   2850
            Width           =   2625
         End
         Begin VB.TextBox txtFolder 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   1
            Left            =   1725
            TabIndex        =   17
            Top             =   3120
            Width           =   2625
         End
         Begin VB.TextBox txtFolder 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   0
            Left            =   1725
            TabIndex        =   11
            Top             =   1980
            Width           =   2625
         End
         Begin VB.TextBox txtParameters 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   0
            Left            =   1725
            TabIndex        =   10
            Top             =   1710
            Width           =   2625
         End
         Begin VB.TextBox txtPath 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   0
            Left            =   1725
            TabIndex        =   8
            Top             =   1440
            Width           =   2625
         End
         Begin Project1.uDropDown uOptions 
            Height          =   285
            Left            =   1725
            TabIndex        =   2
            Top             =   180
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   503
            BackgroundColor =   5786947
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
            Text            =   ""
            SelectionBackgroundColor=   5786947
            SelectionBorderColor=   14737632
         End
         Begin Project1.uDropDown uButtonColor 
            Height          =   240
            Left            =   1725
            TabIndex        =   12
            Top             =   810
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   423
            BackgroundColor =   5786947
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
            Text            =   ""
            SelectionBackgroundColor=   5786947
            SelectionBorderColor=   14737632
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Activation Text:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   16
            Left            =   105
            TabIndex        =   38
            Top             =   555
            Width           =   1365
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Button Text 3:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   14
            Left            =   105
            TabIndex        =   34
            Top             =   3465
            Width           =   1005
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Button Text 2:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   13
            Left            =   105
            TabIndex        =   32
            Top             =   2325
            Width           =   1005
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Button Text 1:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   12
            Left            =   105
            TabIndex        =   30
            Top             =   1185
            Width           =   1005
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Folder Path 3:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   11
            Left            =   105
            TabIndex        =   28
            Top             =   4245
            Width           =   990
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Program Parameters 3:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   10
            Left            =   105
            TabIndex        =   27
            Top             =   3990
            Width           =   1605
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File Path 3:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   9
            Left            =   105
            TabIndex        =   26
            Top             =   3735
            Width           =   795
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Folder Path 2:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   8
            Left            =   105
            TabIndex        =   22
            Top             =   3105
            Width           =   990
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Program Parameters 2:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   7
            Left            =   105
            TabIndex        =   21
            Top             =   2850
            Width           =   1605
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File Path 2:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   6
            Left            =   105
            TabIndex        =   20
            Top             =   2595
            Width           =   795
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Theme Color:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   105
            TabIndex        =   13
            Top             =   810
            Width           =   945
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File Path 1:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   9
            Top             =   1455
            Width           =   795
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Program Parameters 1:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   7
            Top             =   1710
            Width           =   1605
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Folder Path 1:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   105
            TabIndex        =   6
            Top             =   1965
            Width           =   990
         End
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Search Option:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   3
            Top             =   210
            Width           =   1560
         End
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ChangeDropDown As Boolean


Sub Form_Load()
    Dim i As Long

    Dim prevSelected As Long
    prevSelected = uOptions.ListIndex

    uOptions.Clear
    For i = 1 To settingCount
        uOptions.AddItem allSettings(i - 1).sButtonText(0), , , ButtonColors(allSettings(i - 1).sButtonColor)
    Next i

    uButtonColor.Clear
    For i = 0 To UBound(ButtonColors)
        uButtonColor.AddItem CStr(i), , , ButtonColors(i)
    Next i

    txtUpload2560.Text = otherSettings.sUploadArguments

    If prevSelected > -1 Then
        uOptions.ListIndex = prevSelected

    End If
End Sub



Private Sub Form_Resize()
    Dim i As Long
    Dim j As Long


    Dim c As Control


    frSettings.Width = Me.ScaleWidth - 2
    frSearchShortcuts.Width = Me.Width - frSearchShortcuts.Left * 2 - 2 * Screen.TwipsPerPixelX
    frInfo.Width = Me.ScaleWidth - 2
    frUploadSettings.Width = frSearchShortcuts.Width

    uSave.Left = Me.ScaleWidth - uSave.Width - 1

    For Each c In Me.Controls
        'Debug.Print TypeName(c)
        '
        If TypeName(c) = "TextBox" Or TypeName(c) = "uDropDown" Then
            'Debug.Print c.Container.Name
            If c.Container.Name = frSearchShortcuts.Name Or c.Container.Name = frUploadSettings.Name Then
                c.Width = frSearchShortcuts.Width - c.Left - 5 * Screen.TwipsPerPixelX
            End If
        End If

    Next



End Sub

Private Sub txtButtonText_Change(Index As Integer)
    EditSettings
End Sub

Private Sub txtFolder_Change(Index As Integer)
    EditSettings
End Sub

Private Sub txtParameters_Change(Index As Integer)
    EditSettings
End Sub

Private Sub txtPath_Change(Index As Integer)
    EditSettings
End Sub

Private Sub txtUpload2560_Change()
    EditSettings
End Sub

Private Sub uButtonColor_ItemChange(ItemIndex As Long)
    If uOptions.ListIndex > -1 Then
        'uOptions.ItemColor(uOptions.ListIndex) = ButtonColors(ItemIndex)
        frmMain.SetProgramColor ButtonColors(ItemIndex), vbWhite
    End If

    EditSettings
End Sub

Private Sub uClose_Click(Button As Integer, X As Single, Y As Single)
    Me.Hide
    frmMain.SetSearchingMode
End Sub

Private Sub uOptions_ItemChange(ItemIndex As Long)
    Dim i As Long

    ChangeDropDown = True

    For i = 0 To UBound(allSettings(ItemIndex).sActionPath)
        txtButtonText(i).Text = allSettings(ItemIndex).sButtonText(i)
        txtPath(i).Text = allSettings(ItemIndex).sActionPath(i)
        txtFolder(i).Text = allSettings(ItemIndex).sActionFolder(i)
        txtParameters(i).Text = allSettings(ItemIndex).sActionParameters(i)
    Next i

    uButtonColor.ListIndex = allSettings(ItemIndex).sButtonColor
    txtActivationText.Text = allSettings(ItemIndex).sActivationText

    'txtPath.Text = allSettings(ItemIndex).sButtonColor

    frmMain.SetProgramColor ButtonColors(uButtonColor.ListIndex), vbWhite

    ChangeDropDown = False
End Sub


Sub EditSettings()
    Dim ItemIndex As Long
    Dim i As Long

    If ChangeDropDown Then Exit Sub

    ItemIndex = uOptions.ListIndex
    If ItemIndex < 0 Then Exit Sub

    For i = 0 To UBound(allSettings(ItemIndex).sActionPath)
        allSettings(ItemIndex).sButtonText(i) = txtButtonText(i).Text
        allSettings(ItemIndex).sActionPath(i) = txtPath(i).Text
        allSettings(ItemIndex).sActionFolder(i) = txtFolder(i).Text
        allSettings(ItemIndex).sActionParameters(i) = txtParameters(i).Text
    Next i

    allSettings(ItemIndex).sButtonColor = uButtonColor.ListIndex
    allSettings(ItemIndex).sActivationText = txtActivationText.Text

    otherSettings.sUploadArguments = txtUpload2560.Text

End Sub

Private Sub uSave_Click(Button As Integer, X As Single, Y As Single)
    frmMain.SaveAllSettings
    Me.Hide
End Sub
