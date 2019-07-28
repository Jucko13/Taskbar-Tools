Attribute VB_Name = "Globals"
Option Explicit

'Enum ActType
'    OpenUrl = 0
'    OpenProgram = 1
'
'End Enum


Type Sett
    sButtonColor As Long
    sButtonFontColor As Long

    sButtonText(0 To 2) As String
    sActionPath(0 To 2) As String
    sActionFolder(0 To 2) As String
    sActionParameters(0 To 2) As String

    sActivationText As String
End Type

Type Sett2
    sUploadArguments As String
End Type

Global allSettings() As Sett
Global otherSettings As Sett2

Global settingCount As Long
Global ActiveSetting As Long

Global Const AppName As String = "TaskBarToolsByRicardo"

Global ButtonColors(0 To 79) As Long


