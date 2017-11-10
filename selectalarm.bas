Attribute VB_Name = "SelectAlarm"
' Klepsydra Project
' Open Select File Common Dialog, using WinAPI call (comdlg32.dll)

Option Explicit

Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Function SelectAlarmFileDialog() As String
    Dim strTemp, strTemp1, pathStr, sFilter, Fname As String
    Dim i, n, j, lReturn As Long
    Dim OpenFile As OPENFILENAME
    
    OpenFile.lStructSize = Len(OpenFile)
    sFilter = "Wave Audio (*.wav)" & Chr(0) & "*.WAV" & Chr(0)
    
    With OpenFile
        .lpstrFilter = sFilter
        .nFilterIndex = 1
        .lpstrFile = String(257, 0)
        .nMaxFile = Len(OpenFile.lpstrFile) - 1
        .lpstrFileTitle = OpenFile.lpstrFile
        .nMaxFileTitle = OpenFile.nMaxFile
        .lpstrTitle = "Select Alarm Sound File"
        .flags = 0
    End With
    
    lReturn = GetOpenFileName(OpenFile)
    
    If lReturn = 0 Then
        SelectAlarmFileDialog = "alarm.wav"
    Else
        SelectAlarmFileDialog = OpenFile.lpstrFile
    End If
End Function
