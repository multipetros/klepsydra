Attribute VB_Name = "Sound"
' Klepsydra Project
' PlaySound function usin WinAPI calls (winmm.dll)

Option Explicit

Public Const SND_APPLICATION As Long = &H80
Public Const SND_ALIAS As Long = &H10000
Public Const SND_ID As Long = &H110000
Public Const SND_ASYNC As Long = &H1
Public Const SND_FILENAME As Long = &H20000
Public Const SND_LOOP As Long = &H8
Public Const SND_MEMORY As Long = &H4
Public Const SND_NODEFAULT As Long = &H2
Public Const SND_NOSTOP As Long = &H10
Public Const SND_NOWAIT As Long = &H2000
Public Const SND_PURGE As Long = &H40
Public Const SND_RESOURCE As Long = &H40004
Public Const SND_SYNC As Long = &H0

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" ( _
    ByVal lpszName As String, _
    ByVal hModule As Long, _
    ByVal dwFlags As Long _
    ) As Long

