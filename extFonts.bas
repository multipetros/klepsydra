Attribute VB_Name = "ExtFonts"
' Klepsydra Project
' Add & Remove external font resources using WinAPI call (gdi32)

Option Explicit

Private Declare Function AddFontResourceEx Lib "gdi32" _
                         Alias "AddFontResourceExA" ( _
                           ByVal sFileName As String, _
                           ByVal lFlags As Long, _
                           ByVal lReserved As Long _
                         ) As Long

Private Declare Function RemoveFontResourceEx Lib "gdi32" _
                         Alias "RemoveFontResourceExA" ( _
                           ByVal sFileName As String, _
                           ByVal lFlags As Long, _
                           ByVal lReserved As Long _
                         ) As Long

Const FR_PRIVATE As Long = &H10

Public Function LoadFont(path As String) As Long
    LoadFont = AddFontResourceEx(path, FR_PRIVATE, 0&)
End Function

Public Function UnloadFont(path As String) As Long
    UnloadFont = RemoveFontResourceEx(path, FR_PRIVATE, 0&)
End Function
