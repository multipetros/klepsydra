Attribute VB_Name = "IniRW"
' Klepsydra Project
' INI file Read & Write functions using WinAPI calls

Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String _
    ) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String _
    ) As Long
                        
Public Function IniWrite(sSection As String, sKeyName As String, sNewString As String, sINIFileName As String) As Boolean
    Call WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName)
    IniWrite = (Err.Number = 0)
End Function

Public Function IniRead(sSection As String, sKeyName As String, sINIFileName As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    IniRead = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), sINIFileName))
End Function

