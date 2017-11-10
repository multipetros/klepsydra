Attribute VB_Name = "OpenURL"
' Klepsydra Project
' Open Browser function using WinAPI call (shell32.dll)

Option Explicit

Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Public Function OpenBrowser(url As String) As Long
    OpenBrowser = ShellExecute(0, "open", url, 0, 0, 1)
End Function

