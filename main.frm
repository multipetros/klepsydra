VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Klepsydra"
   ClientHeight    =   1320
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   2670
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2670
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerColorBlink 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   2160
      Top             =   720
   End
   Begin VB.ComboBox ComboMins 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.ComboBox ComboSecs 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.ComboBox ComboHours 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton CommandStart 
      Caption         =   "S&tart"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer TimerCountdown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   720
   End
   Begin VB.CommandButton CommandStop 
      Cancel          =   -1  'True
      Caption         =   "Stop"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton CommandDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label LabelCountdown 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Begin VB.Menu MenuMute 
         Caption         =   "&Mute"
         Shortcut        =   ^M
      End
      Begin VB.Menu MenuLoop 
         Caption         =   "&Loop Alarm"
         Checked         =   -1  'True
         Shortcut        =   ^L
      End
      Begin VB.Menu MenuAlarmSound 
         Caption         =   "Ala&rm Sound"
         Shortcut        =   ^S
      End
      Begin VB.Menu MenuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MenuLanguage 
      Caption         =   "Lang&uage"
      Begin VB.Menu MenuEnglish 
         Caption         =   "&English"
      End
      Begin VB.Menu MenuGreek 
         Caption         =   "&Greek"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MenuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MenuLicense 
         Caption         =   "Li&cense"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Klepsydra Project Main Form
' Copyright (c) 2017, Petros Kyladitis <www.multipetros.gr>
'
' Klepsydra is a Countdown timer prgram with sound alarm
' It's open source, distributed under the GNU GPL3

Option Explicit

Dim Countdown As Date
Dim AlarmFile As String
Dim langID As Integer
Dim iniPath As String

Private Sub Form_Load()
    Dim i, iniRNum As Integer
    Dim iniR As String
    
    iniPath = SpecialFolder(feUserAppData) & "\klepsydra.ini"
    
    For i = 0 To 60
        If i < 24 Then
            ComboHours.AddItem (Format(i, "00"))
        End If
        ComboSecs.AddItem (Format(i, "00"))
        ComboMins.AddItem (Format(i, "00"))
    Next i
    
    iniR = IniRead("main", "hours", iniPath)
    iniRNum = CInt(IIf(IsNumeric(iniR), iniR, 0))
    If iniRNum > -1 And iniRNum < 25 Then
        ComboHours.ListIndex = iniRNum
    End If
    
    iniR = IniRead("main", "mins", iniPath)
    iniRNum = CInt(IIf(IsNumeric(iniR), iniR, 0))
    If iniRNum > -1 And iniRNum < 61 Then
        ComboMins.ListIndex = iniRNum
    End If
    
    iniR = IniRead("main", "secs", iniPath)
    iniRNum = CInt(IIf(IsNumeric(iniR), iniR, 10))
    If iniRNum > -1 And iniRNum < 61 Then
        ComboSecs.ListIndex = iniRNum
    Else
        ComboSecs.ListIndex = 5
    End If
    
    iniR = IniRead("main", "loop", iniPath)
    If iniR = "False" Then
        MenuLoop.Checked = False
    Else
        MenuLoop.Checked = True
    End If
    
    iniR = IniRead("main", "alarm", iniPath)
    If iniR = "" Then
        AlarmFile = "alarm.wav"
    End If
    
    iniR = IniRead("main", "mute", iniPath)
    If iniR = "True" Then
        MenuMute.Checked = True
    Else
        MenuMute.Checked = False
    End If
    
    iniR = IniRead("main", "language", iniPath)
    If iniR = "gr" Then
        langID = 200
        LoadStrings
        MenuGreek.Checked = True
        MenuEnglish.Checked = False
    Else
        langID = 100
        LoadStrings
        MenuGreek.Checked = False
        MenuEnglish.Checked = True
    End If
    
    ShowStartControls
End Sub
Private Sub LoadStrings()
    MenuFile.Caption = LoadResString(langID + 1)
    MenuExit.Caption = LoadResString(langID + 2)
    MenuLoop.Caption = LoadResString(langID + 3)
    MenuAlarmSound.Caption = LoadResString(langID + 4)
    MenuHelp.Caption = LoadResString(langID + 5)
    MenuAbout.Caption = LoadResString(langID + 6)
    MenuLicense.Caption = LoadResString(langID + 7)
    CommandStart.Caption = LoadResString(langID + 8)
    CommandStop.Caption = LoadResString(langID + 9)
    CommandDone.Caption = LoadResString(langID + 10)
    MenuLanguage.Caption = LoadResString(langID + 11)
    MenuEnglish.Caption = LoadResString(langID + 12)
    MenuGreek.Caption = LoadResString(langID + 13)
    Me.Caption = LoadResString(langID + 18)
    MenuMute.Caption = LoadResString(langID + 19)
End Sub

Private Sub CommandDone_Click()
    Dim Wnd As Long
    PlaySound vbNullString, 0, 0
    Wnd = SetTopMostWindow(Me.hwnd, False)
    ShowStartControls
End Sub

Private Sub CommandStop_Click()
    TimerCountdown.Enabled = False
    ShowStartControls
End Sub

Private Sub CommandStart_Click()
    Countdown = ComboHours.List(ComboHours.ListIndex) & ":" _
        & ComboMins.List(ComboMins.ListIndex) & ":" _
        & ComboSecs.List(ComboSecs.ListIndex)
    TimerCountdown.Enabled = True
    ShowStopControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim iniW As Boolean
    iniW = IniWrite("main", "hours", ComboHours.Text, iniPath)
    iniW = IniWrite("main", "mins", ComboMins.Text, iniPath)
    iniW = IniWrite("main", "secs", ComboSecs.Text, iniPath)
    iniW = IniWrite("main", "secs", ComboSecs.Text, iniPath)
    iniW = IniWrite("main", "loop", MenuLoop.Checked, iniPath)
    iniW = IniWrite("main", "alarm", AlarmFile, iniPath)
    iniW = IniWrite("main", "mute", MenuMute.Checked, iniPath)
    Dim langStr As String
    If langID = 200 Then
        langStr = "gr"
    Else
        langStr = "en"
    End If
    iniW = IniWrite("main", "language", langStr, iniPath)
End Sub

Private Sub MenuAbout_Click()
    MsgBox LoadResString(langID + 14), vbInformation + vbOKOnly, LoadResString(langID + 15)
End Sub

Private Sub MenuAlarmSound_Click()
    AlarmFile = SelectAlarmFileDialog()
End Sub

Private Sub MenuEnglish_Click()
    MenuEnglish.Checked = True
    MenuGreek.Checked = False
    langID = 100
    LoadStrings
End Sub

Private Sub MenuExit_Click()
    Unload Me
End Sub

Private Sub MenuGreek_Click()
    MenuGreek.Checked = True
    MenuEnglish.Checked = False
    langID = 200
    LoadStrings
End Sub

Private Sub MenuLicense_Click()
    Dim LicFile As String
    Dim TaskID As Double
    LicFile = App.Path & "\license.txt"
    If FileExists(LicFile) = True Then
        On Error GoTo ErrMsg
        TaskID = Shell("notepad.exe " & LicFile, vbNormalFocus)
        Exit Sub
    End If
ErrMsg:
    Dim res As VbMsgBoxResult
    Dim browse As Long
    res = MsgBox(LoadResString(langID + 16), vbCritical + vbYesNo, LoadResString(langID + 17))
    If res = vbYes Then
        browse = OpenBrowser("http://www.gnu.org/licenses/gpl.html")
    End If
End Sub

Private Sub MenuLoop_Click()
    MenuLoop.Checked = Not MenuLoop.Checked
End Sub

Private Sub MenuMute_Click()
    MenuMute.Checked = Not MenuMute.Checked
End Sub

Private Sub TimerColorBlink_Timer()
    Dim bColor As Long
    Dim fColor As Long
    bColor = LabelCountdown.BackColor
    fColor = LabelCountdown.ForeColor
    LabelCountdown.ForeColor = bColor
    LabelCountdown.BackColor = fColor
End Sub

Private Sub TimerCountdown_Timer()
    Dim TimeStr As String
    
    Countdown = Countdown - (1 / 24 / 60 / 60)
    TimeStr = Format(Countdown, "hh:mm:ss")
    LabelCountdown.Caption = TimeStr
    If TimeStr = "00:00:00" Then
        Dim Snd, SndParams, Wnd As Long
        If MenuLoop.Checked Then
            SndParams = SND_FILENAME Or SND_LOOP Or SND_ASYNC
        Else
            SndParams = SND_FILENAME Or SND_ASYNC
        End If
        TimerCountdown.Enabled = False
        Me.SetFocus
        Wnd = SetTopMostWindow(Me.hwnd, True)
        If MenuMute.Checked = False Then
            Snd = PlaySound(AlarmFile, 0, SndParams)
        End If
        ShowDoneControls
    End If
End Sub

Private Sub ShowStartControls()
    TimerColorBlink.Enabled = False
    CommandStart.Visible = True
    CommandStop.Visible = False
    CommandDone.Visible = False
    LabelCountdown.Visible = False
    ComboHours.Visible = True
    ComboMins.Visible = True
    ComboSecs.Visible = True
End Sub

Private Sub ShowStopControls()
    LabelCountdown.ForeColor = vbButtonText
    LabelCountdown.BackColor = vbButtonFace
    CommandStart.Visible = False
    CommandStop.Visible = True
    CommandDone.Visible = False
    LabelCountdown.Visible = True
    ComboHours.Visible = False
    ComboMins.Visible = False
    ComboSecs.Visible = False
End Sub

Private Sub ShowDoneControls()
    TimerColorBlink.Enabled = True
    CommandStart.Visible = False
    CommandStop.Visible = False
    CommandDone.Visible = True
    LabelCountdown.Visible = True
    ComboHours.Visible = False
    ComboMins.Visible = False
    ComboSecs.Visible = False
End Sub
