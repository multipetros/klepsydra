VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Klepsydra"
   ClientHeight    =   1320
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   2880
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandResume 
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      Picture         =   "main.frx":7532
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton CommandPause 
      Height          =   375
      Left            =   1560
      Picture         =   "main.frx":7591
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer TimerColorBlink 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   2280
      Top             =   720
   End
   Begin VB.ComboBox ComboMins 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.ComboBox ComboSecs 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.ComboBox ComboHours 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton CommandStart 
      Height          =   375
      Left            =   840
      Picture         =   "main.frx":75E3
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer TimerCountdown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   720
   End
   Begin VB.CommandButton CommandStop 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   840
      Picture         =   "main.frx":7635
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton CommandDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   840
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
      Left            =   360
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
      Begin VB.Menu MenuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuShutdown 
         Caption         =   "S&hutdown PC"
         Shortcut        =   ^H
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
' Copyright (c) 2017-2018, Petros Kyladitis <www.multipetros.gr>
'
' Klepsydra is a Countdown timer prgram with sound alarm
' It's open source, distributed under the GNU GPL3

Option Explicit

Private WithEvents TaskBarProgress As ITaskBarList3
Attribute TaskBarProgress.VB_VarHelpID = -1
' Task Bar Progress consts
Private Const TBPF_NOPROGRESS = 0
Private Const TBPF_INDETERMINATE = 1
Private Const TBPF_NORMAL = 2
Private Const TBPF_ERROR = 4
Private Const TBPF_PAUSED = 8

' program file names
Private Const FILENAME_ALARM = "alarm.wav"
Private Const FILENAME_INI = "klepsydra.ini"
Private Const FILENAME_FONT = "digital7.ttf"
Private Const FILENAME_LICENSE = "license.txt"

Private Const INI_SECT_MAIN = "main"
Private Const INI_HOURS = "hours"
Private Const INI_MINS = "mins"
Private Const INI_SECS = "secs"
Private Const INI_LOOP = "loop"
Private Const INI_ALARM = "alarm"
Private Const INI_MUTE = "mute"
Private Const INI_SHUTDOWN = "shutdown"
Private Const INI_LANGUAGE = "language"
Private Const INI_TOP = "top"
Private Const INI_LEFT = "left"

Private Const LANGUAGE_EN = "en"
Private Const LANGUAGE_EL = "el"
Private Const LANGUAGE_EN_ID = 100
Private Const LANGUAGE_EL_ID = 200

Private Const COUNTDOWN_FONT_NAME = "Digital-7"
Private Const COUNTDOWN_FONT_SIZE = 24

Dim Countdown As Date
Dim CountdownSecs As Integer
Dim AlarmFile As String
Dim langID As Integer
Dim iniPath As String

Private Sub Form_Load()
    Dim i, iniRNum As Integer
    Dim iniRNumL As Long
    Dim iniR As String
    
    iniPath = SpecialFolder(feUserAppData) & "\" & FILENAME_INI
    
    For i = 0 To 60
        If i < 24 Then
            ComboHours.AddItem (Format(i, "00"))
        End If
        ComboSecs.AddItem (Format(i, "00"))
        ComboMins.AddItem (Format(i, "00"))
    Next i
    
    iniR = IniRead(INI_SECT_MAIN, INI_HOURS, iniPath)
    iniRNum = CInt(IIf(IsNumeric(iniR), iniR, 0))
    If iniRNum > -1 And iniRNum < 25 Then
        ComboHours.ListIndex = iniRNum
    End If
    
    iniR = IniRead(INI_SECT_MAIN, INI_MINS, iniPath)
    iniRNum = CInt(IIf(IsNumeric(iniR), iniR, 0))
    If iniRNum > -1 And iniRNum < 61 Then
        ComboMins.ListIndex = iniRNum
    End If
    
    iniR = IniRead(INI_SECT_MAIN, INI_SECS, iniPath)
    iniRNum = CInt(IIf(IsNumeric(iniR), iniR, 10))
    If iniRNum > -1 And iniRNum < 61 Then
        ComboSecs.ListIndex = iniRNum
    Else
        ComboSecs.ListIndex = 5
    End If
    
    iniR = IniRead(INI_SECT_MAIN, INI_LOOP, iniPath)
    If iniR = "False" Then
        MenuLoop.Checked = False
    Else
        MenuLoop.Checked = True
    End If
    
    iniR = IniRead(INI_SECT_MAIN, INI_ALARM, iniPath)
    If iniR = "" Then
        AlarmFile = FILENAME_ALARM
    Else
        AlarmFile = iniR
    End If
    
    iniR = IniRead(INI_SECT_MAIN, INI_MUTE, iniPath)
    If iniR = "True" Then
        MenuMute.Checked = True
    Else
        MenuMute.Checked = False
    End If
    
    iniR = IniRead(INI_SECT_MAIN, INI_SHUTDOWN, iniPath)
    If iniR = "True" Then
        MenuShutdown.Checked = True
    Else
        MenuShutdown.Checked = False
    End If
    
    iniR = IniRead(INI_SECT_MAIN, INI_LANGUAGE, iniPath)
    If iniR = LANGUAGE_EL Then
        langID = LANGUAGE_EL_ID
        LoadStrings
        MenuGreek.Checked = True
        MenuEnglish.Checked = False
    Else
        langID = LANGUAGE_EN_ID
        LoadStrings
        MenuGreek.Checked = False
        MenuEnglish.Checked = True
    End If
    
    If LoadFont(FILENAME_FONT) > 0 Then
        LabelCountdown.Font.Name = COUNTDOWN_FONT_NAME
    End If
    
    LabelCountdown.Font.Size = COUNTDOWN_FONT_SIZE
    
    iniR = IniRead(INI_SECT_MAIN, INI_LEFT, iniPath)
    iniRNumL = CLng(IIf(IsNumeric(iniR), iniR, 0))
    If iniRNumL > 0 Then
        Me.Left = iniRNumL
    End If
    
    iniR = IniRead(INI_SECT_MAIN, INI_TOP, iniPath)
    iniRNumL = CLng(IIf(IsNumeric(iniR), iniR, 0))
    If iniRNumL > 0 Then
        Me.Top = iniRNumL
    End If
        
    ShowStartControls
    Set TaskBarProgress = New ITaskBarList3
End Sub
Private Sub LoadStrings()
    MenuFile.Caption = LoadResString(langID + 1)
    MenuExit.Caption = LoadResString(langID + 2)
    MenuLoop.Caption = LoadResString(langID + 3)
    MenuAlarmSound.Caption = LoadResString(langID + 4)
    MenuHelp.Caption = LoadResString(langID + 5)
    MenuAbout.Caption = LoadResString(langID + 6)
    MenuLicense.Caption = LoadResString(langID + 7)
    CommandStart.ToolTipText = LoadResString(langID + 8)
    CommandStop.ToolTipText = LoadResString(langID + 9)
    CommandDone.Caption = LoadResString(langID + 10)
    MenuLanguage.Caption = LoadResString(langID + 11)
    MenuEnglish.Caption = LoadResString(langID + 12)
    MenuGreek.Caption = LoadResString(langID + 13)
    Me.Caption = LoadResString(langID + 18)
    MenuMute.Caption = LoadResString(langID + 19)
    CommandPause.ToolTipText = LoadResString(langID + 20)
    CommandResume.ToolTipText = LoadResString(langID + 21)
    MenuShutdown.Caption = LoadResString(langID + 22)
End Sub

Private Sub CommandDone_Click()
    Dim wnd As Long
    PlaySound vbNullString, vbNull, 0
    wnd = SetTopMostWindow(Me.hwnd, False)
    TaskBarProgress.SetProgressState hwnd, TBPF_NOPROGRESS
    ShowStartControls
End Sub

Private Sub CommandResume_Click()
    TimerCountdown.Enabled = True
    CommandStop.Enabled = True
    CommandPause.Visible = True
    CommandResume.Visible = False
End Sub

Private Sub CommandPause_Click()
    TimerCountdown.Enabled = False
    CommandStop.Enabled = False
    CommandPause.Visible = False
    CommandResume.Visible = True
    TaskBarProgress.SetProgressState hwnd, TBPF_PAUSED
End Sub

Private Sub CommandStop_Click()
    TimerCountdown.Enabled = False
    TaskBarProgress.SetProgressState hwnd, TBPF_NOPROGRESS
    ShowStartControls
End Sub

Private Sub CommandStart_Click()
    Countdown = ComboHours.List(ComboHours.ListIndex) & ":" _
        & ComboMins.List(ComboMins.ListIndex) & ":" _
        & ComboSecs.List(ComboSecs.ListIndex)
    CountdownSecs = TimeValue(Countdown) * 86400
    TimerCountdown.Enabled = True
    ShowStopControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim iniW As Boolean
    iniW = IniWrite(INI_SECT_MAIN, INI_HOURS, ComboHours.Text, iniPath)
    iniW = IniWrite(INI_SECT_MAIN, INI_MINS, ComboMins.Text, iniPath)
    iniW = IniWrite(INI_SECT_MAIN, INI_SECS, ComboSecs.Text, iniPath)
    iniW = IniWrite(INI_SECT_MAIN, INI_LOOP, MenuLoop.Checked, iniPath)
    iniW = IniWrite(INI_SECT_MAIN, INI_ALARM, AlarmFile, iniPath)
    iniW = IniWrite(INI_SECT_MAIN, INI_MUTE, MenuMute.Checked, iniPath)
    iniW = IniWrite(INI_SECT_MAIN, INI_SHUTDOWN, MenuShutdown.Checked, iniPath)
    iniW = IniWrite(INI_SECT_MAIN, INI_TOP, CStr(Me.Top), iniPath)
    iniW = IniWrite(INI_SECT_MAIN, INI_LEFT, CStr(Me.Left), iniPath)
    Dim langStr As String
    If langID = LANGUAGE_EL_ID Then
        langStr = LANGUAGE_EL
    Else
        langStr = LANGUAGE_EN
    End If
    iniW = IniWrite(INI_SECT_MAIN, INI_LANGUAGE, langStr, iniPath)
    'release font resource
    UnloadFont (FILENAME_FONT)
    'stop music if playing
    PlaySound vbNullString, vbNull, 0
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
    langID = LANGUAGE_EN_ID
    LoadStrings
End Sub

Private Sub MenuExit_Click()
    Unload Me
End Sub

Private Sub MenuGreek_Click()
    MenuGreek.Checked = True
    MenuEnglish.Checked = False
    langID = LANGUAGE_EL_ID
    LoadStrings
End Sub

Private Sub MenuLicense_Click()
    Dim LicFile As String
    Dim TaskID As Double
    LicFile = App.path & "\" & FILENAME_LICENSE
    If FileExists(LicFile) = True Then
        On Error GoTo ErrMsgNotepad
        TaskID = Shell("notepad.exe " & LicFile, vbNormalFocus)
        Exit Sub
    End If
ErrMsgNotepad:
    Dim browse As Long
    If MsgBox(LoadResString(langID + 16), vbCritical + vbYesNo, LoadResString(langID + 17)) = vbYes Then
        browse = OpenBrowser("http://www.gnu.org/licenses/gpl.html")
    End If
End Sub

Private Sub MenuLoop_Click()
    MenuLoop.Checked = Not MenuLoop.Checked
End Sub

Private Sub MenuMute_Click()
    MenuMute.Checked = Not MenuMute.Checked
End Sub

Private Sub MenuShutdown_Click()
    MenuShutdown.Checked = Not MenuShutdown.Checked
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
    Dim LeftSecs As Integer
    Dim Snd, SndParams, wnd As Long
    
    Countdown = Countdown - (1 / 24 / 60 / 60)
    
    TimeStr = Format(Countdown, "hh:mm:ss")
    LabelCountdown.Caption = TimeStr
    Me.Caption = TimeStr & " - " & LoadResString(langID + 18)
    
    LeftSecs = CountdownSecs - TimeValue(Countdown) * 86400
    TaskBarProgress.SetProgressState hwnd, TBPF_NORMAL
    TaskBarProgress.SetProgressValue hwnd, LeftSecs, CountdownSecs
    
    If TimeStr = "00:00:00" Then
        TimerCountdown.Enabled = False
        If MenuLoop.Checked Then
            SndParams = SND_FILENAME Or SND_ASYNC Or SND_LOOP
        Else
            SndParams = SND_FILENAME Or SND_ASYNC
        End If
        If MenuMute.Checked = False Then
            PlaySound AlarmFile, vbNull, SndParams
        End If
        Me.SetFocus
        wnd = SetTopMostWindow(Me.hwnd, True)
        ShowDoneControls
        If MenuShutdown.Checked = True Then
            On Error GoTo ErrMsgShutdown
            Shell "shutdown -s -f -t 30", vbHide
            If MsgBox(LoadResString(langID + 23), vbQuestion & vbYesNo, LoadResString(langID + 22)) = vbYes Then
                Shell "shutdown -a", vbHide
            End If
        End If
    End If
    Exit Sub
ErrMsgShutdown:
    MsgBox LoadResString(langID + 24), vbCritical, LoadResString(langID + 25)
End Sub

Private Sub ShowStartControls()
    TimerColorBlink.Enabled = False
    CommandStart.Visible = True
    CommandStop.Visible = False
    CommandDone.Visible = False
    CommandPause.Visible = False
    CommandResume.Visible = False
    LabelCountdown.Visible = False
    ComboHours.Visible = True
    ComboMins.Visible = True
    ComboSecs.Visible = True
    Me.Caption = LoadResString(langID + 18)
End Sub

Private Sub ShowStopControls()
    LabelCountdown.ForeColor = vbButtonText
    LabelCountdown.BackColor = vbButtonFace
    CommandStart.Visible = False
    CommandStop.Visible = True
    CommandDone.Visible = False
    CommandPause.Visible = True
    CommandResume.Visible = False
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
    CommandPause.Visible = False
    CommandResume.Visible = False
    LabelCountdown.Visible = True
    ComboHours.Visible = False
    ComboMins.Visible = False
    ComboSecs.Visible = False
End Sub
