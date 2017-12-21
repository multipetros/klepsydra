VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITaskBarList3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_COMMAND As Long = &H111
Private Const THBN_CLICKED As Long = &H1800
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const AsmMain As String = "558BEC83C4FC8D45FC50FF7514FF7510FF750CFF75086800000000B800000000FFD08B45FCC9C21000"
Private ASMArr() As Byte
Private oldproc As Long, wnd As Long

Private Const CLSID_TaskbarList As String = "{56FDF344-FD6D-11d0-958A-006097C9A090}"
Private Const IID_ITaskbarList3 As String = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"
Private Enum ITaskbarList3Members
                                '/* ITaskbarList methods */
    HrInit_ = 3                 'STDMETHOD( HrInit )( THIS ) PURE;
    AddTab_ = 4                 'STDMETHOD( AddTab )( THIS_ HWND ) PURE;
    DeleteTab_ = 5              'STDMETHOD( DeleteTab )( THIS_ HWND ) PURE;
    ActivateTab_ = 6            'STDMETHOD( ActivateTab )( THIS_ HWND ) PURE;
    SetActiveAlt_ = 7           'STDMETHOD( SetActiveAlt )( THIS_ HWND ) PURE;
                                '/* ITaskbarList2 methods */
    MarkFullscreenWindow_ = 8   'STDMETHOD( MarkFullscreenWindow )( THIS_ HWND, BOOL ) PURE;
                                '/* ITaskbarList3 methods */
    SetProgressValue_ = 9       'STDMETHOD( SetProgressValue )( THIS_ HWND, ULONGLONG, ULONGLONG ) PURE;
    SetProgressState_ = 10      'STDMETHOD( SetProgressState )( THIS_ HWND, TBPFLAG ) PURE;
    RegisterTab_ = 11           'STDMETHOD( RegisterTab )( THIS_ HWND, HWND ) PURE;
    UnregisterTab_ = 12         'STDMETHOD( UnregisterTab )( THIS_ HWND ) PURE;
    SetTabOrder_ = 13           'STDMETHOD( SetTabOrder )( THIS_ HWND, HWND ) PURE;
    SetTabActive_ = 14          'STDMETHOD( SetTabActive )( THIS_ HWND, HWND, DWORD ) PURE;
    ThumbBarAddButtons_ = 15    'STDMETHOD( ThumbBarAddButtons )( THIS_ HWND, UINT, LPTHUMBBUTTON ) PURE;
    ThumbBarUpdateButtons_ = 16 'STDMETHOD( ThumbBarUpdateButtons )( THIS_ HWND, UINT, LPTHUMBBUTTON ) PURE;
    ThumbBarSetImageList_ = 17  'STDMETHOD( ThumbBarSetImageList )( THIS_ HWND, HIMAGELIST ) PURE;
    SetOverlayIcon_ = 18        'STDMETHOD( SetOverlayIcon )( THIS_ HWND, HICON, LPCWSTR ) PURE;
    SetThumbnailTooltip_ = 19   'STDMETHOD( SetThumbnailTooltip )( THIS_ HWND, LPCWSTR ) PURE;
    SetThumbnailClip_ = 20      'STDMETHOD( SetThumbnailClip )( THIS_ HWND, RECT * ) PURE;
'                                '/* ITaskbarList4 methods */
'    SetTabProperties_ = 21      'STDMETHOD( SetTabProperties )( THIS_ HWND, STPFLAG ) PURE;
End Enum
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Enum THUMBBUTTONMASK
    THB_BITMAP = 1
    THB_ICON = 2
    THB_TOOLTIP = 4
    THB_FLAGS = 8
End Enum
Private Enum THUMBBUTTONFLAGS
    THBF_ENABLED = 0
    THBF_DISABLED = 1
    THBF_DISMISSONCLICK = 2
    THBF_NOBACKGROUND = 4
    THBF_HIDDEN = 8
    THBF_NONINTERACTIVE = 16
End Enum
Private Type THUMBBUTTON
    dwMask As THUMBBUTTONMASK
    iId As Long
    iBitmap As Long
    hIcon As Long
    szTip As String * 260
    dwFlags As THUMBBUTTONFLAGS
End Type
Private Obj As Long

Public Event ButtonPressed(ByVal Index As Integer)

Public Function ButtonsWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
    Case WM_COMMAND:
        'HiWord of wParam
        If wParam \ &H10000 = THBN_CLICKED Then _
            RaiseEvent ButtonPressed(wParam And &HFFFF&)
    End Select
    ButtonsWindowProc = CallWindowProc(oldproc, hwnd, uMsg, wParam, lParam)
End Function

Private Sub Class_Initialize()
    Dim CLSID As Guid, InterfaceGuid As Guid
    Call CLSIDFromString(StrConv(CLSID_TaskbarList, vbUnicode), CLSID)
    Call IIDFromString(StrConv(IID_ITaskbarList3, vbUnicode), InterfaceGuid)
    Call CoCreateInstance(CLSID, 0, 1, InterfaceGuid, Obj)
End Sub

Private Sub Class_Terminate()
    If wnd Then StopSubclass wnd, oldproc
    Call CallInterface(Obj, unk_Release, 0)
End Sub

Public Function SetProgressState(ByVal hwnd As Long, ByVal tbpFlags As Long) As Long
    SetProgressState = CallInterface(Obj, SetProgressState_, 2, hwnd, tbpFlags)
End Function

Public Function SetProgressValue(ByVal hwnd As Long, ByVal ullCompleted As Long, ByVal ullTotal As Long) As Long
    SetProgressValue = CallInterface(Obj, SetProgressValue_, 5, hwnd, ullCompleted, 0, ullTotal, 0)
End Function

Public Function SetOverlayIcon(ByVal hwnd As Long, ByVal hIcon As Long, Optional ByRef pszDescription As String) As Long
    SetOverlayIcon = CallInterface(Obj, SetOverlayIcon_, 3, hwnd, hIcon, StrPtr(pszDescription))
End Function

Public Function SetThumbnailClip(ByVal hwnd As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Long
    Dim rc As RECT
    rc.Left = Left
    rc.Top = Top
    rc.Right = Right
    rc.Bottom = Bottom
    SetThumbnailClip = CallInterface(Obj, SetThumbnailClip_, 2, hwnd, VarPtr(rc))
End Function

Public Function SetThumbnailTooltip(ByVal hwnd As Long, ByRef pszTip As String) As Long
    SetThumbnailTooltip = CallInterface(Obj, SetThumbnailTooltip_, 2, hwnd, StrPtr(pszTip))
End Function

Public Function ThumbBarAddButtons(ByVal hwnd As Long, ByVal cButtons As Long, ByRef icons() As Long) As Long
    Dim i As Integer, icn(6) As Long
    If cButtons < 1 Or cButtons > 7 Or wnd Then Exit Function
    
    ReDim buttons(cButtons - 1) As THUMBBUTTON
    For i = 0 To UUBound(icons)
        icn(i) = icons(i)
    Next i
    For i = 0 To cButtons - 1
        buttons(i).dwMask = THB_ICON
        buttons(i).iId = i
        buttons(i).hIcon = icn(i)
    Next i
    ThumbBarAddButtons = CallInterface(Obj, ThumbBarAddButtons_, 3, hwnd, cButtons, VarPtr(buttons(0)))
    
    wnd = hwnd
    StartSubclass ASMArr, wnd, oldproc
End Function

Private Sub StartSubclass(ByRef ASM() As Byte, ByVal hwnd As Long, ByRef OldWndProc As Long, Optional ByVal ProcNumber As Long)                 ' Сабклассинг с пом. ASM (автора не знаю...)
    Dim lng As Long, tPtr As Long
    lng = Len(AsmMain) \ 2&
    ReDim ASM(0 To lng - 1)
    For lng = 0 To lng - 1
        ASM(lng) = Val("&H" & Mid$(AsmMain, (lng) * 2& + 1, 2&))
    Next lng
    Call CopyMemory(tPtr, ByVal ObjPtr(Me), 4&)
    Call CopyMemory(lng, ByVal tPtr + &H1C + (4& * ProcNumber), 4&)
    Call CopyMemory(ASM(23), ObjPtr(Me), 4&)
    Call CopyMemory(ASM(28), lng, 4&)
    OldWndProc = SetWindowLong(hwnd, &HFFFC, VarPtr(ASM(0)))
End Sub

Private Sub StopSubclass(ByVal hwnd As Long, ByVal OldWndProc As Long)
    Call SetWindowLong(hwnd, &HFFFC, OldWndProc)
End Sub

Private Function UUBound(arr) As Long
On Error GoTo 1:
    UUBound = UBound(arr)
Exit Function
1:  UUBound = -1
End Function
