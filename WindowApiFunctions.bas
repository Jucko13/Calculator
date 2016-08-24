Attribute VB_Name = "WindowApiFunctions"
Option Explicit

Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function defWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function GetCapture Lib "user32" () As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As Rect) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetRect Lib "user32" (ByRef lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook&, ByVal lpfn&, ByVal hmod&, ByVal dwThreadId&) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook&) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ()
Declare Sub CopyMemoryFromMinMaxInfo Lib "kernel32" Alias "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, ByVal cbCopy As Long)
Declare Sub CopyMemoryToMinMaxInfo Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long


Declare Function GetAsyncKeyState Lib "user32" (ByVal VKey As Long) As Integer
Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Global Const SPI_GETWORKAREA = 48
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Global Const LWA_ALPHA As Long = &H2

Global Const MOUSEEVENTF_LEFTDOWN = &H2  '  left button down /This is actually MOUSEEVENTF_LEFTDOWN
Global Const MOUSEEVENTF_LEFTUP = &H4  '  left button up /This is actually MOUSEEVENTF_LEFTUP
Global Const MOUSEEVENTF_RIGHTDOWN = &H8
Global Const MOUSEEVENTF_RIGHTUP = &H10
Global Const MOUSEEVENTF_MIDDLEDOWN = &H20
Global Const MOUSEEVENTF_MIDDLEUP = &H40
Global Const MOUSEEVENTF_WHEEL = &H800
Global Const MOUSEEVENTF_ABSOLUTE As Long = &H8000

Global Const VK_LBUTTON = &H1

Global Const vbEmptyErr As String = "01101981" * 1

Global Const GWL_EXSTYLE = -20
Global Const GWL_STYLE = (-16)
Global Const GWL_STYLING = (-16)
Global Const GWL_WNDPROC = (-4)

Global Const HC_ACTION = 0
Global Const HC_NOREMOVE = 3

Global Const HTBOTTOM = 15
Global Const HTBOTTOMLEFT = 16
Global Const HTBOTTOMRIGHT = 17
Global Const HTCAPTION = 2
Global Const HTLEFT = 10
Global Const HTTOP = 12
Global Const HTTOPLEFT = 13
Global Const HTTOPRIGHT = 14

Global Const SW_HIDE As Long = 0
Global Const SW_MAXIMIZE As Long = 3
Global Const SW_MINIMIZE As Long = 6
Global Const SW_RESTORE As Long = 9
Global Const SW_SHOW As Long = 5
Global Const SW_SHOWMAXIMIZED As Long = 3
Global Const SW_SHOWMINIMIZED As Long = 2
Global Const SW_SHOWMINNOACTIVE As Long = 7
Global Const SW_SHOWNA As Long = 8
Global Const SW_SHOWNOACTIVATE As Long = 4
Global Const SW_SHOWNORMAL As Long = 1

Global Const SWP_FRAMECHANGED = &H20
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOOWNERZORDER = &H200
Global Const SWP_NOSIZE = &H1
Global Const SWP_NOZORDER = &H4
Global Const SWP_SHOWWINDOW As Long = &H40

Enum CP
    [Black] = &H0
    [Green] = &HFF00&
    [Green Dark] = &H8000&
    [Green Light] = &H80FF80
    [Gray] = &HC0C0C0
    [Gray Dark] = &H404040
    [Gray Light] = &HE0E0E0
    [Blue] = &HFF0000
    [Blue Dark] = &H800000
    [Blue Light] = &HFF8080
    [Red] = &HFF&
    [Red Dark] = &H80&
    [Red Light] = &H8080FF
    [Cyan] = &HFFFF00
    [Cyan Dark] = &H808000
    [Cyan Light] = &HFFFF80
    [Yellow] = &HFFFF&
    [Yellow Dark] = &H8080&
    [Yellow Light] = &H80FFFF
End Enum

Global Const WS_BORDER = &H800000
Global Const WH_KEYBOARD_LL = 13
Global Const WH_MOUSE As Long = 14
Global Const WM_ACTIVATE As Integer = &H6
Global Const WM_ACTIVATEAPP As Integer = &H1C
Global Const WM_ASKCBFORMATNAME As Integer = &H30C
Global Const WM_CANCELJOURNAL As Integer = &H4B
Global Const WM_CANCELMODE As Integer = &H1F
Global Const WM_CHANGECBCHAIN As Integer = &H30D
Global Const WM_CHAR As Integer = &H102
Global Const WM_CHARTOITEM As Integer = &H2F
Global Const WM_CHILDACTIVATE As Integer = &H22
Global Const WM_CLEAR As Integer = &H303
Global Const WM_CLOSE As Integer = &H10
Global Const WM_COMMAND As Integer = &H111
Global Const WM_COMMNOTIFY As Integer = &H44
Global Const WM_COMPACTING As Integer = &H41
Global Const WM_COMPAREITEM As Integer = &H39
Global Const WM_CONTEXTMENU As Integer = &H7B
Global Const WM_COPY As Integer = &H301
Global Const WM_COPYDATA As Integer = &H4A
Global Const WM_CREATE As Integer = &H1
Global Const WM_CTLCOLORBTN As Integer = &H135
Global Const WM_CTLCOLORDLG As Integer = &H136
Global Const WM_CTLCOLOREDIT As Integer = &H133
Global Const WM_CTLCOLORLISTBOX As Integer = &H134
Global Const WM_CTLCOLORMSGBOX As Integer = &H132
Global Const WM_CTLCOLORSCROLLBAR As Integer = &H137
Global Const WM_CTLCOLORSTATIC As Integer = &H138
Global Const WM_CUT As Integer = &H300
Global Const WM_DEADCHAR As Integer = &H103
Global Const WM_DELETEITEM As Integer = &H2D
Global Const WM_DESTROY As Integer = &H2
Global Const WM_DESTROYCLIPBOARD As Integer = &H307
Global Const WM_DEVMODECHANGE As Integer = &H1B
Global Const WM_DISPLAYCHANGE As Integer = &H7E
Global Const WM_DRAWCLIPBOARD As Integer = &H308
Global Const WM_DRAWITEM As Integer = &H2B
Global Const WM_DROPFILES As Integer = &H233
Global Const WM_ENABLE As Integer = &HA
Global Const WM_ENDSESSION As Integer = &H16
Global Const WM_ENTERIDLE As Integer = &H121
Global Const WM_ENTERMENULOOP As Integer = &H211
Global Const WM_ENTERSIZEMOVE As Integer = &H231&
Global Const WM_ERASEBKGND As Integer = &H14
Global Const WM_EXITMENULOOP As Integer = &H212
Global Const WM_EXITSIZEMOVE As Integer = &H232&
Global Const WM_FONTCHANGE As Integer = &H1D
Global Const WM_GETDLGCODE As Integer = &H87
Global Const WM_GETFONT As Integer = &H31
Global Const WM_GETHOTKEY As Integer = &H33
Global Const WM_GETICON As Integer = &H7F
Global Const WM_GETMINMAXINFO As Integer = &H24
Global Const WM_GETOBJECT As Integer = &H3D
Global Const WM_GETTEXT As Integer = &HD
Global Const WM_GETTEXTLENGTH As Integer = &HE
Global Const WM_HELP As Integer = &H53
Global Const WM_HOTKEY As Integer = &H312
Global Const WM_HSCROLL As Integer = &H114
Global Const WM_HSCROLLCLIPBOARD As Integer = &H30E
Global Const WM_ICONERASEBKGND As Integer = &H27
Global Const WM_INITDIALOG As Integer = &H110
Global Const WM_INITMENU As Integer = &H116
Global Const WM_INITMENUPOPUP As Integer = &H117
Global Const WM_INPUTLANGCHANGE As Integer = &H51
Global Const WM_INPUTLANGCHANGEREQUEST As Integer = &H50
Global Const WM_KEYDOWN As Integer = &H100
Global Const WM_KEYFIRST As Integer = &H100
Global Const WM_KEYLAST As Integer = &H108
Global Const WM_KEYUP As Integer = &H101
Global Const WM_KILLFOCUS As Integer = &H8
Global Const WM_LBUTTONDBLCLK As Integer = &H203
Global Const WM_LBUTTONDOWN As Integer = &H201
Global Const WM_LBUTTONUP As Integer = &H202
Global Const WM_MBUTTONDBLCLK As Integer = &H209
Global Const WM_MBUTTONDOWN As Integer = &H207
Global Const WM_MBUTTONUP As Integer = &H208
Global Const WM_MDIACTIVATE As Integer = &H222
Global Const WM_MDICASCADE As Integer = &H227
Global Const WM_MDICREATE As Integer = &H220
Global Const WM_MDIDESTROY As Integer = &H221
Global Const WM_MDIGETACTIVE As Integer = &H229
Global Const WM_MDIICONARRANGE As Integer = &H228
Global Const WM_MDIMAXIMIZE As Integer = &H225
Global Const WM_MDINEXT As Integer = &H224
Global Const WM_MDIREFRESHMENU As Integer = &H234
Global Const WM_MDIRESTORE As Integer = &H223
Global Const WM_MDISETMENU As Integer = &H230
Global Const WM_MDITILE As Integer = &H226
Global Const WM_MEASUREITEM As Integer = &H2C
Global Const WM_MENUCHAR As Integer = &H120
Global Const WM_MENUSELECT As Integer = &H11F
Global Const WM_MOUSEACTIVATE As Integer = &H21
Global Const WM_MOUSEFIRST As Integer = &H200
Global Const WM_MOUSELAST As Integer = &H209
Global Const WM_MOUSEMOVE As Integer = &H200
Global Const WM_MOUSEWHEEL As Integer = &H20A
Global Const WM_MOVE As Integer = &H3
Global Const WM_NCACTIVATE As Integer = &H86
Global Const WM_NCCALCSIZE As Integer = &H83
Global Const WM_NCCREATE As Integer = &H81
Global Const WM_NCDESTROY As Integer = &H82
Global Const WM_NCHITTEST As Integer = &H84
Global Const WM_NCLBUTTONDBLCLK As Integer = &HA3
Global Const WM_NCLBUTTONDOWN As Integer = &HA1
Global Const WM_NCLBUTTONUP As Integer = &HA2
Global Const WM_NCMBUTTONDBLCLK As Integer = &HA9
Global Const WM_NCMBUTTONDOWN As Integer = &HA7
Global Const WM_NCMBUTTONUP As Integer = &HA8
Global Const WM_NCMOUSEMOVE As Integer = &HA0
Global Const WM_NCPAINT As Integer = &H85
Global Const WM_NCRBUTTONDBLCLK As Integer = &HA6
Global Const WM_NCRBUTTONDOWN As Integer = &HA4
Global Const WM_NCRBUTTONUP As Integer = &HA5
Global Const WM_NEXTDLGCTL As Integer = &H28
Global Const WM_NOTIFY As Integer = &H4E
Global Const WM_NOTIFYFORMAT As Integer = &H55
Global Const WM_NULL As Integer = &H0
Global Const WM_OTHERWINDOWCREATED As Integer = &H42
Global Const WM_OTHERWINDOWDESTROYED As Integer = &H43
Global Const WM_PAINT As Integer = &HF
Global Const WM_PAINTCLIPBOARD As Integer = &H309
Global Const WM_PAINTICON As Integer = &H26
Global Const WM_PALETTECHANGED As Integer = &H311
Global Const WM_PALETTEISCHANGING As Integer = &H310
Global Const WM_PARENTNOTIFY As Integer = &H210
Global Const WM_PASTE As Integer = &H302
Global Const WM_PENWINFIRST As Integer = &H380
Global Const WM_PENWINLAST As Integer = &H38F
Global Const WM_POWER As Integer = &H48
Global Const WM_PRINT As Integer = &H317
Global Const WM_PRINTCLIENT As Integer = &H318
Global Const WM_QUERYDRAGICON As Integer = &H37
Global Const WM_QUERYENDSESSION As Integer = &H11
Global Const WM_QUERYNEWPALETTE As Integer = &H30F
Global Const WM_QUERYOPEN As Integer = &H13
Global Const WM_QUEUESYNC As Integer = &H23
Global Const WM_QUIT As Integer = &H12
Global Const WM_RBUTTONDBLCLK As Integer = &H206
Global Const WM_RBUTTONDOWN As Integer = &H204
Global Const WM_RBUTTONUP As Integer = &H205
Global Const WM_RENDERALLFORMATS As Integer = &H306
Global Const WM_RENDERFORMAT As Integer = &H305
Global Const WM_SETCURSOR As Integer = &H20
Global Const WM_SETFOCUS As Integer = &H7
Global Const WM_SETFONT As Integer = &H30
Global Const WM_SETHOTKEY As Integer = &H32
Global Const WM_SETICON As Integer = &H80
Global Const WM_SETREDRAW As Integer = &HB
Global Const WM_SETTEXT As Integer = &HC
Global Const WM_SHOWWINDOW As Integer = &H18
Global Const WM_SIZE As Integer = &H5
Global Const WM_SIZECLIPBOARD As Integer = &H30B
Global Const WM_SPOOLERSTATUS As Integer = &H2A
Global Const WM_STYLECHANGED As Integer = &H7D
Global Const WM_STYLECHANGING As Integer = &H7C
Global Const WM_SYNCPAINT As Integer = &H88
Global Const WM_SYSCHAR As Integer = &H106
Global Const WM_SYSCOLORCHANGE As Integer = &H15
Global Const WM_SYSCOMMAND As Integer = &H112
Global Const WM_SYSDEADCHAR As Integer = &H107
Global Const WM_SYSKEYDOWN As Integer = &H104
Global Const WM_SYSKEYUP As Integer = &H105
Global Const WM_TCARD As Integer = &H52
Global Const WM_TIMECHANGE As Integer = &H1E
Global Const WM_TIMER As Integer = &H113
Global Const WM_UNDO As Integer = &H304
Global Const WM_USER As Integer = &H400
Global Const WM_USERCHANGED As Integer = &H54
Global Const WM_VKEYTOITEM As Integer = &H2E
Global Const WM_VSCROLL As Integer = &H115
Global Const WM_VSCROLLCLIPBOARD As Integer = &H30A
Global Const WM_WINDOWPOSCHANGED As Integer = &H47
Global Const WM_WINDOWPOSCHANGING As Integer = &H46
Global Const WM_WININICHANGE As Integer = &H1A

Global Const WS_DLGFRAME = &H400000
Global Const WS_EX_LAYERED As Long = &H80000
Global Const WS_EX_STATICEDGE = &H20000
Global Const WS_THICKFRAME = &H40000

Global Const LLKHF_ALTDOWN As Integer = &H20

Type POINTAPI
    X       As Long
    Y       As Long
End Type

Type Point
    X       As Long
    Y       As Long
End Type

Type Rect
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Type CWPSTRUCT
    lParam As Long
    wParam As Long
    Message As Long
    hWnd As Long
End Type

Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    Time As Long
    dwExtraInfo As Long
End Type

Type MSLLHOOKSTRUCT
    pt As Point
    mouseData As Integer
    flags As Integer
    Time As Integer
    dwExtraInfo As Integer
End Type

Type MINMAXINFO
    ptReserved      As POINTAPI
    ptMaxSize       As POINTAPI
    ptMaxPosition   As POINTAPI
    ptMinTrackSize  As POINTAPI
    ptMaxTrackSize  As POINTAPI
End Type

Type FormSize
    Width As Long
    Height As Long
    Top As Long
    Left As Long
End Type

Enum UpdateMode
    [UpdateAll] = 2
    [UpdatePos] = 0
    [UpdateWindowHandle] = 1
    [UseNewSearch] = 3
    [WindowExist] = 4
End Enum

Enum WhatMode
    [Char - Return Chars] = 1
    [Char - Return Empty] = 3
    [Number - Return Empty] = 2
    [Number - Return Numbers] = 0
End Enum

Type WINDOWPLACEMENT
    flags               As Long
    Length              As Long
    ptMaxPosition       As POINTAPI
    ptMinPosition       As POINTAPI
    rcNormalPosition    As Rect
    showCmd             As Long
End Type

Global Const HWND_TOPMOST       As Long = -1
Global Const HWND_NOTOPMOST     As Long = -2
Global Const NotDesingTime      As Boolean = True
Global Const KeyUp1             As Boolean = False
Global Const KeyDown1           As Boolean = True
Global Const CTTT1              As String = "|" & "BTW|1|" & "btw|2|" & "n1|3|" & "np|4|" & "N1|5|" & "NP|6|" & "WTF|7|" & "wtf|8|"
Global Const CTTT2              As String = "|1|By The Way" & "|2|by the way" & "|3|Nice One!" & "|4|No Problem" & "|5|Nice One!" & "|6|No Problem" & "|" & "|7|What The Fack!" & "|8|what the fack!" & "|"
Global Const WHEEL_DELTA        As Long = 120

Global ChangeTheTextTo1         As String
Global ChangeTheTextTo2         As String
Global FormLoadingOnStartUp     As Boolean
Global gHW                      As Long
Global HideOnMinimize           As Boolean
Global IsPause                  As Boolean
Global IsPlaying                As Boolean
Global kbd_Hook                 As Long
Global ListDirty                As Boolean
Global lpPrevWndProc            As Long
Global lSave                    As String
Global MayAutoComplete          As Boolean
Global mouse_Hook               As Long
Global mouse_Hook2              As Long
Global NewFocusWindow           As String
Global OldFocusWindow           As String
Global RecWKey                  As Boolean
Global RecDelayTime             As Boolean
Global RecKeyStrokes            As Boolean
Global RecMouseClicks           As Boolean
Global RecMouseMove             As Boolean
Global RecOnTop                 As Boolean
Global Recording                As Boolean
Global RecShowToolTip           As Boolean
Global RecStickPointer          As Boolean
Global RecWgfDelay              As Boolean
Global WindowHndl               As Long
Global WindowHndlHeight         As Long
Global WindowHndlLeft           As Long
Global WindowHndlPos            As Rect
Global WindowHndlTop            As Long
Global WindowHndlWidth          As Long
Global WindowString()           As String

Global infoBarVisible           As Boolean

Global IsMaximized              As Boolean
Global Size                     As FormSize

Dim old_hwnd                    As Long
Dim m_PrevProc                  As Long
Dim m_SecProc                   As Long
Dim m_PrevProc2                 As Long

Global IsSizing                 As Boolean

Private mGetVisible As Boolean

'Public Function GetTheWindowsDirectory() As String
'
'    Dim strWindowsDir As String
'    Dim lngWindowsDirLength As Long
'
'    strWindowsDir = Space(250)
'    lngWindowsDirLength = GetWindowsDirectory(strWindowsDir, 250)
'    strWindowsDir = Left(strWindowsDir, lngWindowsDirLength)
'    GetTheWindowsDirectory = strWindowsDir
'End Function

Property Let SetVisible(Value1 As Boolean)
    mGetVisible = Value1
End Property

Property Get GetVisible() As Boolean
    GetVisible = mGetVisible
End Property

Function getFocusWindow()
    Dim foreground_hwnd As Long
    Dim txt As String
    Dim Length As Long
    
    foreground_hwnd = GetForegroundWindow()
    WindowHndl = foreground_hwnd
    txt = Space$(1024)
    Length = GetWindowText(foreground_hwnd, txt, Len(txt))
    txt = Left$(txt, Length)
    getFocusWindow = txt
    DoEvents
End Function

Sub LoadDataIntoFile(DataName As Integer, FileName As String)
    Dim myArray() As Byte
    Dim myFile As Long
    If Dir(FileName) = "" Then
        myArray = LoadResData(DataName, "CUSTOM")
        myFile = FreeFile
        Open FileName For Binary Access Write As #myFile
        Put #myFile, , myArray
        Close #myFile
    End If
End Sub

Public Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim sSave As String
    Dim i As Long
    On Error GoTo Endiii:
    sSave = Space$(GetWindowTextLength(hWnd) + 1)
    GetWindowText hWnd, sSave, Len(sSave)
    sSave = Left$(sSave, Len(sSave) - 1)
    If sSave <> "" Then lSave = lSave & sSave & vbCrLf
Endiii:
    EnumChildProc = 1
End Function

Sub SetFocusToWindow(strCaption As String, Optional strClassName As String = vbNullString)
Dim hWinodw As Long
    hWinodw = FindWindow(strClassName, strCaption)
    
    If hWinodw <> 0 Then
        SetForegroundWindow hWinodw
    End If
End Sub

Sub SetTrans(ByVal OB_Hwnd As Long, ByVal OB_Val As Integer)
   Dim Attrib As Long
   Attrib = GetWindowLong(OB_Hwnd, GWL_EXSTYLE)
   SetWindowLong OB_Hwnd, GWL_EXSTYLE, Attrib Or WS_EX_LAYERED
   SetLayeredWindowAttributes OB_Hwnd, RGB(0, 255, 0), OB_Val, LWA_ALPHA
End Sub

Sub SetTheWindowPosAndShow(strCaption As String, QHndl As Long, X As Long, Y As Long, X1 As Long, Y1 As Long)
On Error Resume Next
Dim RetVal As Long
Dim hWindow As Long
    hWindow = FindWindow(vbNullString, strCaption)
    
    If hWindow = QHndl Then
        RetVal = ShowWindow(QHndl, SW_RESTORE)
        SetForegroundWindow QHndl
        SetWindowPos QHndl, 0, X, Y, X1, Y1, 0
    ElseIf hWindow <> QHndl And hWindow <> 0 Then
        RetVal = ShowWindow(hWindow, SW_RESTORE)
        SetForegroundWindow hWindow
        SetWindowPos hWindow, 0, X, Y, X1, Y1, 0
    ElseIf hWindow = 0 Then
        ShowMsgBox "Window not found!"
    End If
End Sub

Function GetRectPosOfWindow()
    getFocusWindow
    GetWindowRect WindowHndl, WindowHndlPos
    GetRectPosOfWindow = WindowHndlPos.Left & ", " & WindowHndlPos.Top & ", " & WindowHndlPos.Right & ", " & WindowHndlPos.Bottom
End Function

Function OnlyWhat(Str As String, Optional mode As WhatMode = [Number - Return Numbers], Optional SilenceMode As Boolean = False) As String
    Dim iStr As Integer
    Dim sStr As String
    Dim TotalStr As String
    Dim RubishStr As String
    
    If Str = "" Then OnlyWhat = "": Exit Function
    
    If mode = [Char - Return Chars] Or mode = [Char - Return Empty] Then
    ''------------------------------------------------
        For iStr = 0 To Len(Str)
            sStr = Mid(Str, iStr + 1, 1)
            If InStr(1, AlphaBet, sStr) > 0 Then
                TotalStr = TotalStr & sStr
            Else
                RubishStr = RubishStr & sStr
            End If
        Next iStr
        
        If Len(RubishStr) > 0 Then
            If SilenceMode = False Then ShowMsgBox ("These chars are not allowed: " & "''" & RubishStr & "''" & vbCrLf & "Only Alfabetic Chars!" & vbCrLf & "These are allowed: " & AlphaBet)
            If mode = [Char - Return Empty] Then OnlyWhat = "" Else OnlyWhat = TotalStr
        Else
            OnlyWhat = TotalStr
        End If
    ''------------------------------------------------
    ElseIf mode = [Number - Return Numbers] Or mode = [Number - Return Empty] Then
    ''------------------------------------------------
        For iStr = 0 To Len(Str)
            sStr = Mid(Str, iStr + 1, 1)
            If InStr(1, Numbers, sStr) > 0 Then
                TotalStr = TotalStr & sStr
            Else
                RubishStr = RubishStr & sStr
            End If
        Next iStr
        If Len(RubishStr) > 0 Then
            If SilenceMode = False Then ShowMsgBox ("These chars are not allowed: " & "''" & RubishStr & "''" & vbCrLf & "Only Numbers are allowed!" & vbCrLf & vbCrLf & "These are allowed: " & vbCrLf & "''" & Numbers & "''")
            If mode = [Number - Return Empty] Then OnlyWhat = "" Else OnlyWhat = TotalStr
        Else
            OnlyWhat = TotalStr
        End If
    ''------------------------------------------------
    End If
End Function


Function IsCheckProcess(sAppName As String) As Boolean
 Const strComputer = "."
 Dim objWMIService
 Dim colProcesses
 
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where Name = '" & sAppName & "'")
    If colProcesses.Count = 0 Then
        IsCheckProcess = False
    Else
        IsCheckProcess = True
    End If
End Function



