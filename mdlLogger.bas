Attribute VB_Name = "mdlLogger"
Option Explicit


Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetCapture Lib "user32" () As Long

Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFilename As String) As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long

Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook&, ByVal lpfn&, ByVal hmod&, ByVal dwThreadId&) As Long
Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook&) As Long

Global Const PI As String = "3.14159265358979323846264338327950288419716939937510582097494459230781640628620899862803482534211706798214808651328230664709384460955058223172535940812848111745028410270193852110555964462294895493038196442881097566593344612847564823378678316527120190914564856692346034861045432664821339360726024914127372458700660631558817488152092096282925409171536436789259036001133053054882046652138414695194151160943305727036575959195309218611738193261179310511854807446237996274956735188575272489122793818301194912983367336244065664308602139494639522473719070217986094370277053921717629317675238467481846766940513200056812714526356082778577134275778960917363717872146844090122495343014654958537105079227968925892354201995611212902196086403441815981362977477130996051870721134999999837297804995105973173281609631859502445945534690830264252230825334468503526193118817101000313783875288658753320838142061717766914730359825349042875546873115956286388235378759375195778185778053217122680661300192787661119"

   Declare Function RegCloseKey Lib "advapi32.dll" _
   (ByVal hKey As Long) As Long
   Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
   "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
   As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
   As Long, phkResult As Long, lpdwDisposition As Long) As Long
   Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
   "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
   Long) As Long
   Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As String, lpcbData As Long) As Long
   Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, lpData As _
   Long, lpcbData As Long) As Long
   Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As Long, lpcbData As Long) As Long
   Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
   String, ByVal cbData As Long) As Long
   Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
   ByVal cbData As Long) As Long

Global Const HKEY_LOCAL_MACHINE As Long = &H80000002
Global Const REG_SZ As Long = 1

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Const VK_LBUTTON = &H1
Const VK_RBUTTON = &H2
Const VK_1BUTTON = &H31
Const VK_2BUTTON = &H32
Const VK_3BUTTON = &H33
Const VK_7BUTTON = &H37

Const WM_KEYDOWN As Integer = &H100
Const WM_KEYFIRST As Integer = &H100
Const WM_KEYLAST As Integer = &H108
Const WM_KEYUP As Integer = &H101

Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    Time As Long
    dwExtraInfo As Long
End Type

Global Calculating As Boolean
Global HowMany As Integer

Private hHook As Long
Private IsHooked As Boolean
Private kb_struct As KBDLLHOOKSTRUCT

Global Calculated As Boolean
Global TempStr As String
Global PressedCalc As String
Global MayLog As Boolean
Global TypedText As String

Global Calculation As String
Global objScript As ScriptControl


Dim Ystr As String

Private bAlt                As Boolean
Private bControl            As Boolean
Private bEscape             As Boolean
Const bLog                    As Boolean = True
Private bShift              As Boolean
Private bWindows            As Boolean
Private bW                  As Boolean

Private sText               As String
Private EditingIsBusy       As Boolean
Private ShiftDown           As Boolean
Private AltDown             As Boolean
Private ControlDown         As Boolean
Private EscapeDown          As Boolean
Private WindowsDown         As Boolean
Private WDown               As Boolean



Public Const REG_DWORD As Long = 4

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_USERS = &H80000003

Global Const WH_KEYBOARD_LL = 13
Global Const HC_ACTION = 0
Global Const HC_NOREMOVE = 3

Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_ARENA_TRASHED = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259

Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_ALL_ACCESS = &H3F

Public Const REG_OPTION_NON_VOLATILE = 0

Function GetFileContent(strFileName As String) As String
    Dim iFile As Integer
    iFile = FreeFile()
    
    Open strFileName For Input As #iFile
        GetFileContent = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
    Close #iFile
End Function

Function SetValueEx(ByVal hKey As Long, sValueName As String, _
lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, _
                                           lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, _
lType, lValue, 4)
        End Select
End Function

Sub SetKeyValue(sKeyName As String, sValueName As String, _
vValueSetting As Variant, lValueType As Long)
    Dim lRetVal As Long         'result of the SetValueEx function
    Dim hKey As Long         'handle of open key

    'open the specified key
    lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, _
                              KEY_SET_VALUE, hKey)
    lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)
End Sub

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

Public Sub SetKeyboardHook()
Dim i As Integer

    If IsHooked Then
'        ShowMsgbox "Don't hook WH_KEYBOARD_LL twice or you will be unable to unhook it."
    Else
        hHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
        IsHooked = True
    End If
End Sub

Public Sub RemoveKeyboardHook()
    UnhookWindowsHookEx hHook
    IsHooked = False
End Sub

Public Function LowLevelKeyboardProc(ByVal uCode As Long, ByVal wParam As Long, lParam As KBDLLHOOKSTRUCT) As Long
    Dim cns As String
    Dim cnt As Integer
    Static KeyCounter As Integer
    If uCode >= 0 And uCode = HC_ACTION Then
    If EditingIsBusy = True Then GoTo LATED:
    If Calculating = True Then GoTo LATED:
        Select Case lParam.vkCode
    
'            Case &HA4
'                If wParam = WM_KEYUP And bLog Then
'                    LowLevelKeyboardProc = CallNextHookEx(hHook, uCode, wParam, lParam)
'                    Exit Function
'                End If



            Case &HA0, &HA1 'Shift
                bShift = wParam = WM_KEYDOWN
                If bShift = True Then
                    ShiftDown = True
                Else
                    ShiftDown = False
                End If
                DoEvents

            Case VK_ESCAPE ' EscapeKey
                bEscape = wParam = WM_KEYDOWN
                If bEscape = True Then
                    If EscapeDown = False Then
                            MayLog = False
                    End If
                    EscapeDown = True
                Else
                    EscapeDown = False
                End If
                DoEvents

            Case VK_LWIN, VK_RWIN ' WindowsKey
                bWindows = wParam = WM_KEYDOWN
                If bWindows = True Then
                    WindowsDown = True
                Else
                    WindowsDown = False
                End If
                DoEvents
        End Select

        Select Case lParam.vkCode

            Case vbKeyA To vbKeyZ
                If bLog Then
                        If ShiftDown Then
                            If wParam = WM_KEYUP Then
                                If MayLog = True Then
                                    TypedText = TypedText & Chr$(lParam.vkCode)
                                End If
                                
                            End If
    
                        ElseIf bAlt Then

                        ElseIf bControl Then

                        Else
                            If wParam = WM_KEYUP Then
                                If MayLog = True Then
                                    TypedText = TypedText & Chr$((lParam.vkCode + 32))
                                End If
                            End If
                        End If

                End If


            Case vbKey0 To vbKey9
                If bLog Then
                    If wParam = WM_KEYUP And ShiftDown = True Then
                        If lParam.vkCode = vbKey9 Then MayLog = True
                        If MayLog = True Then
                            Select Case lParam.vkCode
                                Case vbKey0
                                    cns = ")"
                                Case vbKey1
                                    cns = "!"
                                Case vbKey2
                                    cns = "@"
                                Case vbKey3
                                    cns = "#"
                                Case vbKey4
                                    cns = "$"
                                Case vbKey5
                                    cns = "%"
                                Case vbKey6
                                    cns = "^"
                                Case vbKey7
                                    cns = "&"
                                Case vbKey8
                                    cns = "*"
                                Case vbKey9
                                    cns = "("
                            End Select
                            TypedText = TypedText & cns
                        End If
                    ElseIf wParam = WM_KEYUP Then
                        If MayLog = True Then
                            TypedText = TypedText & (lParam.vkCode - 48) & ""
                        End If
                    End If
                End If

            Case 96 To 105
                If bLog Then
                    If wParam = WM_KEYUP Then
                        If MayLog = True Then
                            TypedText = TypedText & (lParam.vkCode - 96) & ""
                        End If
                    End If
                End If


            Case 187 To 191
                If bLog And bAlt Then
                    Select Case lParam.vkCode
                        Case 187
                            cnt = lParam.vkCode - 126
                        Case Else
                            cnt = lParam.vkCode - 144
                    End Select
                    If MayLog = True Then
                        TypedText = TypedText & Chr$(cnt)
                    End If
                ElseIf bLog Then
                    If wParam = WM_KEYUP Then
                        Select Case lParam.vkCode
                            Case 187
                                cnt = lParam.vkCode - 126
                            Case Else
                                cnt = lParam.vkCode - 144
                        End Select
                        If MayLog = True Then
                            TypedText = TypedText & Chr$(cnt)
                        End If
                    End If
                End If

            Case VK_SPACE
                If bLog = True And wParam = WM_KEYUP Then
                    If MayLog = True Then
                        TypedText = TypedText & " "
                    End If
                End If

            Case VK_RETURN
                If bLog = True And wParam = WM_KEYDOWN And ShiftDown = True Then
                    If MayLog = True Then
                        'Do While GetAsyncKeyState(vbKeyShift): Loop
                        'Do While ShiftDown = True: Loop
                        If Len(TypedText) > 0 Then
                            With Form1
                                MayLog = False
                                .Text1.Text = TypedText
                                .Text2.Text = .CheckCalculation(.Text1.Text)
                                Ystr = .Text2.Text
                                If Ystr = "Syntax Error" Then
                                    SendKeys "=" & Ystr
                                Else
                                    SendKeys ("{backspace " & Len(TypedText) & "}")
                                    SendKeys Ystr
                                    .Text2.Text = Ystr
                                    .Text1.Text = TypedText
                                    Ystr = ""
                                    TypedText = ""
                                End If
                            End With
                        End If
                        LowLevelKeyboardProc = -1
                        Exit Function
                    End If
                    
                ElseIf bLog = True And wParam = WM_KEYUP And ShiftDown = True Then
                    LowLevelKeyboardProc = -1
                    Exit Function
                End If
                    

            Case VK_BACK
                If bLog = True And wParam = WM_KEYUP Then
                    If MayLog = True Then
                        If Len(TypedText) > 0 Then
                            TypedText = Mid(TypedText, 1, Len(TypedText) - 1)
                        End If
                    End If

                End If

            Case VK_DECIMAL
                If bLog = True And wParam = WM_KEYUP Then
                    If MayLog = True Then
                        TypedText = TypedText & "*"
                    End If
                End If

            Case 96 To 105
                If bLog = True And wParam = WM_KEYUP Then
                    If MayLog = True Then
                        TypedText = TypedText & Chr(lParam.vkCode - 48) & ""
                    End If
                End If

            Case 106, 107, 108, 109, 111
                If bLog = True And wParam = WM_KEYUP Then
                    If MayLog = True Then
                        If lParam.vkCode = 106 Then
                            TypedText = TypedText & "*"
                        ElseIf lParam.vkCode = 107 Then
                            TypedText = TypedText & "+"
                        ElseIf lParam.vkCode = 108 Then
                            TypedText = TypedText & "/"
                        ElseIf lParam.vkCode = 109 Then
                            TypedText = TypedText & "-"
                        ElseIf lParam.vkCode = 111 Then
                            TypedText = TypedText & "/"
                        End If

                    End If
                End If
            
            Case VK_ESCAPE
                If bLog = True And wParam = WM_KEYUP Then
                    TypedText = ""
                    Ystr = ""
                End If
        End Select
End If

LATED:
    LowLevelKeyboardProc = CallNextHookEx(hHook, uCode, wParam, lParam)

Exit Function
PressedEnter:
    LowLevelKeyboardProc = -1
End Function



'Function Check_For_Sin(Str As String) As String
'Dim midTmp As Double
'Dim midTmp123 As String
'Dim MidEndstr As String
'
'MidEndstr = Str
'        Do
'            HowMany = HowMany + 1
'            midTmp123 = midString(MidEndstr, "asin(", ")", HowMany)
'            midTmp = objScript.Eval(midTmp123)
'            If Len(midTmp123) <> 0 Then MidEndstr = Replace(MidEndstr, ("asin(" & midTmp123 & ")"), Round(Asin(midTmp), 15), , , vbTextCompare)
'        Loop Until (midString(MidEndstr, "asin(", ")", HowMany)) = 0
'HowMany = HowMany - 1
'        Do
'            HowMany = HowMany + 1
'            midTmp123 = midString(MidEndstr, "sin(", ")", HowMany)
'            midTmp = objScript.Eval(midTmp123)
'            If Len(midTmp123) <> 0 Then MidEndstr = Replace(MidEndstr, ("sin(" & midTmp123 & ")"), Round(Sin(midTmp123 * Pi / 180), 15), , , vbTextCompare)
'        Loop Until (midString(MidEndstr, "sin(", ")", HowMany)) = 0
'HowMany = HowMany - 1
'Check_For_Sin = MidEndstr
'End Function
'
'
'
'Function Check_For_Cos(Str As String) As String
'Dim midTmp As Double
'Dim midTmp123 As String
'Dim MidEndstr As String
'
'MidEndstr = Str
'        Do
'            HowMany = HowMany + 1
'            midTmp123 = midString(MidEndstr, "acos(", ")", HowMany)
'            midTmp = objScript.Eval(midTmp123)
'            If Len(midTmp123) <> 0 Then MidEndstr = Replace(MidEndstr, "acos(" & midTmp123 & ")", Round((aAcos(objScript.Eval(midTmp)) / Pi) * 180, 15), , , vbTextCompare)
'        Loop Until (midString(MidEndstr, "acos(", ")", HowMany)) = 0
'HowMany = HowMany - 1
'        Do
'            HowMany = HowMany + 1
'            midTmp123 = midString(MidEndstr, "cos(", ")", HowMany)
'            midTmp = objScript.Eval(midTmp123)
'            If Len(midTmp123) <> 0 Then MidEndstr = Replace(MidEndstr, ("cos(" & midTmp123 & ")"), Round(Cos((objScript.Eval(midTmp) * Pi) / 180), 15), , , vbTextCompare)
'        Loop Until (midString(MidEndstr, "cos(", ")", HowMany)) = 0
'HowMany = HowMany - 1
'Check_For_Cos = MidEndstr
'End Function
'
'
'Function Check_For_Tan(Str As String) As String
'Dim midTmp As Double
'Dim midTmp123 As String
'Dim MidEndstr As String
'
'MidEndstr = Str
'        Do
'            HowMany = HowMany + 1
'            midTmp123 = midString(MidEndstr, "atan(", ")", HowMany)
'            midTmp = objScript.Eval(midTmp123)
'            If Len(midTmp123) <> 0 Then MidEndstr = ReplaceRev(MidEndstr, "atan(" & midTmp123 & ")", Round((Atn(midTmp) / Pi) * 180, 15))
'        Loop Until (midString(MidEndstr, "atan(", ")", HowMany)) = 0
'HowMany = HowMany - 1
'        Do
'        HowMany = HowMany + 1
'            midTmp123 = midString(MidEndstr, "tan(", ")", HowMany)
'            midTmp = objScript.Eval(midTmp123)
'            If Len(midTmp123) <> 0 Then MidEndstr = ReplaceRev(MidEndstr, "tan(" & midTmp123 & ")", Round(Tan((midTmp / 180) * Pi), 15))
'        Loop Until (midString(MidEndstr, "tan(", ")", HowMany)) = 0
'HowMany = HowMany - 1
'
'Check_For_Tan = MidEndstr
'End Function

Function ReplaceRev(ExString As String, Find1 As String, Replace1 As String)
Dim Place1 As Integer
Dim Place2 As Integer
Dim TempStr As String

Place1 = InStrRev(ExString, Find1, , vbTextCompare)

TempStr = Replace(ExString, Find1, Replace1, , 1, vbTextCompare)
ReplaceRev = TempStr
End Function



Function Asin(X As Double) As Double
Dim ix As Double
Dim tmpINSAS As Double
ix = X

    tmpINSAS = Sqr((-ix * ix + 1))
    If tmpINSAS = 0 Then
        Asin = (Atn(ix) * 180) / PI * 2
    Else
        Asin = (Atn(ix / tmpINSAS) * 180) / PI
    End If
    'Asin = (X * 180) / Pi
End Function

Function aAcos(X As Double) As Double
'Dim ix As Double
'ix = x '(x * Pi) / 180

    aAcos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function
