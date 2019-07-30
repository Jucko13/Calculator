VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "winapi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Function GetTickCountA Lib "kernel32" Alias "GetTickCount" () As Long


Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public commondialog As commondialog
Attribute commondialog.VB_VarHelpID = -1





'Private Declare Sub RtlMoveMemory Lib "kernel32" (dest As Any, src As Any, ByVal L As Long)
'Private Declare Function CallWindowProcA Lib "user32" (ByVal addr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
'Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
'Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
'
'Function CallApiByName(ByVal sLib As String, ByVal sMod As String, ParamArray Params() As Variant) As Long
'On Error Resume Next
'    Dim lPtr                As Long
'    Dim bvASM(&HEC00& - 1)  As Byte
'    Dim I                   As Long
'    Dim lMod                As Long
'
'    lMod = GetProcAddress(LoadLibraryA(sLib), sMod)
'    If lMod = 0 Then Exit Function
'
'    lPtr = VarPtr(bvASM(0))
'    RtlMoveMemory ByVal lPtr, &H59595958, &H4:              lPtr = lPtr + 4
'    RtlMoveMemory ByVal lPtr, &H5059, &H2:                  lPtr = lPtr + 2
'
'
'
'    For I = UBound(Params) To 0 Step -1
'    Debug.Print VarType(Params(I)(0))
'
'        RtlMoveMemory ByVal lPtr, &H68, &H1:                lPtr = lPtr + 1
'        RtlMoveMemory ByVal lPtr, CLng(Params(I)), &H4:     lPtr = lPtr + 4
'    Next
'    RtlMoveMemory ByVal lPtr, &HE8, &H1:                    lPtr = lPtr + 1
'    RtlMoveMemory ByVal lPtr, lMod - lPtr - 4, &H4:         lPtr = lPtr + 4
'    RtlMoveMemory ByVal lPtr, &HC3, &H1:                    lPtr = lPtr + 1
'    CallApiByName = CallWindowProcA(VarPtr(bvASM(0)), 0, 0, 0, 0)
'
'End Function
'
'Public Function ParamArrayDelegated(ParamArray prms() As Variant) As Variant
'Dim arrPrms() As Variant, arrWrk() As Variant
''When prms(0) is Array, supposed is delegated from another function
'  arrPrms = prms
'  Do While VarType(arrPrms(0)) >= vbArray And UBound(arrPrms) < 1
'    arrWrk = arrPrms(0)
'    arrPrms = arrWrk
'  Loop
'  ParamArrayDelegated = arrPrms
'End Function
'
'Function CallApiByName2(ByVal sLib As String, ByVal sMod As String, ParamArray Params() As Variant) As Long
' CallApiByName2 = CallApiByName(sLib, sMod, Params(0))
'End Function

Public Function GetClipboardText() As String
    GetClipboardText = Clipboard.GetText
End Function

Public Sub SetClipboardText(str As String)
    Clipboard.Clear
    Clipboard.SetText str
End Sub

Public Function GetTickCount() As Long
    GetTickCount = GetTickCountA()
End Function


Public Function NewApiCall(strDllName As String, strApiName As String, lngNumParameters As Long) As ApiCall
    Set NewApiCall = New ApiCall
    NewApiCall.Init strDllName, strApiName, lngNumParameters
End Function


Public Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As Long: FindWindow = FindWindowA(lpClassName, lpWindowName): End Function

Public Sub help(Optional index As Long = 0)
    Select Case index
        Case 0
            MsgBox "Gebruik winapi.help(0 tot 2) voor het vinden van alle keywords en functions." & vbCrLf & _
            vbCrLf & "1 - List of Custom Commands (functionlist.txt)" & _
            vbCrLf & "2 - List of Integrated VB Commands" & _
            vbCrLf & "3 - List of Constants and Keywords" & _
            vbCrLf & "4 - List of Operators", vbInformation
        
        Case 1
            showcustomcommands
            
        Case 2
            showcommands
            
        Case 3
            showkeywords
            
        Case 4
            showoperators
            
            
            
    End Select
End Sub

Public Sub showkeywords()
    Dim s As String
    Dim I As Long
    
    For I = 0 To UBound(ExternalConstants)
        s = s & ExternalConstants(I) & "   "
        If (I + 1) Mod 4 = 0 Then
            s = s & vbCrLf
        End If
    Next I
    
    MsgBox s, vbInformation, "List of Constants and Keywords"
End Sub

Public Sub showcommands()
    Dim s As String
    Dim I As Long
    
    For I = 0 To UBound(ExternalFunctions)
        s = s & ExternalFunctions(I) & "   "
        If (I + 1) Mod 4 = 0 Then
            s = s & vbCrLf
        End If
    Next I
    
    MsgBox s, vbInformation, "List of Integrated VB Commands"
End Sub

Public Sub showoperators()
    Dim s As String
    Dim I As Long
    
    For I = 0 To UBound(ExternalOperators)
        s = s & ExternalOperators(I) & "   "
    Next I
    
    MsgBox s, vbInformation, "List of Operators"
End Sub


Public Sub showcustomcommands()
    Dim s As String
    Dim I As Long
    
    For I = 0 To UBound(ExternalCustomFunctions)
        s = s & ExternalCustomFunctions(I) & vbCrLf
    Next I
    
    MsgBox s, vbInformation, "List of Custom Commands"
    
End Sub

Public Sub initialize(ByRef com As commondialog)
    Set commondialog = com
End Sub


Public Function ColorPicker(Optional cc As Variant) As String
    Dim openColor As Long
    
    Select Case VarType(cc)
    Case 8
        If InStr(1, cc, "#") > 0 Then openColor = HexToColor(Replace(cc, "#", ""))
        
    Case 2, 3, 4, 5, 6, 17
        openColor = CLng(cc)
    Case Else
        openColor = 0
        
    End Select
    
    commondialog.Color = openColor
    commondialog.flags = &H1 + &H2
    
    commondialog.ShowColor
    
    ColorPicker = LongToRGB(commondialog.Color) & " : " & commondialog.Color & " : &H" & Hex(commondialog.Color)
End Function

Function LongToRGB(ByVal c As Long) As String
    Dim r As Long, g As Long, b As Long
    
    r = c Mod &H100
    c = c \ &H100
    g = c Mod &H100
    c = c \ &H100
    b = c Mod &H100
    
    LongToRGB = "rgb(" & r & ", " & g & ", " & b & ") : #" & IIf(r < 16, "0", "") & Hex(r) & IIf(g < 16, "0", "") & Hex(g) & IIf(b < 16, "0", "") & Hex(b)
End Function

Function HexToColor(ByRef HexColor As String) As Long
    ' variable size byte array
    Dim bytHex() As Byte
    ' we only accept one length, 6 characters = 12 bytes
    If LenB(HexColor) = 12 Then
        ' convert string to byte array
        bytHex = HexColor
        ' if a value is now higher than 57, we reduce it by 7
        If bytHex(0) > &H39 Then bytHex(0) = bytHex(0) - 7
        If bytHex(2) > &H39 Then bytHex(2) = bytHex(2) - 7
        If bytHex(4) > &H39 Then bytHex(4) = bytHex(4) - 7
        If bytHex(6) > &H39 Then bytHex(6) = bytHex(6) - 7
        If bytHex(8) > &H39 Then bytHex(8) = bytHex(8) - 7
        If bytHex(10) > &H39 Then bytHex(10) = bytHex(10) - 7
        ' this function is "stupid", it assumes it gets correct data...
        '  makes it faster, but you can give it any string that is 6 characters long, no error, ever!
        '  we take 4 bits for each six characters, and place it in the correct position of a Long,
        '  making up 24 bits that are required to represent a color value
        HexToColor = ((bytHex(0) And &HF&) * &H10&) Or (bytHex(2) And &HF&) _
            Or ((bytHex(4) And &HF&) * &H1000&) Or ((bytHex(6) And &HF&) * &H100&) _
            Or ((bytHex(8) And &HF&) * &H100000) Or ((bytHex(10) And &HF&) * &H10000)
    End If
End Function

Function GetProperties(obj As Variant, isClass As Boolean) As String
    GetProperties = ObjectProperties(obj, isClass)
End Function
