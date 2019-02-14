VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub RtlMoveMemory Lib "kernel32" (dest As Any, src As Any, ByVal L As Long)
Private Declare Function CallWindowProcA Lib "user32" (ByVal addr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long


Dim dllName As String
Dim apiName As String
Dim numParameters As Long
Dim moduleHandle As Long
Dim isInitialized As Boolean
Dim parameters() As Variant

Private Sub Class_Initialize()
    isInitialized = False
    
End Sub


Sub Init(strDllName As String, strApiName As String, lngNumParameters As Long)
    dllName = strDllName
    apiName = strApiName
    numParameters = lngNumParameters
    
    moduleHandle = GetProcAddress(LoadLibraryA(strDllName), strApiName)
    If moduleHandle = 0 Then
        Err.Raise 100, "ApiCall", "Could not dynamically load the API " & strApiName & " from library " & strDllName & "."
        Exit Sub
    End If
    
    Erase parameters
    
    isInitialized = True
End Sub

'Function CallApi(ParamArray Params() As Variant)
'    'For i = 0 To UBound(Params)
'    '    If VarType(Params) = vbString Then
'    '        Params(i) = StrPtr(Params(i))
'    '    Else
'    '        Params(i) = VarPtr(Params(i))
'    '    End If
'    'Next i
'
'    CallApi = CallApiByName(Params(0))
'End Function

Public Function P(parameter As Variant) As ApiCall
    Set P = Me
    
    Dim newIndex As Long
    
    If IsArrayInitialized(parameters) Then
        newIndex = UBound(parameters) + 1
    Else
        newIndex = 0
    End If
    
    ReDim Preserve parameters(0 To newIndex)
    parameters(newIndex) = parameter
End Function

Public Function IsArrayInitialized(arr) As Boolean

  Dim rv As Long

  On Error Resume Next

  rv = UBound(arr)
  IsArrayInitialized = (Err.Number = 0)

End Function

Public Function E() As Long
    If IsArrayInitialized(parameters) Then
        E = ExecuteApi(parameters)
        Erase parameters
    Else
        E = 0
    End If
End Function

'Public Function CallByParameters(ParamArray Params() As Variant) As Long
'    CallByParameters = ExecuteApi(Params(0))
'End Function




Public Function ExecuteApi(Params() As Variant) As Long
    
    Dim lPtr                As Long
    Dim bvASM(&HEC00& - 1)  As Byte
    Dim i                   As Long
    
    ExecuteApi = 0
    
    If Not isInitialized Then
        Err.Raise 100, "ApiCall", "Class is not initialized. Call 'Init(strDllName, strApiName, lngNumParameters) to initialize the class."
        Exit Function
    End If
    
    
    If UBound(Params) <> numParameters - 1 Then
        Err.Raise 101, "ApiCall", "The number of parameters should be " & numParameters & " instead of " & (UBound(Params) + 1)
        
        Exit Function
    End If
    
    On Error Resume Next
    lPtr = VarPtr(bvASM(0))
    RtlMoveMemory ByVal lPtr, &H59595958, &H4:              lPtr = lPtr + 4
    RtlMoveMemory ByVal lPtr, &H5059, &H2:                  lPtr = lPtr + 2
    For i = UBound(Params) To 0 Step -1
        RtlMoveMemory ByVal lPtr, &H68, &H1:                lPtr = lPtr + 1
        
        If VarType(Params(i)) = vbString Then
            RtlMoveMemory ByVal lPtr, CLng(StrPtr(Params(i))), &H4:     lPtr = lPtr + 4
        Else
            RtlMoveMemory ByVal lPtr, CLng(Params(i)), &H4:     lPtr = lPtr + 4
        End If
    Next
    RtlMoveMemory ByVal lPtr, &HE8, &H1:                    lPtr = lPtr + 1
    RtlMoveMemory ByVal lPtr, moduleHandle - lPtr - 4, &H4:         lPtr = lPtr + 4
    RtlMoveMemory ByVal lPtr, &HC3, &H1:                    lPtr = lPtr + 1
    
    ExecuteApi = CallWindowProcA(VarPtr(bvASM(0)), 0, 0, 0, 0)

End Function
