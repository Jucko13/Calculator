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


Public Function GetTickCount() As Long: GetTickCount = GetTickCountA: End Function
Public Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As Long: FindWindow = FindWindowA(lpClassName, lpWindowName): End Function

Public Sub help(Optional Index As Long = 0)
    Select Case Index
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
    Dim i As Long
    
    For i = 0 To UBound(ExternalConstants)
        s = s & ExternalConstants(i) & "   "
        If (i + 1) Mod 4 = 0 Then
            s = s & vbCrLf
        End If
    Next i
    
    MsgBox s, vbInformation, "List of Constants and Keywords"
End Sub

Public Sub showcommands()
    Dim s As String
    Dim i As Long
    
    For i = 0 To UBound(ExternalFunctions)
        s = s & ExternalFunctions(i) & "   "
        If (i + 1) Mod 4 = 0 Then
            s = s & vbCrLf
        End If
    Next i
    
    MsgBox s, vbInformation, "List of Integrated VB Commands"
End Sub

Public Sub showoperators()
    Dim s As String
    Dim i As Long
    
    For i = 0 To UBound(ExternalOperators)
        s = s & ExternalOperators(i) & "   "
    Next i
    
    MsgBox s, vbInformation, "List of Operators"
End Sub


Public Sub showcustomcommands()
    Dim s As String
    Dim i As Long
    
    For i = 0 To UBound(ExternalCustomFunctions)
        s = s & ExternalCustomFunctions(i) & vbCrLf
    Next i
    
    MsgBox s, vbInformation, "List of Custom Commands"
    
End Sub

Public Sub initialize(ByRef com As commondialog)
    Set commondialog = com
End Sub


Public Function ColorPicker(Optional cc As Long = 0) As String
    commondialog.Color = cc
    commondialog.flags = &H1 + &H2
    
    commondialog.ShowColor
    
    ColorPicker = LongToRGB(commondialog.Color) & " : "
    
    Dim tmps As String
    tmps = Hex(commondialog.Color)
    If Len(tmps) < 6 Then
        tmps = String(6 - Len(tmps), "0") & tmps
    End If
    
    tmps = "#" & tmps
    
    ColorPicker = ColorPicker & tmps & " : " & commondialog.Color
End Function

Function LongToRGB(ByVal c As Long) As String
    Dim r, g, b
    
    r = c Mod &H100
    c = c \ &H100
    g = c Mod &H100
    c = c \ &H100
    b = c Mod &H100
    LongToRGB = "rgb(" & r & ", " & g & ", " & b & ")"
End Function
