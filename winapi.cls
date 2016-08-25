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

Private c As CommonDialog
Attribute c.VB_VarHelpID = -1

Public Function GetTickCount() As Long: GetTickCount = GetTickCountA: End Function
Public Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As Long: FindWindow = FindWindowA(lpClassName, lpWindowName): End Function


Public Sub initialize(ByRef com As CommonDialog)
    Set c = com
End Sub


Public Function ColorPicker(Optional cc As Long = 0) As String
    c.Color = cc
    c.flags = &H1 + &H2
    
    c.ShowColor
    
    ColorPicker = LongToRGB(c.Color) & " : "
    
    Dim tmps As String
    tmps = Hex(c.Color)
    If Len(tmps) < 6 Then
        tmps = String(6 - Len(tmps), "0") & tmps
    End If
    
    tmps = "#" & tmps
    
    ColorPicker = ColorPicker & tmps & " : " & c.Color
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
