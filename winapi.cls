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


Public Function GetTickCount() As Long: GetTickCount = GetTickCountA: End Function
Public Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As Long: FindWindow = FindWindowA(lpClassName, lpWindowName): End Function


