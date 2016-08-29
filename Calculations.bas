Attribute VB_Name = "Calculations"
Option Explicit

Public Function CharExecution(pObject As Object, isclass As Boolean) As String
    CharExecution = ""
    Dim TLI         As TLIApplication
    Dim lInterface  As InterfaceInfo
    Dim lMember     As MemberInfo
    Dim ClassInfo As InterfaceInfo
    Dim FilteredMembers As SearchResults
    Dim FilteredItem As SearchItem
    Dim ClassName As String
    
    
    
    Set TLI = New TLIApplication
    Set lInterface = TLI.InterfaceInfoFromObject(pObject)

    Set ClassInfo = TLI.InterfaceInfoFromObject(pObject)
    Set FilteredMembers = ClassInfo.Members.GetFilteredMembers
    
    ClassName = Replace$(ClassInfo.Name, "_", "")
    
    For Each lMember In lInterface.Members
        If lMember.Name <> "winapi" Then

            
            CharExecution = CharExecution & ParseMember(lMember, IIf(isclass, ClassName, "")) & vbCrLf
        End If
    Next

    
    Set pObject = Nothing
    Set lInterface = Nothing
    Set TLI = Nothing
  End Function

   '================================================================================

Function GetVariableName(Index As Long) As String
    Select Case Index
        Case vbNull
            GetVariableName = " As Null"
        Case vbInteger
            GetVariableName = " As Integer"
        Case vbLong
            GetVariableName = " As Long"
        Case vbSingle
            GetVariableName = " As Single"
        Case vbDouble
            GetVariableName = " As Double"
        Case vbCurrency
            GetVariableName = " As Currency"
        Case vbDate
            GetVariableName = " As Date"
        Case vbString
            GetVariableName = " As String"
        Case vbObject
            GetVariableName = " As Object"
        Case vbError
            GetVariableName = " As Error"
        Case vbBoolean
            GetVariableName = " As Boolean"
        Case vbVariant
            GetVariableName = "" '" As Variant"
        Case vbDataObject
            GetVariableName = " As DataObject"
        Case vbDecimal
            GetVariableName = " As Decimal"
        Case vbByte
            GetVariableName = " As Byte"
        Case vbUserDefinedType
            GetVariableName = " As UserDefinedType"
        Case vbArray
            GetVariableName = " As Array"
        Case Else
            GetVariableName = " As UNKNOWN"
    End Select
End Function

Private Function ParseMember(lMember As MemberInfo, ClassName As String) As String

    Dim Parameters As String
    Dim ParameterInf As ParameterInfo
    
    For Each ParameterInf In lMember.Parameters
       Parameters = Parameters & IIf(Parameters <> "", ", ", "") & ParameterInf.Name & GetVariableName(ParameterInf.VarTypeInfo)
    Next
    
    If ClassName <> "" Then ClassName = ClassName & "."
 
    Select Case lMember.InvokeKind
        Case INVOKE_FUNC
            If lMember.ReturnType.VarType <> VT_VOID Then
                ParseMember = "Function " & ClassName & lMember.Name & "( " & Parameters & " )" & GetVariableName(lMember.ReturnType.VarType)
            Else
                ParseMember = "Sub " & ClassName & lMember.Name & "( " & Parameters & " )"
            End If
        Case INVOKE_PROPERTYGET
            ParseMember = "Property Get" & ClassName & lMember.Name & "( " & Parameters & " )" & GetVariableName(lMember.ReturnType.VarType)
        Case INVOKE_PROPERTYPUT
            ParseMember = "Property Let" & ClassName & lMember.Name & "( " & Parameters & " )"
        Case INVOKE_PROPERTYPUTREF
            ParseMember = "Property Set" & ClassName & lMember.Name & "( " & Parameters & " )"
        Case INVOKE_CONST
            ParseMember = "Const " & ClassName
        Case INVOKE_EVENTFUNC
            ParseMember = "Event " & ClassName & lMember.Name & "( " & Parameters & " )"
        Case Else
            ParseMember = ClassName & lMember.Name
    End Select
End Function

'Public Const PI As Double = 3.14159265358979

' argument in radians
Public Function Arccos(ByVal x As Double) As Double
   If Abs(x) <> 1 Then
       Arccos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    Else
       Arccos = IIf(x = 1, 0, Atn(1) * 4)
    End If
End Function

' argument in radians
Public Function Arcsin(ByVal x As Double) As Double
   If Abs(x) <> 1 Then
       Arcsin = Atn(x / Sqr(-x * x + 1))
    Else
       Arcsin = IIf(x = 1, Atn(1) * 2, Atn(1) * 6)
    End If
End Function

'usefull c function missing in vb. X & Y are triangle's cathets/
Public Function Atan2(ByVal y As Double, ByVal x As Double) As Double
   If x = 0 And y = 0 Then
      Atan2 = 0
   Else
      Atan2 = Atn(y / x) - PI * (x < 0)
   End If
End Function

'conversion RAD<->DEG
Public Function Rad(ByVal x As Double) As Double
  Rad = x * Atn(1) / 45#
End Function

Public Function Deg(ByVal x As Double) As Double
  Deg = x * 45# / Atn(1)
End Function

'Helpfull functions to compute sin, cos, tan
'with argument in degrees
Public Function Sind(ByVal x As Double) As Double
   Sind = Sin(Rad(x))
End Function

Public Function aSind(ByVal x As Double) As Double
   aSind = Deg(Arcsin(x))
End Function

Public Function Cosd(ByVal x As Double) As Double
   Cosd = Cos(Rad(x))
End Function

Public Function aCosd(ByVal x As Double) As Double
   aCosd = Deg(Arccos(x))
End Function

Public Function Tand(ByVal x As Double) As Double
   Tand = Tan(Rad(x))
End Function

Public Function aTand(ByVal x As Double) As Double
   aTand = Deg(Atn(x))
End Function

Public Function NormalizeAngle(ByVal x As Double) As Double
   Dim ret As Double
   ret = x - Int(x / 360#) * 360#
   If ret < 0 Then ret = ret + 360
   NormalizeAngle = ret
End Function
