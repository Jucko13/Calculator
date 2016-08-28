Attribute VB_Name = "Calculations"
Option Explicit

Public Function CharExecution(pObject As Object) As String
    CharExecution = ""
    Dim TLI         As TLIApplication
    Dim lInterface  As InterfaceInfo
    Dim lMember     As MemberInfo
    Dim ClassInfo As Object
    Dim FilteredMembers As Object
    
    Set TLI = New TLIApplication
    Set lInterface = TLI.InterfaceInfoFromObject(pObject)

    Set ClassInfo = TLI.InterfaceInfoFromObject(pObject)
    Set FilteredMembers = ClassInfo.Members.GetFilteredMembers

    For Each lMember In lInterface.Members
        'If WhatIsIt(lMember) = "Property Get" Then
            'CharExecution = CharExecution & "*****" & lMember.Name & " : " & TLI.InvokeHook(pObject, lMember.Name, INVOKE_PROPERTYGET)
        'End If
        CharExecution = CharExecution & lMember.Name & " : " & WhatIsIt(lMember) & vbCrLf
    Next
    Set pObject = Nothing
    Set lInterface = Nothing
    Set TLI = Nothing
  End Function

   '================================================================================

 Private Function WhatIsIt(lMember As MemberInfo) As String
  Select Case lMember.InvokeKind
    Case INVOKE_FUNC
        If lMember.ReturnType.VarType <> VT_VOID Then
            WhatIsIt = "Function"
        Else
            WhatIsIt = "Method"
        End If
    Case INVOKE_PROPERTYGET
        WhatIsIt = "Property Get"
    Case INVOKE_PROPERTYPUT
        WhatIsIt = "Property Let"
    Case INVOKE_PROPERTYPUTREF
        WhatIsIt = "Property Set"
    Case INVOKE_CONST
        WhatIsIt = "Const"
    Case INVOKE_EVENTFUNC
        WhatIsIt = "Event"
    Case Else
        
        
        WhatIsIt = lMember.InvokeKind & " (Unknown) " & TypeName(lMember)
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
