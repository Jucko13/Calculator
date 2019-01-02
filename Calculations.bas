Attribute VB_Name = "Calculations"
Option Explicit

Public Sub MergeSort(ByRef pvarArray As Variant, Optional pvarMirror As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngMid As Long
    Dim L As Long
    Dim r As Long
    Dim O As Long
    Dim varSwap As Variant
 
    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
        ReDim pvarMirror(plngLeft To plngRight)
    End If
    lngMid = plngRight - plngLeft
    Select Case lngMid
        Case 0
        Case 1
            If pvarArray(plngLeft) > pvarArray(plngRight) Then
                varSwap = pvarArray(plngLeft)
                pvarArray(plngLeft) = pvarArray(plngRight)
                pvarArray(plngRight) = varSwap
            End If
        Case Else
            lngMid = lngMid \ 2 + plngLeft
            MergeSort pvarArray, pvarMirror, plngLeft, lngMid
            MergeSort pvarArray, pvarMirror, lngMid + 1, plngRight
            ' Merge the resulting halves
            L = plngLeft ' start of first (left) half
            r = lngMid + 1 ' start of second (right) half
            O = plngLeft ' start of output (mirror array)
            Do
                If pvarArray(r) < pvarArray(L) Then
                    pvarMirror(O) = pvarArray(r)
                    r = r + 1
                    If r > plngRight Then
                        For L = L To lngMid
                            O = O + 1
                            pvarMirror(O) = pvarArray(L)
                        Next
                        Exit Do
                    End If
                Else
                    pvarMirror(O) = pvarArray(L)
                    L = L + 1
                    If L > lngMid Then
                        For r = r To plngRight
                            O = O + 1
                            pvarMirror(O) = pvarArray(r)
                        Next
                        Exit Do
                    End If
                End If
                O = O + 1
            Loop
            For O = plngLeft To plngRight
                pvarArray(O) = pvarMirror(O)
            Next
    End Select
End Sub


Public Function CharExecution(pObject As Object, isclass As Boolean) As String
    CharExecution = ""
    Dim TLI         As TLIApplication
    Dim lInterface  As InterfaceInfo
    Dim lMember     As MemberInfo
    'Dim ClassInfo As TypeInfo
    Dim FilteredMembers As SearchResults
    Dim FilteredItem As SearchItem
    Dim ClassName As String
    
    
    
    Set TLI = New TLIApplication
    Set lInterface = TLI.InterfaceInfoFromObject(pObject)

    'Set ClassInfo = TLI.ClassInfoFromObject(pObject)
    
    Set FilteredMembers = lInterface.Members.GetFilteredMembers
    
'    Dim i As Long
'    Dim s As SearchItem
'
'    For i = 1 To FilteredMembers.Count
'        Set s = FilteredMembers.Item(i)
'
'        Debug.Print FilteredMembers.Item(i).Name; " "; s.Constant; " "; s.Hidden
'    Next i
    
    
    
    
    ClassName = Replace$(lInterface.Name, "_", "")
    
    For Each lMember In lInterface.Members
        'If lMember.Name <> "winapi" Then
            CharExecution = CharExecution & ParseMember(lMember, IIf(isclass, ClassName, "")) & vbCrLf
        'End If
    Next
    
    Debug.Print CharExecution
    
    Set pObject = Nothing
    Set lInterface = Nothing
    Set TLI = Nothing
  End Function

   '================================================================================

Function GetVariableName(index As Long) As String
    Select Case index
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
        Parameters = Parameters & IIf(Parameters <> "", ", ", "")
    
        If ParameterInf.flags And PARAMFLAG_FOPT Then
            Parameters = Parameters & "Optional "
        End If
        
        If ParameterInf.flags And PARAMFLAG_FOUT Then
            Parameters = Parameters & "ByRef "
        Else
            Parameters = Parameters & "ByVal "
        End If
        
        Parameters = Parameters & ParameterInf.Name & GetVariableName(ParameterInf.VarTypeInfo)
        
        If ParameterInf.flags And PARAMFLAG_FOPT Then
            If ParameterInf.Default Then
                
                Parameters = Parameters & " = "
                If ParameterInf.VarTypeInfo = VT_BSTR Then
                    Parameters = Parameters & """" & ParameterInf.DefaultValue & """"
                Else
                    Parameters = Parameters & ParameterInf.DefaultValue
                End If
            End If
        End If
        
        
    Next
    
    If ClassName <> "" Then ClassName = ClassName & "."
    
    On Error Resume Next
    
    Select Case lMember.DescKind
        Case DESCKIND_VARDESC
            ParseMember = "dim " & lMember.Name & GetVariableName(lMember.ReturnType.VarType)
            Debug.Print lMember.CustomDataCollection.Item(0).Value
            Exit Function
            
        Case DESCKIND_FUNCDESC
        
        Case Else
        
    End Select
    
    
    Select Case lMember.InvokeKind
        Case INVOKE_FUNC
            If lMember.ReturnType.VarType <> VT_VOID Then
                ParseMember = "Function " & ClassName & lMember.Name & "( " & Parameters & " )" & GetVariableName(lMember.ReturnType.VarType)
            Else
                ParseMember = "Sub " & ClassName & lMember.Name & "( " & Parameters & " )"
            End If
        Case INVOKE_PROPERTYGET
            ParseMember = "Property Get " & ClassName & lMember.Name & "( " & Parameters & " )" & GetVariableName(lMember.ReturnType.VarType)
        Case INVOKE_PROPERTYPUT
            ParseMember = "Property Let " & ClassName & lMember.Name & "( " & Parameters & " )"
        Case INVOKE_PROPERTYPUTREF
            ParseMember = "Property Set " & ClassName & lMember.Name & "( " & Parameters & " )"
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
Public Function Arccos(ByVal X As Double) As Double
   If Abs(X) <> 1 Then
       Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    Else
       Arccos = IIf(X = 1, 0, Atn(1) * 4)
    End If
End Function

' argument in radians
Public Function Arcsin(ByVal X As Double) As Double
   If Abs(X) <> 1 Then
       Arcsin = Atn(X / Sqr(-X * X + 1))
    Else
       Arcsin = IIf(X = 1, Atn(1) * 2, Atn(1) * 6)
    End If
End Function

'usefull c function missing in vb. X & Y are triangle's cathets/
Public Function Atan2(ByVal Y As Double, ByVal X As Double) As Double
   If X = 0 And Y = 0 Then
      Atan2 = 0
   Else
      Atan2 = Atn(Y / X) - PI * (X < 0)
   End If
End Function

'conversion RAD<->DEG
Public Function Rad(ByVal X As Double) As Double
  Rad = X * Atn(1) / 45#
End Function

Public Function Deg(ByVal X As Double) As Double
  Deg = X * 45# / Atn(1)
End Function

'Helpfull functions to compute sin, cos, tan
'with argument in degrees
Public Function Sind(ByVal X As Double) As Double
   Sind = Sin(Rad(X))
End Function

Public Function aSind(ByVal X As Double) As Double
   aSind = Deg(Arcsin(X))
End Function

Public Function Cosd(ByVal X As Double) As Double
   Cosd = Cos(Rad(X))
End Function

Public Function aCosd(ByVal X As Double) As Double
   aCosd = Deg(Arccos(X))
End Function

Public Function Tand(ByVal X As Double) As Double
   Tand = Tan(Rad(X))
End Function

Public Function aTand(ByVal X As Double) As Double
   aTand = Deg(Atn(X))
End Function

Public Function NormalizeAngle(ByVal X As Double) As Double
   Dim ret As Double
   ret = X - Int(X / 360#) * 360#
   If ret < 0 Then ret = ret + 360
   NormalizeAngle = ret
End Function
