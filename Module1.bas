Attribute VB_Name = "FractionCalc"
Sub DecToFrac(DecimalNum As Double, Numerator As Long, Denom As Long)
    
    
   ' The BigNumber constant can be adjusted to handle larger fractional parts
   Const BigNumber = 1000
   Const SmallNumber = 0.000000000001

   Dim Inverse As Double, FractionalPart As Double
   Dim WholePart As Long, SwapTemp As Long

   Inverse = 1 / DecimalNum
   WholePart = Int(Inverse)
   FractionalPart = Frac(Inverse)

   If 1 / (FractionalPart + SmallNumber) < BigNumber Then
        ' Notice that DecToFrac is called recursively.
        Call DecToFrac(FractionalPart, Numerator, Denom)
        Numerator = Denom * WholePart + Numerator

        SwapTemp = Numerator
        Numerator = Denom
        Denom = SwapTemp
   Else ' If 1 / (FractionalPart + SmallNumber) > BigNumber
        ' Recursion stops when the final value of FractionalPart is 0 or
        ' close enough.  SmallNumber is added to prevent division by 0.
        Numerator = 1
        Denom = Int(Inverse)
   End If
End Sub

' This function is used by DecToFrac and DecToProperFact

Function Frac(x As Double) As Double
    Frac = Abs(Abs(x) - Int(Abs(x)))
End Function

' This additional procedure handles "improper" fractions and returns
' them in mixed form (a b/c) when the numerator is larger than the denominator

Sub DecToProperFrac(x As Double, a As Long, b As Long, c As Long)
   If x > 1 Then a = Int(x)
   If Frac(x) <> 0 Then
      Call DecToFrac(Frac(x), b, c)
   End If
End Sub

Public Function Dec2Frac(ByVal f As Double) As String
On Error GoTo EndIt:
   Dim df As Double
   Dim lUpperPart As Long
   Dim lLowerPart As Long
   
   lUpperPart = 1
   lLowerPart = 1
   
   df = lUpperPart / lLowerPart

   While (Round(df, 14) <> f)
      If (df < f) Then
         lUpperPart = lUpperPart + 1
         
      Else
         lLowerPart = lLowerPart + 1
         lUpperPart = f * lLowerPart
      End If
      df = lUpperPart / lLowerPart
   Wend
Dec2Frac = CStr(lUpperPart) & "/" & CStr(lLowerPart)
Exit Function
EndIt:
Dec2Frac = f
End Function


Function GetFraction(ByVal d As Double) As String
        
        Dim Denom As Double
        Dim Numer As Double
        Dim a As Double
        Dim b As Double
        Dim t As Double
        Dim tmpStr As String
        
        tmpStr = CStr(d)
        If InStr(1, tmpStr, ",") < 1 Then
            GetFraction = ""
            Exit Function
        End If
        tmpStr = Split(tmpStr, ",")(1)
        ' Get the initial denominator: 1 * (10 ^ decimal portion length)
        Denom = (1 * (10 ^ Len(tmpStr)))
        ' Get the initial numerator: integer portion of the number
        Numer = (tmpStr)
        
        Dim i As Long
        Dim x As Long
        
        
        Dim RepeatCheck As String
        
        If Len(tmpStr) > 7 Then
            For i = Len(tmpStr) / 2 To 1 Step -1
                RepeatCheck = Mid(tmpStr, 1, i)
                For x = i + 1 To Len(tmpStr) Step i
                    If Mid(tmpStr, x, i) = RepeatCheck Then
                        GoTo ReTime:
                    Else
                        Exit For
                    End If
                Next x
            Next i
            GoTo NotPosible
        End If
        
        ' Use the Euclidean algorithm to find the gcd
        a = Numer
        b = Denom
        t = 0 ' t is a value holder
        
        GoTo Euclidean:
        
ReTime:
        x = 10 ^ Len(RepeatCheck)
        x = x - 1
        
        
        a = CLng(RepeatCheck)
        b = x
        Numer = a
        Denom = b
        t = 0
Euclidean:


        
        
        ' Euclidean algorithm
        While b <> 0
            t = b
            b = a Mod b
            a = t
        Wend

        'Get whole part of the number
        Dim Whole As String
        Whole = Split(CStr(d), ",")(0)

        If Whole = 0 Then
            GetFraction = (Numer / a) & "/" & (Denom / a)
        Else
            GetFraction = Whole & " " & (Numer / a) & "/" & (Denom / a)
        End If
        ' Return our answer
        
        
        Exit Function
NotPosible:
        GetFraction = d
    End Function
    
