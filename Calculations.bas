Attribute VB_Name = "Calculations"
Option Explicit

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
