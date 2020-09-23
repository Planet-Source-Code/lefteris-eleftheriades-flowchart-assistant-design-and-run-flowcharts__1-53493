Attribute VB_Name = "trigDeviations"
Option Explicit
Private Const Pi = 3.14159265358979
'Taken from: mk:@MSITStore:C:\Program Files\Microsoft Visual Studio\MSDN98\98VSa\1033\office95.chm::/html/S11624.HTM
Public Function Arcsin(x#) As Double
    If Abs(x) = 1 Then
        Arcsin = x * 1.5707963267949
    Else
        Arcsin = Atn(x / Sqr(-x * x + 1))
    End If
End Function

Public Function Arccos(x#) As Double
    If x = -1 Then
        Arccos = 3.14159265359
    ElseIf x = 1 Then
        Arccos = 0
    Else
        Arccos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    End If
End Function

'It's true you can know if a line is clicked, by calling this function on its container control
Public Function LineMouseEvent(LineName As Line, x As Single, y As Single, Optional Aura% = 0) As Boolean
  'You are propably wandering wtf is aura. if Aura is big then the line will be considered as clicked even if it is not exactly clicked on. if it's zero the event will only trigger on exact click
  On Error GoTo ExitF
  Dim XMin&, XMax&, YMin&, YMax&, Gradient As Currency, HalfBordWid&
  If LineName.X1 < LineName.X2 Then
     XMin = LineName.X1
     XMax = LineName.X2
  Else
     XMin = LineName.X2
     XMax = LineName.X1
  End If
  If LineName.Y1 < LineName.Y2 Then
     YMin = LineName.Y1
     YMax = LineName.Y2
  Else
     YMin = LineName.Y2
     YMax = LineName.Y1
  End If
  HalfBordWid = (LineName.BorderWidth + Aura) / 2
  If x >= XMin - HalfBordWid And x <= XMax + HalfBordWid And y >= YMin - HalfBordWid And y <= YMax + HalfBordWid Then
       'calculate the line vector equation and check the Y values
       If LineName.X2 - LineName.X1 = 0 Then
          LineMouseEvent = True
       Else
          Gradient = (LineName.Y2 - LineName.Y1) / (LineName.X2 - LineName.X1)
          'Line Equation is: Y - LineName.Y2 = Gradient * (X - LineName.X2)
           LineMouseEvent = CBool(Abs(Gradient * (x - LineName.X2) - (y - LineName.Y2)) < (LineName.BorderWidth + Aura))
       End If
  End If
ExitF:
End Function

Public Function SuitableSide(cx&, cy&, W&, H&, x&, y&) As Byte
Dim A As Currency
  If cx - x <> 0 Then
     A = Atn((cy - y) / (cx - x))
  Else
     If cy > y Then
        A = Pi / 2
     Else
        A = -Pi / 2
     End If
  End If

If cx >= x Then
   If cy < y Then
      A = 2 * Pi + A
   End If
Else
      A = Pi + A
End If

If A < Atn(H / W) Then SuitableSide = 1
If A > Atn(H / W) Then SuitableSide = 2
If A > Pi - Atn(H / W) Then SuitableSide = 3
If A > Atn(H / W) + Pi Then SuitableSide = 4
If A > 2 * Pi - Atn(H / W) Then SuitableSide = 1
End Function

