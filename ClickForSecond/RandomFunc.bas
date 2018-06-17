Attribute VB_Name = "RandomFunc"
Option Explicit

Public Function ProCele(Index As Double) As Boolean
Dim num1 As Double
    Randomize
    num1 = CDbl(Rnd)
    If num1 < Index Then
        ProCele = True
        Else
        ProCele = False
    End If
End Function
