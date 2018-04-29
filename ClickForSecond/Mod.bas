Attribute VB_Name = "Mod"
Option Explicit

Public Function InputItem(num As Integer) As Integer
Dim I%
    For I = 0 To num
        InputItem = NumTotalS(I)
    Next I
End Function

Public Sub UpdateItem(UItem())
Dim I%
    For I = 0 To NumTopI
        UItem(I) = NumTotalS(I)
    Next I
End Sub
