Attribute VB_Name = "RandomFunc"
Option Explicit

Public Function ProCele(ByVal Index As Double) As Boolean
Dim num As Double, I%
    Randomize
    num = Int((10000 + 1) * Rnd)
    Debug.Print "抽取随机数为:" & num
    Index = Int(Index * 10000)
    Debug.Print "指定概率为:" & Index
    ProCele = IIf(num < Index, True, False)
    If ProCele Then
        Debug.Print "审查通过"
        Else
        Debug.Print "审查不通过"
    End If
End Function

Public Sub RunRandomE(ind As Integer)
Dim I%, REnum1, REnum2
    REnum1 = ""
    REnum2 = ""
    If Not ProCele(Reventlist(2, ind)) Then Exit Sub
    Randomize
    Select Case ind
        Case 0:
            REnum1 = Int((3600 + 1) * Rnd)
            For I = ResearchF.Resing.ListCount - 1 To 0 Step -1
                If RO(I).TimeNow > 0 Then O(I).TimeNow = O(I).TimeNow + REnum1
            Next I
        Case 1:
            REnum1 = Int((NumTopI + 1) * Rnd)
            NumTotalI(REnum1) = NumTotalI(REnum1) + 1
    End Select
    UpdEve StrEnc(StrEnc(Reventlist(1, ind), StrMem1, REnum1), StrMem2, REnum2)
End Sub
