Attribute VB_Name = "ForBuildCopy"
Option Explicit
Public Sub Buildref()
Dim I%
    BuildingF.Building.Clear
    BuildingF.Builded.Clear
    BuildingF.Buildable.Clear
    For I = 0 To NumTopB
        If BuildTI(0, I) Then BuildingF.Building.AddItem NameB(0, I)
        If NumTotalBN(I) Then BuildingF.Buildable.AddItem NameB(0, I)
        If NumTotalB(I) Then BuildingF.Builded.AddItem NameB(0, I)
    Next I
End Sub

Public Function showBuildde(ind As String) As String
Dim NumB%, strI, strN, I%
    NumB = BuildNum(ind)
    If NumB < 0 Then
        showBuildde = "点击建筑项目显示描述" & vbCrLf & "点击'建造'按钮以开始建造": Exit Function
        Else: showBuildde = StrEnc(NameB(1, NumB), StrCrlf, vbCrLf) & vbCrLf & "消耗" & BuildV(NumB) & "s" & ",研究时长" & BuildT(NumB) & "s"
    End If
    strI = Split(BuildVI(0, NumB), "|")
    strN = Split(BuildVI(1, NumB), "|")
    If UBound(strN) > 0 Then
        showBuildde = showBuildde & vbCrLf & "以及:"
        For I = 0 To UBound(strN) - 1
            showBuildde = showBuildde & " " & NameI(strI(I)) & ":" & strN(I)
        Next I
    End If
    showBuildde = ind & vbCrLf & showBuildde
End Function

Public Function BuildNum(ind As String) As Integer
Dim IRM%
    BuildNum = -1
    For IRM = 0 To NumTopB
        If NameB(0, IRM) = ind Then BuildNum = IRM: Exit Function
    Next IRM
End Function





