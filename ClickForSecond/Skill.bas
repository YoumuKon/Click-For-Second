Attribute VB_Name = "Skill"
Option Explicit

Public Sub RunSkill(ind As Integer)
Dim RunST%, ST%, Num, succeed%
    succeed = -1
    Select Case ind
        Case 0:
        Num = InputBox("请输入技能使用数量", "喝枸杞茶")
        If Num <> "" Then
            If MsgBox("技能需要消耗" & Num & "个枸杞茶" & Chr(13) & "现在有" & NumTotalI(7) & "个枸杞茶" & Chr(13) & "确定要使用技能吗?", _
            vbYesNo, "喝枸杞茶") = vbYes Then
                If BuyCheck((Num), NumTotalI(7)) Then
                    ShopF.NumI(7) = "目前共" & NumTotalI(7) & "个"
                    For RunST = ResearchF.Resing.ListCount - 1 To 0 Step -1
                        ST = ResNum(ResearchF.Resing.List(RunST))
                        If ResTI(1, ST) > 0 Then ResTI(1, ST) = ResTI(1, ST) - 60 * Num
                        If ResTI(1, ST) < 0 Then ResTI(1, ST) = 0
                    Next RunST
                    succeed = 0
                    Else: MsgBox "物品数不够!", 16, "使用失败"
                End If
            End If
            Else: MsgBox "请输入物品数!", 16, "使用失败"
        End If
        Case Else: MsgBox "技能无效!", 16, "使用失败"
    End Select
    If succeed <> -1 Then
        MsgBox "技能使用成功!", 0, "使用成功"
        UpdEve StrEnc(StrEnc(EventList(5), "&U", UserN), "&Mem1", NameS(0, succeed))
        Select Case succeed
            Case 0: UpdEve StrEnc(StrEnc(NameS(1, succeed), "&Mem1", Num), "&Mem2", 60 * Num)
        End Select
    End If
End Sub
