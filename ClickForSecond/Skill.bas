Attribute VB_Name = "Skill"
Option Explicit

Public Sub RunSkill(ind As Integer)
Dim RunST%, ST%, Num, succeed%
    succeed = -1
    Select Case ind
        Case 0:
        Num = InputBox("�����뼼��ʹ������", "����轲�")
        If Num <> "" Then
            If MsgBox("������Ҫ����" & Num & "����轲�" & Chr(13) & "������" & NumTotalI(7) & "����轲�" & Chr(13) & "ȷ��Ҫʹ�ü�����?", _
            vbYesNo, "����轲�") = vbYes Then
                If BuyCheck((Num), NumTotalI(7)) Then
                    ShopF.NumI(7) = "Ŀǰ��" & NumTotalI(7) & "��"
                    For RunST = ResearchF.Resing.ListCount - 1 To 0 Step -1
                        ST = ResNum(ResearchF.Resing.List(RunST))
                        If ResTI(1, ST) > 0 Then ResTI(1, ST) = ResTI(1, ST) - 60 * Num
                        If ResTI(1, ST) < 0 Then ResTI(1, ST) = 0
                    Next RunST
                    succeed = 0
                    Else: MsgBox "��Ʒ������!", 16, "ʹ��ʧ��"
                End If
            End If
            Else: MsgBox "��������Ʒ��!", 16, "ʹ��ʧ��"
        End If
        Case Else: MsgBox "������Ч!", 16, "ʹ��ʧ��"
    End Select
    If succeed <> -1 Then
        MsgBox "����ʹ�óɹ�!", 0, "ʹ�óɹ�"
        UpdEve StrEnc(StrEnc(EventList(5), "&U", UserN), "&Mem1", NameS(0, succeed))
        Select Case succeed
            Case 0: UpdEve StrEnc(StrEnc(NameS(1, succeed), "&Mem1", Num), "&Mem2", 60 * Num)
        End Select
    End If
End Sub
