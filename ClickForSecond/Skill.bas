Attribute VB_Name = "Skill"
Option Explicit

Public Sub RunSkill(ind As Integer)
Dim RunST%, ST%, num, succeed%
    succeed = -1
    Select Case ind
        Case 0:
        num = InputBox("�����뼼��ʹ������", "����轲�")
        If num <> "" Then
            If MsgBox("������Ҫ����" & num & "����轲�" & Chr(13) & "������" & NumTotalI(7) & "����轲�" & Chr(13) & "ȷ��Ҫʹ�ü�����?", _
            vbYesNo, "����轲�") = vbYes Then
                If BuyCheck(Int(num), NumTotalI(7)) Then
                    ShopF.NumI(7) = "Ŀǰ��" & NumTotalI(7) & "��"
                    For RunST = ResearchF.Resing.ListCount - 1 To 0 Step -1
                        ST = ResNum(ResearchF.Resing.List(RunST))
                        If RO(ST).TimeNow > 0 Then RO(ST).TimeNow = RO(ST).TimeNow - 60 * num
                        If RO(ST).TimeNow < 0 Then RO(ST).TimeNow = 0
                        OnlineTime = OnlineTime + 60 * num
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
        UpdEve StrEnc(StrEnc(EventList(5), StrUser, UserN), StrMem1, NameS(0, succeed))
        Select Case succeed
            Case 0: UpdEve StrEnc(StrEnc(NameS(1, succeed), StrMem1, num), StrMem2, 60 * num)
        End Select
    End If
End Sub
