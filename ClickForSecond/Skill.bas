Attribute VB_Name = "Skill"
Option Explicit

Public Sub RunSkill(ind As Integer, Optional Int0 As Integer)
Dim RunST%, ST%
    Select Case ind
        Case 0:
        If BuyCheck((Int0), NumTotalS(6)) Then
            For RunST = ResearchF.Resing.ListCount - 1 To 0 Step -1
                ST = -1
                Do While ST = -1
                    ST = ResNum(ResearchF.Resing.List(RunST))
                Loop
                If ResTI(1, ST) > 0 Then ResTI(1, ST) = ResTI(1, ST) - 60 * Int0
                If ResTI(1, ST) < 0 Then ResTI(1, ST) = 0
            Next RunST
            MsgBox "����ʹ�óɹ�!", 0, "ʹ�óɹ�"
            Else: MsgBox "��Ʒ������!", 16, "ʹ��ʧ��"
        End If
        ShopF.NumI(6) = "Ŀǰ��" & NumTotalS(6) & "��"
        Case Else: MsgBox "�޴˼���!", 16, "ʹ��ʧ��"
    End Select
End Sub
