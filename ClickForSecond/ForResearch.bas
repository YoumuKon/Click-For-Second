Attribute VB_Name = "ForResearch"
Option Explicit
Public Sub ResRef()
Dim IRR%
    ResearchF.Resing.Clear
    ResearchF.Resed.Clear
    ResearchF.Resable.Clear
    For IRR = 0 To NumTopR
        If ResTI(0, IRR) Then ResearchF.Resing.AddItem NameR(IRR)
        If NumTotalRN(IRR) Then ResearchF.Resable.AddItem NameR(IRR)
        If NumTotalR(IRR) Then ResearchF.Resed.AddItem NameR(IRR)
    Next IRR
End Sub

Public Function showde(ind As String) As String
    Select Case ind
        '������о�ʱֱ��ճ��
        Case NameR(0): showde = "�����̵������ڿ��۾�." & vbCrLf _
        & "����" & ResV(0) & "s" & ",�о�ʱ��" & ResT(0) & "s"
        Case NameR(1): showde = "�����̵����������ı����й���." & vbCrLf _
        & "����" & ResV(1) & "s" & ",�о�ʱ��" & ResT(1) & "s"
        Case NameR(2): showde = "�����̵�������е�ֱ���װ." & vbCrLf _
        & "����" & ResV(2) & "s" & ",�о�ʱ��" & ResT(2) & "s"
        Case NameR(3): showde = "�����̵�������ͨѼ���." & vbCrLf _
        & "����" & ResV(3) & "s" & ",�о�ʱ��" & ResT(3) & "s"
        Case NameR(4): showde = "�����̵�������ͧ." & vbCrLf _
        & "����" & ResV(4) & "s" & ",�о�ʱ��" & ResT(4) & "s"
        Case NameR(5): showde = "�����̵�������Aloha 'Oe���ڽ���Ƭ." & vbCrLf _
        & "����" & ResV(5) & "s" & ",�о�ʱ��" & ResT(5) & "s"
        Case NameR(6): showde = "���ڿ��۾�����Ϊ�����խ���۾�." & vbCrLf _
        & "����" & ResV(6) & "s" & ",�о�ʱ��" & ResT(6) & "s" & vbCrLf & "Ч��+50%"
        Case NameR(7): showde = "�������ı����й�������Ϊ����������ѡ��." & vbCrLf _
        & "����" & ResV(7) & "s" & ",�о�ʱ��" & ResT(7) & "s" & vbCrLf & "Ч��+50%"
        Case NameR(8): showde = "�������е��ֱ�����Ϊ�����ֱ�." & vbCrLf _
        & "����" & ResV(8) & "s" & ",�о�ʱ��" & ResT(8) & "s" & vbCrLf & "Ч��+50%"
        Case NameR(9): showde = "��Ѽ�������Ϊ��ЧѼ���." & vbCrLf _
        & "����" & ResV(9) & "s" & ",�о�ʱ��" & ResT(9) & "s" & vbCrLf & "Ч��+50%"
        Case NameR(10): showde = "����ͧ�Ĳ�������Ϊ���ϲ���." & vbCrLf _
        & "����" & ResV(10) & "s" & ",�о�ʱ��" & ResT(10) & "s" & vbCrLf & "Ч��+50%"
        Case NameR(11): showde = "����Aloha 'Oe��ת¼���Ŵ�." & vbCrLf _
        & "����" & ResV(11) & "s" & ",�о�ʱ��" & ResT(11) & "s" & vbCrLf & "Ч��+50%"
        Case NameR(12): showde = "�����̵�������轲�." & vbCrLf _
        & "����" & ResV(12) & "s" & ",�о�ʱ��" & ResT(12) & "s"
        Case NameR(13): showde = "�������խ���۾�����Ϊ��裿��۾�." & vbCrLf _
        & "����" & ResV(13) & "s" & ",�о�ʱ��" & ResT(13) & "s" & vbCrLf & "Ч��+100%"
        Case NameR(14): showde = "������������ѡ������Ϊ������������." & vbCrLf _
        & "����" & ResV(14) & "s" & ",�о�ʱ��" & ResT(14) & "s" & vbCrLf & "Ч��+100%"
        Case NameR(15): showde = "�������е��ֱ�����Ϊ��Яԭ����." & vbCrLf _
        & "����" & ResV(15) & "s" & ",�о�ʱ��" & ResT(15) & "s" & vbCrLf & "Ч��+100%"
        Case NameR(16): showde = "��Ѽ�������Ϊ�Զ�Ѽ���." & vbCrLf _
        & "����" & ResV(16) & "s" & ",�о�ʱ��" & ResT(16) & "s" & vbCrLf & "Ч��+100%"
        Case NameR(17): showde = "����ͧ����������һ����" & vbCrLf _
        & "����" & ResV(17) & "s" & ",�о�ʱ��" & ResT(17) & "s" & vbCrLf & "Ч��+100%"
        Case NameR(18): showde = "����Aloha 'Oe��ת¼��DVD." & vbCrLf _
        & "����" & ResV(18) & "s" & ",�о�ʱ��" & ResT(18) & "s" & vbCrLf & "Ч��+100%"
        Case NameR(19): showde = "Ϊ����������һ��������Ƹ�빤��." & vbCrLf _
        & "����" & ResV(19) & "s" & ",�о�ʱ��" & ResT(19) & "s" & vbCrLf & "ÿ�ε��Ч��+1"
        Case NameR(20): showde = "���������ķ�������ΪԱ��������Ƹ����๤��." & vbCrLf _
        & "����" & ResV(20) & "s" & ",�о�ʱ��" & ResT(20) & "s" & vbCrLf & "ÿ�ε��Ч��+3"
        Case NameR(21): showde = "Ϊ���������һ���㳡����߹��˻�����." & vbCrLf _
        & "����" & ResV(21) & "s" & ",�о�ʱ��" & ResT(21) & "s" & vbCrLf & "ÿ�ε��Ч��+40%"
        Case NameR(22): showde = "Ϊ���������һ���������������Ч��." & vbCrLf _
        & "����" & ResV(22) & "s" & ",�о�ʱ��" & ResT(22) & "s" & vbCrLf & "ȫ���Զ�������ƷЧ��+10%"
        Case NameR(23): showde = "�ڹ㳡����ʱ�䷨���Թ��о��߼�֪ʶ����." & vbCrLf _
        & "����" & ResV(23) & "s" & ",�о�ʱ��" & ResT(23) & "s" & vbCrLf & "�����߼��о�"
        Case Else: showde = "����о���Ŀ��ʾ����" & vbCrLf & "���'�о�'��ť�Կ�ʼ�о�"
    End Select
    showde = ind & vbCrLf & showde
End Function

Public Function ResNum(ind As String) As Integer
Dim IRM%
    For IRM = 0 To NumTopR
        If NameR(IRM) = ind Then ResNum = IRM: Exit Function
    Next IRM
    ResNum = -1
End Function

Public Sub showWP(ind As Integer)
    Select Case ind
        Case 0: Main.WorkPlace.Caption = "���Ƕ���һ�����ӵĹ�����"
        Case 1: Main.WorkPlace.Caption = "������Ա���������׵Ĺ�����"
        Case 2: Main.WorkPlace.Caption = "������Ա���������׵Ĺ�����Ա���㳡"
        Case 3: Main.WorkPlace.Caption = "������Ա�����ᡢԱ���㳡���׵Ĺ���������"
        Case Else: Main.WorkPlace.Caption = "���ǹ�����"
    End Select
End Sub

Public Sub CheckRes()
Dim updateR As Boolean
    updateR = False
    'Ŀǰ�ж���
    '��10����Ʒ0ʱ�����о�1��6
    If NumTotalI(0) >= 10 And NumTotalR(0) And Not updCed(6) Then _
    NumTotalRN(6) = True: NumTotalRN(1) = True: updCed(6) = True: updateR = True
    '��10����Ʒ1ʱ�����о�2��7
    If NumTotalI(1) >= 10 And NumTotalR(1) And Not updCed(7) Then _
    NumTotalRN(7) = True: NumTotalRN(2) = True: updCed(7) = True: updateR = True
    '��10����Ʒ2ʱ�����о�3, 12��8
    If NumTotalI(2) >= 10 And NumTotalR(2) And Not updCed(8) Then _
    NumTotalRN(8) = True: NumTotalRN(12) = True: NumTotalRN(3) = True: updCed(8) = True: updateR = True
    '��10����Ʒ3ʱ�����о�4��9
    If NumTotalI(3) >= 10 And NumTotalR(3) And Not updCed(9) Then _
    NumTotalRN(9) = True: NumTotalRN(4) = True: updCed(9) = True: updateR = True
    '��10����Ʒ4ʱ�����о�5��10
    If NumTotalI(4) >= 10 And NumTotalR(4) And Not updCed(10) Then _
    NumTotalRN(10) = True: NumTotalRN(5) = True: updCed(10) = True: updateR = True
    '��10����Ʒ5ʱ�����о�11��23(��Ҫ21����)
    If NumTotalI(5) >= 10 And NumTotalR(5) And Not updCed(11) Then _
    NumTotalRN(11) = True: updCed(11) = True: updateR = True
    If NumTotalI(5) >= 10 And NumTotalR(21) And Not updCed(23) Then NumTotalRN(23) = True: updCed(23) = True: updateR = True
    '---�߼��о�---
    If NumTotalR(23) Then
        '��50����Ʒ0ʱ�����о�13
        If NumTotalI(0) >= 10 And NumTotalR(6) And Not updCed(13) Then _
        NumTotalRN(13) = True: updCed(13) = True: updateR = True
        '��50����Ʒ1ʱ�����о�14
        If NumTotalI(1) >= 10 And NumTotalR(7) And Not updCed(14) Then _
        NumTotalRN(14) = True: updCed(14) = True: updateR = True
        '��50����Ʒ2ʱ�����о�15
        If NumTotalI(2) >= 10 And NumTotalR(8) And Not updCed(15) Then _
        NumTotalRN(15) = True: updCed(15) = True: updateR = True
        '��50����Ʒ3ʱ�����о�16
        If NumTotalI(3) >= 10 And NumTotalR(9) And Not updCed(16) Then _
        NumTotalRN(16) = True: updCed(16) = True: updateR = True
        '��50����Ʒ4ʱ�����о�17
        If NumTotalI(4) >= 10 And NumTotalR(10) And Not updCed(17) Then _
        NumTotalRN(17) = True: updCed(17) = True: updateR = True
        '��50����Ʒ5ʱ�����о�18
        If NumTotalI(5) >= 10 And NumTotalR(11) And Not updCed(18) Then _
        NumTotalRN(18) = True: updCed(18) = True: updateR = True
    End If
    If NumTotalR(19) And Not updCed(19) Then _
    ClickP = ClickP + 1: updCed(19) = True: NumTotalRN(20) = True: Call showWP(0): updateR = True
    If NumTotalR(20) And Not updCed(20) Then _
    ClickP = ClickP + 3: updCed(20) = True: NumTotalRN(21) = True: Call showWP(1): updateR = True
    If NumTotalR(21) And Not updCed(21) Then
        ClickP = ClickP * 1.4: updCed(21) = True
        If NumTotalR(23) Then NumTotalRN(22) = True '22Ϊ�߼��о�
        Call showWP(2): updateR = True
    End If
    If NumTotalR(22) And Not updCed(22) Then _
    ItemPST = ItemPST + 0.1: updCed(22) = True: Call showWP(3): updateR = True
    '�����о��б�
    If updateR Then Call ResRef: Call ResShop
End Sub

