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
        Case NameR(11): showde = "���ڽ���Ƭ��¼��VCD." & vbCrLf _
        & "����" & ResV(11) & "s" & ",�о�ʱ��" & ResT(11) & "s" & vbCrLf & "Ч��+50%"
        Case NameR(12): showde = "Ϊ����������һ��������Ƹ�빤��." & vbCrLf _
        & "����" & ResV(12) & "s" & ",�о�ʱ��" & ResT(12) & "s" & vbCrLf & "ÿ�ε��Ч��+1"
        Case NameR(13): showde = "Ϊ�������ķ�������ΪԱ��������Ƹ����๤��." & vbCrLf _
        & "����" & ResV(13) & "s" & ",�о�ʱ��" & ResT(13) & "s" & vbCrLf & "ÿ�ε��Ч��+3"
        Case NameR(14): showde = "Ϊ���������һ���㳡����߹��˻�����." & vbCrLf _
        & "����" & ResV(14) & "s" & ",�о�ʱ��" & ResT(14) & "s" & vbCrLf & "ÿ�ε��Ч��+40%"
        Case NameR(15): showde = "�����̵�������轲�." & vbCrLf _
        & "����" & ResV(15) & "s" & ",�о�ʱ��" & ResT(15) & "s"
        Case Else: showde = "����о���Ŀ��ʾ����" & vbCrLf & "���'�о�'��ť�Կ�ʼ�о�"
    End Select
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
        Case 2: Main.WorkPlace.Caption = "������Ա���������׵Ĺ������㳡"
        Case Else: Main.WorkPlace.Caption = "���ǹ�����"
    End Select
End Sub

Public Sub CheckRes()
Dim updateR As Boolean
    updateR = False
    'Ŀǰ�ж���
    '��10����Ʒ0ʱ�����о�1��6
    If NumTotalS(0) >= 10 And NumTotalR(0) And Not updCed(0) Then _
    NumTotalRN(6) = True: NumTotalRN(1) = True: updCed(0) = True: updateR = True
    '��10����Ʒ1ʱ�����о�2��7
    If NumTotalS(1) >= 10 And NumTotalR(1) And Not updCed(1) Then _
    NumTotalRN(7) = True: NumTotalRN(2) = True: updCed(1) = True: updateR = True
    '��10����Ʒ2ʱ�����о�3, 15��8
    If NumTotalS(2) >= 10 And NumTotalR(2) And Not updCed(2) Then _
    NumTotalRN(8) = True: NumTotalRN(15) = True: NumTotalRN(3) = True: updCed(2) = True: updateR = True
    '��10����Ʒ3ʱ�����о�4��9
    If NumTotalS(3) >= 10 And NumTotalR(3) And Not updCed(3) Then _
    NumTotalRN(9) = True: NumTotalRN(4) = True: updCed(3) = True: updateR = True
    '��10����Ʒ4ʱ�����о�5��10
    If NumTotalS(4) >= 10 And NumTotalR(4) And Not updCed(4) Then _
    NumTotalRN(10) = True: NumTotalRN(5) = True: updCed(4) = True: updateR = True
    '��10����Ʒ5ʱ�����о�11
    If NumTotalS(5) >= 10 And NumTotalR(5) And Not updCed(5) Then _
    NumTotalRN(11) = True: updCed(5) = True: updateR = True
    '�о�12,13,14��ɺ���¹�������Ϣ
    If NumTotalR(12) And Not updCed(12) Then _
    ClickP = ClickP + 1: updCed(12) = True: NumTotalRN(13) = True: Call showWP(0): updateR = True
    If NumTotalR(13) And Not updCed(13) Then _
    ClickP = ClickP + 3: updCed(13) = True: NumTotalRN(14) = True: Call showWP(1): updateR = True
    If NumTotalR(14) And Not updCed(14) Then _
    ClickP = ClickP * 1.4: updCed(14) = True: Call showWP(2): updateR = True
    '�����о��б�
    If updateR Then Call ResRef: Call ResShop
End Sub

