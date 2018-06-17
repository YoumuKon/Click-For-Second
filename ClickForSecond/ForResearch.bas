Attribute VB_Name = "ForResearch"
Option Explicit
Public Sub ResRef()
Dim IRR%
    ResearchF.Resing.Clear
    ResearchF.Resed.Clear
    ResearchF.Resable.Clear
    For IRR = 0 To NumTopR
        If ResTI(0, IRR) Then ResearchF.Resing.AddItem NameR(0, IRR)
        If NumTotalRN(IRR) Then ResearchF.Resable.AddItem NameR(0, IRR)
        If NumTotalR(IRR) Then ResearchF.Resed.AddItem NameR(0, IRR)
    Next IRR
End Sub

Public Function showde(Ind As String) As String
Dim NumR%, strI, strN, I%
    NumR = ResNum(Ind)
    If NumR < 0 Then
        showde = "����о���Ŀ��ʾ����" & vbCrLf & "���'�о�'��ť�Կ�ʼ�о�": Exit Function
        Else: showde = StrEnc(NameR(1, NumR), "&CL", vbCrLf) & vbCrLf & "����" & ResV(NumR) & "s" & ",�о�ʱ��" & ResT(NumR) & "s"
    End If
    strI = Split(ResVI(0, NumR), "|")
    strN = Split(ResVI(1, NumR), "|")
    If UBound(strN) > 0 Then
        showde = showde & vbCrLf & "�Լ�:"
        For I = 0 To UBound(strN) - 1
            showde = showde & " " & NameI(strI(I)) & ":" & strN(I)
        Next I
    End If
    showde = Ind & vbCrLf & showde
End Function

Public Function ResNum(Ind As String) As Integer
Dim IRM%
    ResNum = -1
    For IRM = 0 To NumTopR
        If NameR(0, IRM) = Ind Then ResNum = IRM: Exit Function
    Next IRM
End Function

Public Sub showWP(Ind As Integer)
    If Ind >= 0 Then
        Main.WorkPlace.Caption = WPevent(Ind + 1)
        Select Case Ind
            Case 0: ClickP = ClickP + 1
            Case 1: ClickP = ClickP + 3
            Case 2: ClickP = ClickP * 1.4
            Case 3: ItemPST = ItemPST + 0.1
        End Select
        Else: Main.WorkPlace.Caption = "���ǹ�����"
    End If
End Sub

Public Sub CheckRes()
Dim updateR As Boolean, I As Integer, J As Integer, strF, strI, strT, str1
Dim CanUpd As Boolean
    updateR = False
    '�ж���
    CanUpd = True
    For I = 0 To NumTopRN
        Call Needcele(ResNeed(I), strF, strI, strT)
        For J = 0 To UBound(strF)
            If Not NumTotalR(strF(J)) Then CanUpd = False
        Next J
        For J = 0 To UBound(strI)
            str1 = Split(strI(J), "*")
            If NumTotalI(str1(0)) < str1(1) Then CanUpd = False
        Next J
        If CanUpd Then
            For J = 0 To UBound(strT)
                If updCed(I) Then
                    GoTo nextJ
                    Else
                    NumTotalRN(strT(J)) = True
                    updateR = True
                End If
nextJ:
            Next J
            updCed(I) = True
        End If
    Next I
    If NumTotalR(26) Then Call showWP(3): GoTo jumpWP
    If NumTotalR(25) Then Call showWP(2): GoTo jumpWP
    If NumTotalR(24) Then Call showWP(1): GoTo jumpWP
    If NumTotalR(23) Then Call showWP(0): GoTo jumpWP
jumpWP:
    '�ظ����
    '���ȼ��������>������>δ���
    For I = 0 To NumTopR
        If NumTotalR(I) Then
            ResTI(0, I) = False
            ResTI(1, I) = 0
            NumTotalRN(I) = False
            ElseIf ResTI(0, I) Then NumTotalRN(I) = False
        End If
    Next I
    '�����о��б�
    If updateR Then Call ResRef: Call ResRefresh
End Sub

