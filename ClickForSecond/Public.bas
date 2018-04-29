Attribute VB_Name = "Public"
Option Explicit
Public Const NumTopS = 4 - 1, NumTopR = 5 - 1
Public ItemV(NumTopS) As Long, ClickP As Integer
Public ResV(NumTopR) As Integer, ResT(NumTopR) As Integer, ResTI(1, NumTopR)
Public NumTotalS(NumTopS) As Integer, sper&, chg%, NumTotalR(NumTopR) As Boolean, NameR(NumTopR) As String
Public NumTotalRN(NumTopR) As Boolean
Public Sub Mainload() '������
    '�о���
    NameR(0) = "�ڿ��۾�����"
    NameR(1) = "�����ı����й�������"
    NameR(2) = "��ͧ����"
    NameR(3) = "�������ֱ���������"
    NameR(4) = "���������ݽ���"
    '��Ʒ����
    ItemV(0) = 10
    ItemV(1) = 20
    ItemV(2) = 45
    ItemV(3) = 90
    '�о�����
    ResV(0) = 20
    ResV(1) = 30
    ResV(2) = 60
    ResV(3) = 150
    ResV(4) = 100
    '�о�ʱ��
    ResT(0) = 10
    ResT(1) = 30
    ResT(2) = 60
    ResT(3) = 115
    ResT(4) = 70
End Sub
Public Sub Refe()
    Main.Total = str(Ts)
    For I = 0 To NumTopS
        ShopF.NumI(I) = "Ŀǰ��" & str(NumTotalS(I)) & "��"
    Next I
    Call NumPer
End Sub

Public Sub NumPer()
    sper = NumTotalS(0) * 1
    sper = sper + NumTotalS(1) * 2
    sper = sper + NumTotalS(2) * 5
    sper = sper + NumTotalS(3) * 10
    Main.Persec = "����1s��������:" & str(sper) & "s"
End Sub

Public Sub ResRef()
    ResearchF.Resing.Clear
    ResearchF.Resed.Clear
    ResearchF.Resable.Clear
    For I = 0 To NumTopR
        If ResTI(0, I) Then ResearchF.Resing.AddItem NameR(I)
        If NumTotalRN(I) Then ResearchF.Resable.AddItem NameR(I)
        If NumTotalR(I) Then ResearchF.Resed.AddItem NameR(I)
    Next I
End Sub

Public Sub ResShop()
    For I = 0 To NumTopS
        ShopF.NumI(I) = "Ŀǰ��" & str(NumTotalS(I)) & "��"
    Next I
    If NumTotalR(0) Then ShopF.BuyI(0).Enabled = True
    If NumTotalR(1) Then ShopF.BuyI(1).Enabled = True
    If NumTotalR(2) Then ShopF.BuyI(2).Enabled = True
    If NumTotalR(3) Then ShopF.BuyI(3).Enabled = True
End Sub

Public Sub saveF()
Dim ResHex As String, ISa As Integer
    ResHex = ResSave()
    Main.Common.DefaultExt = "savesecond"
    Main.Common.ShowSave
    If Main.Common.FileName = "" Then Exit Sub Else
    Open Main.Common.FileName For Output As #1
    Print #1, Main.User & "|" & Ts & "|";
    For ISa = 0 To NumTopS
        Print #1, NumTotalS(ISa) & "|";
    Next ISa
    Print #1, ClickP & "|" & ResHex;
    For ISa = 0 To NumTopR
        Print #1, ResTI(1, ISa) & "|";
    Next ISa
    Close #1
End Sub

Public Sub loadf()
Dim str As String, stuffstr, bitR, IL As Integer, ResTIstuff(NumTopR) As Boolean
    Main.Common.ShowOpen
    If Main.Common.FileName = "" Then Exit Sub Else
    Open Main.Common.FileName For Input As #1
    Line Input #1, str
    Close #1
    '�����ĵ�����
    stuffstr = Split(str, "|", -1)
    '0-�û��� 1-�ܹ����� 2~(4+2)-��Ʒ���� 4+3-���������
    '4+4-ʮ�������о���(�о����+�о���) (4+5)~(4+5+5)-�о�ʣ��
    Main.User = stuffstr(0)
    Ts = stuffstr(1)
    For IL = 0 To NumTopS
        NumTotalS(IL) = stuffstr(IL + 2)
    Next IL
    ClickP = stuffstr(NumTopS + 3)
    bitR = Split(stuffstr(NumTopS + 4), "+", 2)
    Call bitBoo(hexBit(bitR(0)), NumTotalR())
    Call bitBoo(hexBit(bitR(1)), ResTIstuff())
    For IL = 0 To NumTopR
        ResTI(0, IL) = ResTIstuff(IL)
        ResTI(1, IL) = stuffstr(IL + 5)
    Next IL
    Call Refe
End Sub

Public Function showde(ind As String) As String
    Select Case ind
        '������о�ʱֱ��ճ����������
        Case NameR(0): showde = "�����̵������ڿ��۾�." & vbCrLf _
        & "����" & ResV(0) & "s" & ",�о�ʱ��" & ResT(0) & "s"
        Case NameR(1): showde = "�����̵����������ı����й���." & vbCrLf _
        & "����" & ResV(1) & "s" & ",�о�ʱ��" & ResT(1) & "s"
        Case NameR(2): showde = "�����̵�������ͧ." & vbCrLf _
        & "����" & ResV(2) & "s" & ",�о�ʱ��" & ResT(2) & "s"
        Case NameR(3): showde = "�����̵��������ֱ���װ." & vbCrLf _
        & "����" & ResV(3) & "s" & ",�о�ʱ��" & ResT(3) & "s"
        Case NameR(4): showde = "Ϊ���������һ��������Ƹ�빤��." & vbCrLf _
        & "����" & ResV(4) & "s" & ",�о�ʱ��" & ResT(4) & "s" & vbCrLf & "ÿ�ε��Ч��+1"
        Case Else: showde = "����о���Ŀ��ʾ����" & vbCrLf & "���'�о�'��ť�Կ�ʼ�о�"
    End Select
End Function

Public Function ResNum(ind As String) As Integer
    For I = 0 To NumTopR
        If NameR(I) = ind Then ResNum = I: Exit Function
    Next I
    ResNum = -1
End Function
