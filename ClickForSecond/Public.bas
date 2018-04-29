Attribute VB_Name = "Public"
Option Explicit
Public Const NumTopS = 4 - 1, NumTopR = 6 - 1
Public ItemV(NumTopS) As Double, ClickP As Integer, NameI(NumTopS) As String, ItemPS(NumTopS) As Double
Public ResV(NumTopR) As Double, ResT(NumTopR) As Integer, ResTI(1, NumTopR)
Public NumTotalS(NumTopS) As Integer, sper As Double, chg%, NumTotalR(NumTopR) As Boolean, NameR(NumTopR) As String
Public NumTotalRN(NumTopR) As Boolean
Public Sub Mainload() '������
    '�о���
    NameR(0) = "�ڿ��۾�����"
    NameR(1) = "�����ı����й�������"
    NameR(2) = "��ͧ����"
    NameR(3) = "�������ֱ���������"
    NameR(4) = "���������ݽ���"
    NameR(5) = "�����խ���۾�����"
    '��Ʒ����
    ItemV(0) = 10
    ItemV(1) = 20
    ItemV(2) = 45
    ItemV(3) = 90
    '��Ʒ��
    NameI(0) = "�ڿ��۾� " & ItemV(0) & "s"
    NameI(1) = "�����ı����й��� " & ItemV(1) & "s"
    NameI(2) = "��ͧ " & ItemV(2) & "s"
    NameI(3) = "���ֱ���װ " & ItemV(3) & "s"
    '�о�����
    ResV(0) = 20
    ResV(1) = 30
    ResV(2) = 60
    ResV(3) = 150
    ResV(4) = 100
    ResV(5) = 50
    '�о�ʱ��
    ResT(0) = 10
    ResT(1) = 30
    ResT(2) = 60
    ResT(3) = 115
    ResT(4) = 70
    ResT(5) = 35
End Sub
Public Sub Refe()
Dim IR%
    Main.Total = str(Ts)
    For IR = 0 To NumTopS
        ShopF.NumI(IR) = "Ŀǰ��" & str(NumTotalS(IR)) & "��"
    Next IR
    Call NumPer
End Sub

Public Sub NumPer() 'ÿ��������=��ÿ����Ʒ��(��Ʒ����*��ƷЧ��)���
    sper = NumTotalS(0) * 1 * ItemPS(0)
    sper = sper + NumTotalS(1) * 2 * ItemPS(1)
    sper = sper + NumTotalS(2) * 5 * ItemPS(2)
    sper = sper + NumTotalS(3) * 10 * ItemPS(3)
    Main.Persec = "����1s��������:" & str(sper) & "s"
End Sub

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

Public Sub ResShop()
Dim IRS%, updPSed(NumTopS, 1) As Integer
    For IRS = 0 To NumTopS
        ShopF.BuyI(IRS).Caption = NameI(IRS)
        ShopF.NumI(IRS) = "Ŀǰ��" & str(NumTotalS(IRS)) & "��"
    Next IRS
    If NumTotalR(0) Then ShopF.BuyI(0).Enabled = True
    If NumTotalR(1) Then ShopF.BuyI(1).Enabled = True
    If NumTotalR(2) Then ShopF.BuyI(2).Enabled = True
    If NumTotalR(3) Then ShopF.BuyI(3).Enabled = True
    If NumTotalR(0) And NumTotalR(5) And Not updPSed(0, 0) Then _
    NameI(0) = "�����խ���۾� " & ItemV(0) & "s": ItemPS(0) = 1.5: updPSed(0, 0) = False
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
    For ISa = 0 To NumTopS
        Print #1, ItemPS(ISa) & "|";
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
    '0-�û��� 1-�ܹ����� 2~(2+4)-��Ʒ���� 3+4-���������
    '4+4-ʮ�������о���(�о����+�о���) (5+4)~(5+4+6)-�о�ʣ��
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
        & "����" & ResV(4) & "s" & ",�о�ʱ��" & ResT(4) & "s" & vbCrLf & "ÿ�ε��Ч��Ϊ2"
        Case NameR(5): showde = "���ڿ��۾�����Ϊ�����խ���۾�." & vbCrLf _
        & "����" & ResV(5) & "s" & ",�о�ʱ��" & ResT(5) & "s" & vbCrLf & "�ڿ��۾�Ч��+50%"
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
        Case 4: Main.WorkPlace.Caption = "���Ƕ���һ�����ӵĹ�����"
        Case Else: Main.WorkPlace.Caption = "���ǹ�����"
    End Select
End Sub

Public Sub UpdEve(str$)
    Main.EventS = str & vbCrLf & Main.EventS
End Sub
