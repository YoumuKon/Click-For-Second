Attribute VB_Name = "Public"
Option Explicit
Public Const NumTopS = 5 - 1, NumTopR = 13 - 1
Public updCed(NumTopR) As Boolean, updPSed(NumTopS, 1) As Boolean
Public ItemV(NumTopS) As Double, ClickP As Integer, NameI(NumTopS) As String, ItemPS(NumTopS) As Double
Public ResV(NumTopR) As Double, ResT(NumTopR) As Integer, ResTI(1, NumTopR)
Public NumTotalS(NumTopS) As Integer, sper As Double, chg%, NumTotalR(NumTopR) As Boolean, NameR(NumTopR) As String
Public NumTotalRN(NumTopR) As Boolean
Public Sub Mainload() '������
    '�о���
    NameR(0) = "�ڿ��۾�����"
    NameR(1) = "�����ı����й�������"
    NameR(2) = "��ͧ����"
    NameR(3) = "���������ֱ�"
    NameR(4) = "Ѽ�����������"
    NameR(5) = "�����խ���۾�����"
    NameR(6) = "����������ѡ������"
    NameR(7) = "��ͧ���ϸ���"
    NameR(8) = "�������������ֱ�"
    NameR(9) = "Ѽ���Ч������"
    NameR(10) = "���������ݽ���"
    NameR(11) = "������Ա�����Ὠ��"
    NameR(12) = "�������㳡����"
    '��Ʒ����
    ItemV(0) = 10
    ItemV(1) = 20
    ItemV(2) = 45
    ItemV(3) = 90
    ItemV(4) = 185
    '��Ʒ��
    NameI(0) = "�ڿ��۾� " & ItemV(0) & "s"
    NameI(1) = "�����ı����й��� " & ItemV(1) & "s"
    NameI(2) = "��ͨ������ͧ " & ItemV(2) & "s"
    NameI(3) = "��е�ֱ���װ " & ItemV(3) & "s"
    NameI(4) = "Ѽ�����װ " & ItemV(4) & "s"
    '�о�����
    ResV(0) = 20
    ResV(1) = 30
    ResV(2) = 60
    ResV(3) = 150
    ResV(4) = 260
    ResV(5) = 50
    ResV(6) = 85
    ResV(7) = 120
    ResV(8) = 325
    ResV(9) = 535
    ResV(10) = 100
    ResV(11) = 500
    ResV(12) = 1000
    '�о�ʱ��
    ResT(0) = 10
    ResT(1) = 30
    ResT(2) = 60
    ResT(3) = 115
    ResT(4) = 320
    ResT(5) = 35
    ResT(6) = 70
    ResT(7) = 125
    ResT(8) = 345
    ResT(9) = 600
    ResT(10) = 300
    ResT(11) = 1200
    ResT(12) = 3600
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
    Call ResShop
    sper = NumTotalS(0) * 1 * ItemPS(0)
    sper = sper + NumTotalS(1) * 2 * ItemPS(1)
    sper = sper + NumTotalS(2) * 5 * ItemPS(2)
    sper = sper + NumTotalS(3) * 10 * ItemPS(3)
    sper = sper + NumTotalS(3) * 20 * ItemPS(4)
    Main.Persec = "����1s��������:" & str(sper) & "s"
End Sub

Public Sub ResShop()
Dim IRS%
    '������
    If NumTotalR(0) Then ShopF.BuyI(0).Enabled = True
    If NumTotalR(1) Then ShopF.BuyI(1).Enabled = True
    If NumTotalR(2) Then ShopF.BuyI(2).Enabled = True
    If NumTotalR(3) Then ShopF.BuyI(3).Enabled = True
    If NumTotalR(4) Then ShopF.BuyI(4).Enabled = True
    '������
    If NumTotalR(0) And NumTotalR(5) And Not updPSed(0, 0) Then _
    NameI(0) = "�����խ���۾� " & ItemV(0) & "s": ItemPS(0) = 1.5: updPSed(0, 0) = True
    If NumTotalR(1) And NumTotalR(6) And Not updPSed(1, 0) Then _
    NameI(1) = "����������ѡ�� " & ItemV(1) & "s": ItemPS(1) = 1.5: updPSed(1, 0) = True
    If NumTotalR(2) And NumTotalR(7) And Not updPSed(2, 0) Then _
    NameI(2) = "���ϲ�����ͧ " & ItemV(2) & "s": ItemPS(2) = 1.5: updPSed(2, 0) = True
    If NumTotalR(3) And NumTotalR(8) And Not updPSed(3, 0) Then _
    NameI(3) = "�����ֱ���װ " & ItemV(3) & "s": ItemPS(3) = 1.5: updPSed(3, 0) = True
    If NumTotalR(4) And NumTotalR(9) And Not updPSed(4, 0) Then _
    NameI(4) = "��ЧѼ��� " & ItemV(4) & "s": ItemPS(4) = 1.5: updPSed(4, 0) = True
    For IRS = 0 To NumTopS
        ShopF.BuyI(IRS).Caption = NameI(IRS)
        ShopF.NumI(IRS) = "Ŀǰ��" & str(NumTotalS(IRS)) & "��"
    Next IRS
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
    Print #1, ClickP & "|" & ResHex & "|";
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
    '0-�û��� 1-�ܹ����� 2~(2+NumTopS)-��Ʒ���� 3+NumTopS-���������
    '4+NumTopS-ʮ�������о���(�о���+�о���+���о�) (5+NumTopS)~(5+NumTopS+NumTopR)-�о�ʣ��
    Main.User = stuffstr(0)
    Ts = stuffstr(1)
    For IL = 0 To NumTopS
        NumTotalS(IL) = stuffstr(IL + 2)
        If NumTotalR(IL + 5) Then updPSed(1, IL) = True
    Next IL
    ClickP = stuffstr(NumTopS + 3)
    bitR = Split(stuffstr(NumTopS + 4), "+", 3)
    Call bitBoo(hexBit(bitR(0)), NumTotalR())
    Call bitBoo(hexBit(bitR(1)), ResTIstuff())
    Call bitBoo(hexBit(bitR(2)), NumTotalRN())
    For IL = 0 To NumTopR
        ResTI(0, IL) = ResTIstuff(IL)
        ResTI(1, IL) = stuffstr(IL + NumTopS + 5)
        If NumTotalR(IL) Then updCed(IL) = True
    Next IL
    Call Refe
    Call ResRef
End Sub

Public Sub UpdEve(str$)
    Main.EventS = str & vbCrLf & Main.EventS
End Sub
