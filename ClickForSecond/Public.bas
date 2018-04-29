Attribute VB_Name = "Public"
Option Explicit
Public Const NumTopS = 4 - 1, NumTopR = 5 - 1
Public ResV(NumTopR) As Integer, ItemV(NumTopS) As Long, ResT(NumTopR) As Integer, ResTN(NumTopR) As Long, ClickP As Integer
Public NumTotalS(NumTopS) As Integer, sper&, chg%, NumTotalR(NumTopR) As Boolean, NameR(NumTopR) As String
Public NumTotalRN(NumTopR) As Boolean
Public Sub Mainload()
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
    ResearchF.Resable.Clear
    For I = 0 To NumTopR
        If NumTotalRN(I) Then ResearchF.Resable.AddItem NameR(I)
        If NumTotalR(I) Then ResearchF.Resed.AddItem NameR(I)
    Next I
End Sub

Public Sub saveF()
Dim ResHex As String
    ResHex = ResSave()
    Main.Common.DefaultExt = "savesecond"
    Main.Common.ShowSave
    If Main.Common.FileName = "" Then Exit Sub Else
    Open Main.Common.FileName For Output As #1
    Print #1, Main.User & "|" & str(Ts) & "|" _
    & NumTotalS(0) & "|" & NumTotalS(1) & "|" & NumTotalS(2) & "|" & NumTotalS(3) & "|" _
    & ResHex
    Close #1
End Sub

Public Sub loadf()
Dim str As String, stuffstr, bitL As String
    Main.Common.ShowOpen
    If Main.Common.FileName = "" Then Exit Sub Else
    Open Main.Common.FileName For Input As #1
    Line Input #1, str
    Close #1
    '�����ĵ�����
    stuffstr = Split(str, "|", -1)
    '0-�û��� 1-�ܹ����� 2~(4+2)-��Ʒ���� 4+3-��������� 4+4-ʮ�������о���
    Main.User = stuffstr(0)
    Ts = stuffstr(1)
    For I2 = 0 To NumTopS
    NumTotalS(I2) = stuffstr(I2 + 2)
    Next I2
    ClickP = stuffstr(NumTopS + 3)
    bitL = stuffstr(NumTopS + 4)
    Call bitBoo(hexBit(bitL), NumTotalR())
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
        If NameR(I) = ind Then ResNum = I
        Exit Function
    Next I
    ResNum = -1
End Function
