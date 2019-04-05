Attribute VB_Name = "PublicFS"
'�Զ�������
'�о�����Ʒ:(��Ʒ-����-�۸�-ʱ��-�¼�-������Ʒ��Ϣ-����״̬)
Public Type ResearchObject
    Name As String
    Description As String
    Valve As Long
    Time As Long
    Event As String
    NeedItem As Variant
    NeedItemNumber As Variant
    TimeNow As Long
    Status As Statuinfo
End Type
'�о�״̬ö��
Enum Statuinfo
    CFSisdone = -1
    CFSisdoing = 1
    CFSisable = 0
    CFSnone = -2
End Enum
'һ�ж������������
Public Const MaxNum = 1926
'������
Public Const StrMem1 = "&Mem1", StrMem2 = "&Mem2", StrUser = "&U", StrCrLf = "&CL"
'�汾��/��Ʒ����/�������¼���/����ʱ��
Public CFSVersion, SellI, NumWPE, OnlineTime
'��Ʒ��������/�о�����/��������/������������/����¼�����/ͨ���¼�����/�ϳ�������/�о���ϵʽ����/�Զ���Ʒ��
Public NumTopI, NumTopR, NumTopS, NumTopB, NumTopRevent, NumTopE, NumTopC, NumTopRN, NumTopAuto
'�о������ж�/�û���/Ч�������ж�/ǹ������/���ܽ����ж�
Public updCed() As Boolean, UserN As String, updPSed() As Boolean, Shotlist(1, 9), updSkill() As Boolean
'��Ʒ����/�����������/��Ʒ����/��Ʒ���׶���/��ƷЧ��/��Ʒ��Ч��
Public ItemV() As Double, ClickP As Integer, NameI() As String, NameII() As String, ItemPS() As Double, ItemPST As Double
'�о��������
Public RO() As ResearchObject
'��Ʒ����/ÿ����������/ͨ���¼�����/����¼�+����+����/����������
Public NumTotalI(), sper As Double, EventList() As String, Reventlist(), WPevent() As String
'������+��������/�ϳɱ�+�ϳ����ж����/�ϳɳɹ�����/�о���ϵʽ
Public NameS(), Crafting(), CraftP As Double, ResNeed() As String
'��������/����ʱ��/��������״̬/'��������״̬/������+����/������+ʣ��ʱ��/����������Ʒ+����
Public BuildV() As Double, BuildT() As Double, NumTotalBN() As Boolean, NumTotalB() As Boolean, NameB(), BuildTI(), BuildVI(), BO() As ResearchObject
'���������ļ���ַ
Public ConfigA As String, LangA As String
'FileSystemObject����
Public fs

Public Sub Errlog(ErrN As Integer)
Dim lname, Logf, Datename
    Datename = Format(Date, "yyyy-mm-dd")
    lname = App.Path & "\log\" & Datename & ".log"
    If Dir(App.Path & "\log", vbDirectory) = "" Then Set Logf = fs.CreateTextFile(lname)
    Set Logf = fs.OpenTextFile(lname, ForAppending, True)
    Logf.writeline "��������" & Now & "��������"
    Logf.writeline "������ţ�" & Err.Number
    Logf.writeline "����������" & Err.Description
    Logf.writeline "���´��󴦣�" & Err.Source
    Logf.writeline "----------------------------------------------"
    Logf.Close
    MsgBox "�������!" & vbCrLf & "��鿴λ��" & lname & "�ı��棬��������������Ա��", vbCritical, "�������"
    Stop
End Sub

Public Sub Rediming(ind As Byte)
    Select Case ind
        Case 0: '�����ʼ��
        ReDim updCed(MaxNum): ReDim updPSed(2, MaxNum): ReDim updSkill(MaxNum): ReDim NameS(1, MaxNum)
        ReDim NameI(MaxNum): ReDim NameII(2, MaxNum): ReDim ItemPS(MaxNum)
        ReDim NumTotalI(MaxNum): ReDim Reventlist(2, MaxNum): ReDim EventList(MaxNum): ReDim Reventlist(2, MaxNum)
        ReDim BuildV(MaxNum): ReDim BuildT(MaxNum): ReDim NumTotalBN(MaxNum): ReDim NumTotalB(MaxNum): ReDim NameB(1, MaxNum): ReDim BuildTI(1, MaxNum): ReDim BuildVI(1, MaxNum)
        ReDim ResNeed(MaxNum): ReDim Crafting(1, MaxNum): ReDim WPevent(MaxNum): ReDim ItemV(MaxNum): ReDim RO(MaxNum): ReDim BO(MaxNum)
        Case 1:
    End Select
End Sub

Public Sub Refe()
Dim iR%
    Main.Total = Ts
    For iR = 0 To SellI
        ShopF.NumI(iR) = "Ŀǰ��" & NumTotalI(iR) & "��"
    Next iR
    Call NumPer
End Sub

Public Sub NumPer() 'ÿ��������=��ÿ���Զ�������Ʒ��(��Ʒ����*��ƷЧ��)���*��Ʒ��Ч��
    sper = NumTotalI(0) * 1 * ItemPS(0)
    sper = sper + NumTotalI(1) * 2 * ItemPS(1)
    sper = sper + NumTotalI(2) * 5 * ItemPS(2)
    sper = sper + NumTotalI(3) * 10 * ItemPS(3)
    sper = sper + NumTotalI(4) * 20 * ItemPS(4)
    sper = sper + NumTotalI(5) * 50 * ItemPS(5)
    sper = sper + NumTotalI(6) * 100 * ItemPS(5)
    sper = sper * ItemPST
    Main.Persec = "����1s��������:" & str(sper) & "s"
End Sub

Public Sub ResRefresh()
    '����
    If RO(0).Status = CFSisdone Then ShopF.BuyI(0).Enabled = True
    If RO(1).Status = CFSisdone Then ShopF.BuyI(1).Enabled = True
    If RO(2).Status = CFSisdone Then ShopF.BuyI(2).Enabled = True
    If RO(3).Status = CFSisdone Then ShopF.BuyI(3).Enabled = True
    If RO(4).Status = CFSisdone Then ShopF.BuyI(4).Enabled = True
    If RO(5).Status = CFSisdone Then ShopF.BuyI(5).Enabled = True
    If RO(6).Status = CFSisdone Then ShopF.BuyI(6).Enabled = True
    If RO(28).Status = CFSisdone Then ShopF.BuyI(7).Enabled = True
    If RO(29).Status = CFSisdone Then Main.NBuild.Enabled = True
    '����
    Call refshop
    Call CraftingF.RefCraft
End Sub

Public Sub refshop()
Dim I%
    For I = 0 To SellI
        ShopF.BuyI(I).Caption = NameI(I) & str(ItemV(I) * (1 + NumTotalI(I) * (0.25 * ItemPS(I)))) & "s"
        ShopF.NumI(I) = "Ŀǰ��" & NumTotalI(I) & "��"
    Next I
End Sub

Public Sub UpdEve(str As String)
    Main.EventS = Main.EventS & str & vbCrLf
    Main.EventS.SelStart = Len(Main.EventS.Text)
End Sub

Public Sub Shotadd(Name As String, ts1 As Double)
Dim I%
    For I = 8 To 0 Step -1
        Shotlist(0, I + 1) = Shotlist(0, I)
        Shotlist(1, I + 1) = Shotlist(1, I)
    Next I
    Shotlist(0, 0) = Name
    Shotlist(1, 0) = ts1
End Sub

Public Function StrEnc(ByVal str1 As String, ByVal str2 As String, ByVal Mem As String)
Dim inS As String, strS
    inS = InStr(str1, str2)
    If inS <> 0 Then
        strS = Split(str1, str2, 2)
        If UBound(strS) > 0 Then StrEnc = strS(0) & Mem & strS(1)
        Else: StrEnc = str1
    End If
End Function

Public Function CraftNum(Inp As String)
Dim I%
    CraftNum = -1
    For I = 0 To NumTopC
        If NameII(0, I + SellI + 1) = Inp Then CraftNum = I: Exit Function
    Next I
End Function
