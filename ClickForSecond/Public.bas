Attribute VB_Name = "PublicFS"
Option Explicit
'һ�ж������������
Public Const MaxNum = 1926
'�汾��/��Ʒ����/�������¼���
Public CFSVersion, SellI, NumWPE
'��Ʒ��������/�о�����/��������/������������/����¼�����/ͨ���¼�����/�ϳ�������/�о���ϵʽ����
Public NumTopI, NumTopR, NumTopS, NumTopB, NumTopRevent, NumTopE, NumTopC, NumTopRN
'�о������ж�/�û���/Ч�������ж�/ǹ������/���ܽ����ж�
Public updCed() As Boolean, UserN As String, updPSed() As Boolean, Shotlist(1, 9), updSkill() As Boolean
'��Ʒ����/�����������/��Ʒ����/��Ʒ���׶���/��ƷЧ��/��Ʒ��Ч��
Public ItemV() As Double, ClickP As Integer, NameI() As String, NameII() As String, ItemPS() As Double, ItemPST As Double
'�о�����/�о�ʱ��/�о���+ʣ��ʱ��/�о�������Ʒ+����
Public ResV() As Double, ResT() As Double, ResTI(), ResVI() As String
'��Ʒ����/ÿ����������/ͨ���¼�����/����¼�+����+����/����������
Public NumTotalI() As Double, sper As Double, EventList() As String, Reventlist(), WPevent() As String
'�о�����״̬/�о����״̬/�о���+����+��������/������+��������/�ϳɱ�+�ϳ����ж����/�ϳɳɹ�����/�о���ϵʽ
Public NumTotalRN() As Boolean, NumTotalR() As Boolean, NameR() As String, NameS(), Crafting(), CraftP As Double, ResNeed() As String
'��������/����ʱ��/��������״̬/'��������״̬/������+����/������+ʣ��ʱ��/����������Ʒ+����
Public BuildV() As Double, BuildT() As Double, NumTotalBN() As Boolean, NumTotalB() As Boolean, NameB(), BuildTI(), BuildVI()
Public ConfigA As String, LangA As String

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
Dim IRS%
    '������
    If NumTotalR(0) Then ShopF.BuyI(0).Enabled = True
    If NumTotalR(1) Then ShopF.BuyI(1).Enabled = True
    If NumTotalR(2) Then ShopF.BuyI(2).Enabled = True
    If NumTotalR(3) Then ShopF.BuyI(3).Enabled = True
    If NumTotalR(4) Then ShopF.BuyI(4).Enabled = True
    If NumTotalR(5) Then ShopF.BuyI(5).Enabled = True
    If NumTotalR(6) Then ShopF.BuyI(6).Enabled = True
    If NumTotalR(28) Then ShopF.BuyI(7).Enabled = True
    If NumTotalR(30) Then Main.NBuild.Enabled = True
    '������
    If NumTotalR(0) And NumTotalR(7) And Not updPSed(0, 0) Then _
    NameI(0) = NameII(1, 0): ItemPS(0) = 1.5: updPSed(0, 0) = True
    If NumTotalR(1) And NumTotalR(8) And Not updPSed(0, 1) Then _
    NameI(1) = NameII(1, 1): ItemPS(1) = 1.5: updPSed(0, 1) = True
    If NumTotalR(2) And NumTotalR(9) And Not updPSed(0, 2) Then _
    NameI(2) = NameII(1, 2): ItemPS(2) = 1.5: updPSed(0, 2) = True
    If NumTotalR(3) And NumTotalR(10) And Not updPSed(0, 3) Then _
    NameI(3) = NameII(1, 3): ItemPS(3) = 1.5: updPSed(0, 3) = True
    If NumTotalR(4) And NumTotalR(11) And Not updPSed(0, 4) Then _
    NameI(4) = NameII(1, 4): ItemPS(4) = 1.5: updPSed(0, 4) = True
    If NumTotalR(5) And NumTotalR(12) And Not updPSed(0, 5) Then _
    NameI(5) = NameII(1, 5): ItemPS(5) = 1.5: updPSed(0, 5) = True
    If NumTotalR(6) And NumTotalR(13) And Not updPSed(0, 6) Then _
    NameI(6) = NameII(1, 6): ItemPS(6) = 1.5: updPSed(0, 6) = True
    If NumTotalR(0) And NumTotalR(13) And Not updPSed(1, 0) Then _
    NameI(0) = NameII(2, 0): ItemPS(0) = 2.25: updPSed(1, 0) = True
    If NumTotalR(1) And NumTotalR(14) And Not updPSed(1, 1) Then _
    NameI(1) = NameII(2, 1): ItemPS(1) = 2.25: updPSed(1, 1) = True
    If NumTotalR(2) And NumTotalR(15) And Not updPSed(1, 2) Then _
    NameI(2) = NameII(2, 2): ItemPS(2) = 2.25: updPSed(1, 2) = True
    If NumTotalR(3) And NumTotalR(16) And Not updPSed(1, 3) Then _
    NameI(3) = NameII(2, 3): ItemPS(3) = 2.25: updPSed(1, 3) = True
    If NumTotalR(4) And NumTotalR(17) And Not updPSed(1, 4) Then _
    NameI(4) = NameII(2, 4): ItemPS(4) = 2.25: updPSed(1, 4) = True
    If NumTotalR(5) And NumTotalR(18) And Not updPSed(1, 5) Then _
    NameI(5) = NameII(2, 5): ItemPS(5) = 2.25: updPSed(1, 5) = True
    If NumTotalR(5) And NumTotalR(19) And Not updPSed(1, 6) Then _
    NameI(5) = NameII(2, 6): ItemPS(6) = 2.25: updPSed(1, 6) = True
    For IRS = 0 To SellI
        ShopF.BuyI(IRS).Caption = NameI(IRS) & str(ItemV(IRS) * (1 + NumTotalI(IRS) * (0.25 * ItemPS(IRS)))) & "s"
        ShopF.NumI(IRS) = "Ŀǰ��" & NumTotalI(IRS) & "��"
    Next IRS
End Sub

Public Sub UpdEve(str$)
    Main.EventS = str & vbCrLf & Main.EventS
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

Public Sub ResCraft()
Dim I%
    For I = 0 To NumTopC
        If Crafting(1, I) Then CraftingF.CraftList.AddItem NameII(0, I + SellI + 1)
    Next I
End Sub
