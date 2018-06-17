Attribute VB_Name = "PublicFS"
Option Explicit
'一切东西的最大数量
Public Const MaxNum = 1926
'版本号/商品数量/工作区事件数
Public CFSVersion, SellI, NumWPE
'物品种类总数/研究总数/技能总数/建筑种类总数/随机事件总数/通用事件总数/合成物总数/研究关系式总数
Public NumTopI, NumTopR, NumTopS, NumTopB, NumTopRevent, NumTopE, NumTopC, NumTopRN
'研究解锁判定/用户名/效率升级判定/枪毙名单/技能解锁判定
Public updCed() As Boolean, UserN As String, updPSed() As Boolean, Shotlist(1, 9), updSkill() As Boolean
'商品费用/点击续命秒数/商品现名/商品各阶段名/商品效率/商品总效率
Public ItemV() As Double, ClickP As Integer, NameI() As String, NameII() As String, ItemPS() As Double, ItemPST As Double
'研究费用/研究时间/研究中+剩余时间/研究所需物品+数量
Public ResV() As Double, ResT() As Double, ResTI(), ResVI() As String
'物品数量/每秒续命秒数/通用事件总数/随机事件+提醒+概率/工作区提醒
Public NumTotalI() As Double, sper As Double, EventList() As String, Reventlist(), WPevent() As String
'研究解锁状态/研究完成状态/研究名+描述+特殊提醒/技能名+特殊提醒/合成表+合成物判定情况/合成成功概率/研究关系式
Public NumTotalRN() As Boolean, NumTotalR() As Boolean, NameR() As String, NameS(), Crafting(), CraftP As Double, ResNeed() As String
'建筑费用/建筑时间/建筑解锁状态/'建筑建成状态/建筑名+描述/建筑中+剩余时间/建筑所需物品+数量
Public BuildV() As Double, BuildT() As Double, NumTotalBN() As Boolean, NumTotalB() As Boolean, NameB(), BuildTI(), BuildVI()
Public ConfigA As String, LangA As String

Public Sub Refe()
Dim iR%
    Main.Total = Ts
    For iR = 0 To SellI
        ShopF.NumI(iR) = "目前共" & NumTotalI(iR) & "个"
    Next iR
    Call NumPer
End Sub

Public Sub NumPer() '每秒增加量=对每个自动续命商品的(商品增益*商品效率)求和*商品总效率
    sper = NumTotalI(0) * 1 * ItemPS(0)
    sper = sper + NumTotalI(1) * 2 * ItemPS(1)
    sper = sper + NumTotalI(2) * 5 * ItemPS(2)
    sper = sper + NumTotalI(3) * 10 * ItemPS(3)
    sper = sper + NumTotalI(4) * 20 * ItemPS(4)
    sper = sper + NumTotalI(5) * 50 * ItemPS(5)
    sper = sper + NumTotalI(6) * 100 * ItemPS(5)
    sper = sper * ItemPST
    Main.Persec = "现在1s最少能续:" & str(sper) & "s"
End Sub

Public Sub ResRefresh()
Dim IRS%
    '解锁区
    If NumTotalR(0) Then ShopF.BuyI(0).Enabled = True
    If NumTotalR(1) Then ShopF.BuyI(1).Enabled = True
    If NumTotalR(2) Then ShopF.BuyI(2).Enabled = True
    If NumTotalR(3) Then ShopF.BuyI(3).Enabled = True
    If NumTotalR(4) Then ShopF.BuyI(4).Enabled = True
    If NumTotalR(5) Then ShopF.BuyI(5).Enabled = True
    If NumTotalR(6) Then ShopF.BuyI(6).Enabled = True
    If NumTotalR(28) Then ShopF.BuyI(7).Enabled = True
    If NumTotalR(30) Then Main.NBuild.Enabled = True
    '升级区
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
        ShopF.NumI(IRS) = "目前共" & NumTotalI(IRS) & "个"
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
