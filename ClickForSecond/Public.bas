Attribute VB_Name = "PublicFS"
'自定义类型
'研究类物品:(物品-描述-价格-时间-事件-需求物品信息-现在状态)
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
'研究状态枚举
Enum Statuinfo
    CFSisdone = -1
    CFSisdoing = 1
    CFSisable = 0
    CFSnone = -2
End Enum
'一切东西的最大数量
Public Const MaxNum = 1926
'各常量
Public Const StrMem1 = "&Mem1", StrMem2 = "&Mem2", StrUser = "&U", StrCrLf = "&CL"
'版本号/商品数量/工作区事件数/在线时间
Public CFSVersion, SellI, NumWPE, OnlineTime
'物品种类总数/研究总数/技能总数/建筑种类总数/随机事件总数/通用事件总数/合成物总数/研究关系式总数/自动物品数
Public NumTopI, NumTopR, NumTopS, NumTopB, NumTopRevent, NumTopE, NumTopC, NumTopRN, NumTopAuto
'研究解锁判定/用户名/效率升级判定/枪毙名单/技能解锁判定
Public updCed() As Boolean, UserN As String, updPSed() As Boolean, Shotlist(1, 9), updSkill() As Boolean
'商品费用/点击续命秒数/商品现名/商品各阶段名/商品效率/商品总效率
Public ItemV() As Double, ClickP As Integer, NameI() As String, NameII() As String, ItemPS() As Double, ItemPST As Double
'研究相关数据
Public RO() As ResearchObject
'物品数量/每秒续命秒数/通用事件总数/随机事件+提醒+概率/工作区提醒
Public NumTotalI(), sper As Double, EventList() As String, Reventlist(), WPevent() As String
'技能名+特殊提醒/合成表+合成物判定情况/合成成功概率/研究关系式
Public NameS(), Crafting(), CraftP As Double, ResNeed() As String
'建筑费用/建筑时间/建筑解锁状态/'建筑建成状态/建筑名+描述/建筑中+剩余时间/建筑所需物品+数量
Public BuildV() As Double, BuildT() As Double, NumTotalBN() As Boolean, NumTotalB() As Boolean, NameB(), BuildTI(), BuildVI(), BO() As ResearchObject
'各大配置文件地址
Public ConfigA As String, LangA As String
'FileSystemObject定义
Public fs

Public Sub Errlog(ErrN As Integer)
Dim lname, Logf, Datename
    Datename = Format(Date, "yyyy-mm-dd")
    lname = App.Path & "\log\" & Datename & ".log"
    If Dir(App.Path & "\log", vbDirectory) = "" Then Set Logf = fs.CreateTextFile(lname)
    Set Logf = fs.OpenTextFile(lname, ForAppending, True)
    Logf.writeline "本程序于" & Now & "发生错误："
    Logf.writeline "错误代号：" & Err.Number
    Logf.writeline "错误描述：" & Err.Description
    Logf.writeline "大致错误处：" & Err.Source
    Logf.writeline "----------------------------------------------"
    Logf.Close
    MsgBox "程序崩溃!" & vbCrLf & "请查看位于" & lname & "的报告，并反馈给开发人员。", vbCritical, "程序崩溃"
    Stop
End Sub

Public Sub Rediming(ind As Byte)
    Select Case ind
        Case 0: '数组初始化
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
    '解锁
    If RO(0).Status = CFSisdone Then ShopF.BuyI(0).Enabled = True
    If RO(1).Status = CFSisdone Then ShopF.BuyI(1).Enabled = True
    If RO(2).Status = CFSisdone Then ShopF.BuyI(2).Enabled = True
    If RO(3).Status = CFSisdone Then ShopF.BuyI(3).Enabled = True
    If RO(4).Status = CFSisdone Then ShopF.BuyI(4).Enabled = True
    If RO(5).Status = CFSisdone Then ShopF.BuyI(5).Enabled = True
    If RO(6).Status = CFSisdone Then ShopF.BuyI(6).Enabled = True
    If RO(28).Status = CFSisdone Then ShopF.BuyI(7).Enabled = True
    If RO(29).Status = CFSisdone Then Main.NBuild.Enabled = True
    '升级
    Call refshop
    Call CraftingF.RefCraft
End Sub

Public Sub refshop()
Dim I%
    For I = 0 To SellI
        ShopF.BuyI(I).Caption = NameI(I) & str(ItemV(I) * (1 + NumTotalI(I) * (0.25 * ItemPS(I)))) & "s"
        ShopF.NumI(I) = "目前共" & NumTotalI(I) & "个"
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
