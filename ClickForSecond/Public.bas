Attribute VB_Name = "PublicFS"
Option Explicit
Public Const CFSVersion = "Beta1.3"
Public Const NumTopI = 7 - 1, NumTopR = 24 - 1, NumTopS = 2 - 1 '随物品种类/研究总数/技能总数变化
Public updCed(NumTopR) As Boolean, updPSed(NumTopI, 2 - 1) As Boolean, Shotlist(1, 9)
Public ItemV(NumTopI) As Double, ClickP As Integer, NameI(NumTopI) As String, ItemPS(NumTopI) As Double, ItemPST As Double
Public ResV(NumTopR) As Double, ResT(NumTopR) As Double, ResTI(1, NumTopR)
Public NumTotalI(NumTopI) As Double, sper As Double, chg%, NumTotalR(NumTopR) As Boolean, NameR(NumTopR) As String
Public NumTotalRN(NumTopR) As Boolean
Public NumSkill(NumTopS) As String
Public Sub Mainconst() '常数表(常量被注释表待定)
    '商品名
    NameI(0) = "黑框眼镜"
    NameI(1) = "《他改变了中国》"
    NameI(2) = "机械手表套装"
    NameI(3) = "普通鸭嘴笔"
    NameI(4) = "普通材料赛艇"
    NameI(5) = "《Aloha 'Oe》黑胶唱片"
    NameI(6) = "枸杞茶"
    '商品费用
    ItemV(0) = 10
    ItemV(1) = 20
    ItemV(2) = 45
    ItemV(3) = 90
    ItemV(4) = 210
    ItemV(5) = 480
    ItemV(6) = 59
    '研究名
    NameR(0) = "黑框眼镜制造"
    NameR(1) = "《他改变了中国》出版"
    NameR(2) = "整合三个手表"
    NameR(3) = "鸭嘴笔配套生产"
    NameR(4) = "赛艇制造"
    NameR(5) = "“《Aloha 'oe》黑胶唱片”出版"
    NameR(6) = "意大利窄边眼镜制造"
    NameR(7) = "《江泽民文选》出版"
    NameR(8) = "整合三个电子手表"
    NameR(9) = "鸭嘴笔效率升级"
    NameR(10) = "赛艇材料改良"
    NameR(11) = "磁带版唱片出版"
    NameR(12) = "枸杞茶上市"
    NameR(13) = "玳瑁框眼镜制造"
    NameR(14) = "《读懂江泽民》出版"
    NameR(15) = "便携原子钟整合"
    NameR(16) = "自动鸭嘴笔实装"
    NameR(17) = "赛艇外表面镀铪工作"
    NameR(18) = "DVD版唱片出版"
    NameR(19) = "工作区房屋建造"
    NameR(20) = "工作区员工宿舍建造"
    NameR(21) = "工作区广场建造"
    NameR(22) = "工作区工厂建造"
    NameR(23) = "在广场描绘时间法阵"
    '研究费用
    ResV(0) = 20
    ResV(1) = 30
    ResV(2) = 80
    ResV(3) = 150
    ResV(4) = 260
    ResV(5) = 555
    ResV(6) = 50
    ResV(7) = 85
    ResV(8) = 120
    ResV(9) = 325
    ResV(10) = 535
    ResV(11) = 1150
    ResV(12) = 275
    ResV(13) = 225
    ResV(14) = 340
    ResV(15) = 515
    ResV(16) = 1250
    ResV(17) = 2400
    ResV(18) = 5900
    ResV(19) = 100
    ResV(20) = 5000
    ResV(21) = 12000
    ResV(22) = 50000
    ResV(23) = 80000
    '研究时间
    ResT(0) = 10
    ResT(1) = 30
    ResT(2) = 60
    ResT(3) = 115
    ResT(4) = 320
    ResT(5) = 645
    ResT(6) = 35
    ResT(7) = 70
    ResT(8) = 125
    ResT(9) = 345
    ResT(10) = 600
    ResT(11) = 1215
    ResT(12) = 325
    ResT(13) = 155
    ResT(14) = 360
    ResT(15) = 645
    ResT(16) = 1550
    ResT(17) = 3600
    ResT(18) = 8500
    ResT(19) = 300
    ResT(20) = 1200
    ResT(21) = 3600
    ResT(22) = 43200
    ResT(23) = 6400
    '技能名
    NumSkill(0) = "喝枸杞茶"
End Sub

Public Sub Refe()
Dim IR%
    Main.Total = Ts
    For IR = 0 To NumTopI
        ShopF.NumI(IR) = "目前共" & NumTotalI(IR) & "个"
    Next IR
    Call NumPer
End Sub

Public Sub NumPer() '每秒增加量=对每个自动续命商品的(商品增益*商品效率)求和*商品总效率
    sper = NumTotalI(0) * 1 * ItemPS(0)
    sper = sper + NumTotalI(1) * 2 * ItemPS(1)
    sper = sper + NumTotalI(2) * 5 * ItemPS(2)
    sper = sper + NumTotalI(3) * 10 * ItemPS(3)
    sper = sper + NumTotalI(4) * 20 * ItemPS(4)
    sper = sper + NumTotalI(5) * 50 * ItemPS(5)
    sper = sper * ItemPST
    Main.Persec = "现在1s最少能续:" & str(sper) & "s"
End Sub

Public Sub ResShop()
Dim IRS%
    '解锁区
    If NumTotalR(0) Then ShopF.BuyI(0).Enabled = True
    If NumTotalR(1) Then ShopF.BuyI(1).Enabled = True
    If NumTotalR(2) Then ShopF.BuyI(2).Enabled = True
    If NumTotalR(3) Then ShopF.BuyI(3).Enabled = True
    If NumTotalR(4) Then ShopF.BuyI(4).Enabled = True
    If NumTotalR(5) Then ShopF.BuyI(5).Enabled = True
    If NumTotalR(12) Then ShopF.BuyI(6).Enabled = True
    '升级区
    If NumTotalR(0) And NumTotalR(6) And Not updPSed(0, 0) Then _
    NameI(0) = "意大利窄边眼镜": ItemPS(0) = 1.5: updPSed(0, 0) = True
    If NumTotalR(1) And NumTotalR(7) And Not updPSed(1, 0) Then _
    NameI(1) = "《江泽民文选》": ItemPS(1) = 1.5: updPSed(1, 0) = True
    If NumTotalR(2) And NumTotalR(8) And Not updPSed(2, 0) Then _
    NameI(2) = "电子手表套装": ItemPS(2) = 1.5: updPSed(2, 0) = True
    If NumTotalR(3) And NumTotalR(9) And Not updPSed(3, 0) Then _
    NameI(3) = "高效鸭嘴笔": ItemPS(3) = 1.5: updPSed(3, 0) = True
    If NumTotalR(4) And NumTotalR(10) And Not updPSed(4, 0) Then _
    NameI(4) = "复合材料赛艇": ItemPS(4) = 1.5: updPSed(4, 0) = True
    If NumTotalR(5) And NumTotalR(11) And Not updPSed(5, 0) Then _
    NameI(5) = "《Aloha 'Oe》磁带": ItemPS(5) = 1.5: updPSed(5, 0) = True
    If NumTotalR(0) And NumTotalR(13) And Not updPSed(0, 1) Then _
    NameI(0) = "玳瑁框眼镜": ItemPS(0) = 3: updPSed(0, 1) = True
    If NumTotalR(1) And NumTotalR(14) And Not updPSed(1, 1) Then _
    NameI(1) = "《读懂江泽民》": ItemPS(1) = 3: updPSed(1, 1) = True
    If NumTotalR(2) And NumTotalR(15) And Not updPSed(2, 1) Then _
    NameI(2) = "便携原子钟套装": ItemPS(2) = 3: updPSed(2, 1) = True
    If NumTotalR(3) And NumTotalR(16) And Not updPSed(3, 1) Then _
    NameI(3) = "自动鸭嘴笔": ItemPS(3) = 3: updPSed(3, 1) = True
    If NumTotalR(4) And NumTotalR(17) And Not updPSed(4, 1) Then _
    NameI(4) = "镀铪赛艇": ItemPS(4) = 3: updPSed(4, 1) = True
    If NumTotalR(5) And NumTotalR(18) And Not updPSed(5, 1) Then _
    NameI(5) = "《Aloha 'Oe》DVD": ItemPS(5) = 2.25: updPSed(5, 1) = True
    For IRS = 0 To NumTopI
        ShopF.BuyI(IRS).Caption = NameI(IRS) & str(ItemV(IRS) * (1 + NumTotalI(IRS) * 0.1)) & "s"
        ShopF.NumI(IRS) = "目前共" & NumTotalI(IRS) & "个"
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
    For ISa = 0 To NumTopI
        Print #1, NumTotalI(ISa) & "|";
    Next ISa
    Print #1, ClickP & "|" & ResHex & "|";
    For ISa = 0 To NumTopR
        Print #1, ResTI(1, ISa) & "|";
    Next ISa
    For ISa = 0 To NumTopI
        Print #1, ItemPS(ISa) & "|";
    Next ISa
    Print #1, CFSVersion & "|"
    Close #1
    UpdEve "已保存至'" & Main.Common.FileName & "'"
End Sub

Public Sub loadf()
Dim str As String, stuffstr, bitR, IL As Integer, ResTIstuff(NumTopR) As Boolean
    Main.Common.ShowOpen
    If Main.Common.FileName = "" Then Exit Sub Else
    Open Main.Common.FileName For Input As #1
    Line Input #1, str
    Close #1
    '载入文档处理
    stuffstr = Split(str, "|", -1)
    '0-用户名 1-总共秒数 2~(2+NumTopI)-物品数量 3+NumTopI-点击增加量
    '4+NumTopI-十六进制研究数(研究完+研究中+可研究) (5+NumTopI)~(5+NumTopI+NumTopR)-研究剩余 末尾-版本号(不使用)
    Main.User = stuffstr(0)
    Ts = stuffstr(1)
    For IL = 0 To NumTopI
        NumTotalI(IL) = stuffstr(IL + 2)
        If NumTotalR(IL + 5) Then updPSed(IL, 1) = True
    Next IL
    ClickP = stuffstr(NumTopI + 3)
    bitR = Split(stuffstr(NumTopI + 4), "+", 3)
    Call bitBoo(hexBit(bitR(0)), NumTotalR())
    Call bitBoo(hexBit(bitR(1)), ResTIstuff())
    Call bitBoo(hexBit(bitR(2)), NumTotalRN())
    For IL = 0 To NumTopR
        ResTI(0, IL) = ResTIstuff(IL)
        ResTI(1, IL) = stuffstr(IL + NumTopI + 5)
    Next IL
    If NumTotalR(6) Or NumTotalRN(6) Or ResTI(0, 6) Then updCed(6) = True
    If NumTotalR(7) Or NumTotalRN(7) Or ResTI(0, 7) Then updCed(7) = True
    If NumTotalR(8) Or NumTotalRN(8) Or ResTI(0, 8) Then updCed(8) = True
    If NumTotalR(9) Or NumTotalRN(9) Or ResTI(0, 9) Then updCed(9) = True
    If NumTotalR(10) Or NumTotalRN(10) Or ResTI(0, 10) Then updCed(10) = True
    If NumTotalR(11) Or NumTotalRN(11) Or ResTI(0, 11) Then updCed(11) = True
    If NumTotalR(13) Or NumTotalRN(13) Or ResTI(0, 13) Then updCed(13) = True
    If NumTotalR(14) Or NumTotalRN(14) Or ResTI(0, 14) Then updCed(14) = True
    If NumTotalR(15) Or NumTotalRN(15) Or ResTI(0, 15) Then updCed(15) = True
    If NumTotalR(16) Or NumTotalRN(16) Or ResTI(0, 16) Then updCed(16) = True
    If NumTotalR(17) Or NumTotalRN(17) Or ResTI(0, 17) Then updCed(17) = True
    If NumTotalR(18) Or NumTotalRN(18) Or ResTI(0, 18) Then updCed(18) = True
    If NumTotalR(19) Then updCed(19) = True: Call showWP(0)
    If NumTotalR(20) Then updCed(20) = True: Call showWP(1)
    If NumTotalR(21) Then updCed(21) = True: Call showWP(2)
    If NumTotalR(22) Then updCed(22) = True: Call showWP(2)
    Call Refe
    Call ResRef
    Call ResShop
    UpdEve "存档'" & Main.Common.FileName & "'已加载"
End Sub

Public Sub UpdEve(str$)
    Main.EventS = str & vbCrLf & Main.EventS
End Sub

Public Sub Shotadd(name As String, ts1 As Double)
Dim I%
    For I = 8 To 0 Step -1
        Shotlist(0, I + 1) = Shotlist(0, I)
        Shotlist(1, I + 1) = Shotlist(1, I)
    Next I
    Shotlist(0, 0) = name
    Shotlist(1, 0) = ts1
End Sub
