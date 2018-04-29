Attribute VB_Name = "Public"
Option Explicit
Public Const NumTopS = 7 - 1, NumTopR = 16 - 1
Public updCed(NumTopR) As Boolean, updPSed(NumTopS, 1) As Boolean
Public ItemV(NumTopS) As Double, ClickP As Integer, NameI(NumTopS) As String, ItemPS(NumTopS) As Double
Public ResV(NumTopR) As Double, ResT(NumTopR) As Integer, ResTI(1, NumTopR)
Public NumTotalS(NumTopS) As Integer, sper As Double, chg%, NumTotalR(NumTopR) As Boolean, NameR(NumTopR) As String
Public NumTotalRN(NumTopR) As Boolean
Public Sub Mainconst() '常数表
    '商品费用
    ItemV(0) = 10
    ItemV(1) = 20
    ItemV(2) = 45
    ItemV(3) = 90
    ItemV(4) = 210
    ItemV(5) = 480
    ItemV(6) = 59
    '商品名
    NameI(0) = "黑框眼镜 " & ItemV(0) & "s"
    NameI(1) = "《他改变了中国》 " & ItemV(1) & "s"
    NameI(2) = "机械手表套装 " & ItemV(2) & "s"
    NameI(3) = "普通鸭嘴笔 " & ItemV(3) & "s"
    NameI(4) = "普通材料赛艇 " & ItemV(4) & "s"
    NameI(5) = "《Aloha 'Oe》黑胶唱片 " & ItemV(5) & "s"
    NameI(6) = "枸杞茶 " & ItemV(6) & "s"
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
    NameR(11) = "VCD版唱片出版"
    NameR(12) = "工作区房屋建造"
    NameR(13) = "工作区员工宿舍建造"
    NameR(14) = "工作区广场建造"
    NameR(15) = "枸杞茶上市"
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
    ResV(12) = 100
    ResV(13) = 5000
    ResV(14) = 12000
    ResV(15) = 275
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
    ResT(12) = 300
    ResT(13) = 1200
    ResT(14) = 3600
    ResT(15) = 325
End Sub

Public Sub Refe()
Dim IR%
    Main.Total = str(Ts)
    For IR = 0 To NumTopS
        ShopF.NumI(IR) = "目前共" & str(NumTotalS(IR)) & "个"
    Next IR
    Call NumPer
End Sub

Public Sub NumPer() '每秒增加量=对每个商品的(商品增益*商品效率)求和
    sper = NumTotalS(0) * 1 * ItemPS(0)
    sper = sper + NumTotalS(1) * 2 * ItemPS(1)
    sper = sper + NumTotalS(2) * 5 * ItemPS(2)
    sper = sper + NumTotalS(3) * 10 * ItemPS(3)
    sper = sper + NumTotalS(3) * 20 * ItemPS(4)
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
    If NumTotalR(15) Then ShopF.BuyI(6).Enabled = True
    '升级区
    If NumTotalR(0) And NumTotalR(6) And Not updPSed(0, 0) Then _
    NameI(0) = "意大利窄边眼镜 " & ItemV(0) & "s": ItemPS(0) = 1.5: updPSed(0, 0) = True
    If NumTotalR(1) And NumTotalR(7) And Not updPSed(1, 0) Then _
    NameI(1) = "《江泽民文选》 " & ItemV(1) & "s": ItemPS(1) = 1.5: updPSed(1, 0) = True
    If NumTotalR(2) And NumTotalR(8) And Not updPSed(2, 0) Then _
    NameI(2) = "电子手表套装 " & ItemV(2) & "s": ItemPS(2) = 1.5: updPSed(2, 0) = True
    If NumTotalR(3) And NumTotalR(9) And Not updPSed(3, 0) Then _
    NameI(3) = "高效鸭嘴笔 " & ItemV(3) & "s": ItemPS(3) = 1.5: updPSed(3, 0) = True
    If NumTotalR(4) And NumTotalR(10) And Not updPSed(4, 0) Then _
    NameI(4) = "复合材料赛艇 " & ItemV(4) & "s": ItemPS(4) = 1.5: updPSed(4, 0) = True
    If NumTotalR(5) And NumTotalR(11) And Not updPSed(5, 0) Then _
    NameI(5) = "《Aloha 'Oe》VCD版 " & ItemV(5) & "s": ItemPS(5) = 1.5: updPSed(5, 0) = True
    For IRS = 0 To NumTopS
        ShopF.BuyI(IRS).Caption = NameI(IRS)
        ShopF.NumI(IRS) = "目前共" & str(NumTotalS(IRS)) & "个"
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
    '载入文档处理
    stuffstr = Split(str, "|", -1)
    '0-用户名 1-总共秒数 2~(2+NumTopS)-物品数量 3+NumTopS-点击增加量
    '4+NumTopS-十六进制研究数(研究完+研究中+可研究) (5+NumTopS)~(5+NumTopS+NumTopR)-研究剩余
    Main.User = stuffstr(0)
    Ts = stuffstr(1)
    For IL = 0 To NumTopS
        NumTotalS(IL) = stuffstr(IL + 2)
        If NumTotalR(IL + 5) Then updPSed(IL, 1) = True
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
    If NumTotalR(10) Then Call showWP(0)
    If NumTotalR(11) Then Call showWP(1)
    If NumTotalR(12) Then Call showWP(2)
    Call Refe
    Call ResRef
End Sub

Public Sub UpdEve(str$)
    Main.EventS = str & vbCrLf & Main.EventS
End Sub
