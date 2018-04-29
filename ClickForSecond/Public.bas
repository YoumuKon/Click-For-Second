Attribute VB_Name = "Public"
Option Explicit
Public Const NumTopS = 4 - 1, NumTopR = 6 - 1
Public ItemV(NumTopS) As Double, ClickP As Integer, NameI(NumTopS) As String, ItemPS(NumTopS) As Double
Public ResV(NumTopR) As Double, ResT(NumTopR) As Integer, ResTI(1, NumTopR)
Public NumTotalS(NumTopS) As Integer, sper As Double, chg%, NumTotalR(NumTopR) As Boolean, NameR(NumTopR) As String
Public NumTotalRN(NumTopR) As Boolean
Public Sub Mainload() '常数表
    '研究名
    NameR(0) = "黑框眼镜制造"
    NameR(1) = "《他改变了中国》出版"
    NameR(2) = "赛艇出厂"
    NameR(3) = "将三个手表整合起来"
    NameR(4) = "工作区房屋建造"
    NameR(5) = "意大利窄边眼镜制造"
    '商品费用
    ItemV(0) = 10
    ItemV(1) = 20
    ItemV(2) = 45
    ItemV(3) = 90
    '商品名
    NameI(0) = "黑框眼镜 " & ItemV(0) & "s"
    NameI(1) = "《他改变了中国》 " & ItemV(1) & "s"
    NameI(2) = "赛艇 " & ItemV(2) & "s"
    NameI(3) = "三手表套装 " & ItemV(3) & "s"
    '研究费用
    ResV(0) = 20
    ResV(1) = 30
    ResV(2) = 60
    ResV(3) = 150
    ResV(4) = 100
    ResV(5) = 50
    '研究时间
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
        ShopF.NumI(IR) = "目前共" & str(NumTotalS(IR)) & "个"
    Next IR
    Call NumPer
End Sub

Public Sub NumPer() '每秒增加量=对每个商品的(商品增益*商品效率)求和
    sper = NumTotalS(0) * 1 * ItemPS(0)
    sper = sper + NumTotalS(1) * 2 * ItemPS(1)
    sper = sper + NumTotalS(2) * 5 * ItemPS(2)
    sper = sper + NumTotalS(3) * 10 * ItemPS(3)
    Main.Persec = "现在1s最少能续:" & str(sper) & "s"
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
        ShopF.NumI(IRS) = "目前共" & str(NumTotalS(IRS)) & "个"
    Next IRS
    If NumTotalR(0) Then ShopF.BuyI(0).Enabled = True
    If NumTotalR(1) Then ShopF.BuyI(1).Enabled = True
    If NumTotalR(2) Then ShopF.BuyI(2).Enabled = True
    If NumTotalR(3) Then ShopF.BuyI(3).Enabled = True
    If NumTotalR(0) And NumTotalR(5) And Not updPSed(0, 0) Then _
    NameI(0) = "意大利窄边眼镜 " & ItemV(0) & "s": ItemPS(0) = 1.5: updPSed(0, 0) = False
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
    '载入文档处理
    stuffstr = Split(str, "|", -1)
    '0-用户名 1-总共秒数 2~(2+4)-物品数量 3+4-点击增加量
    '4+4-十六进制研究数(研究情况+研究中) (5+4)~(5+4+6)-研究剩余
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
        '添加新研究时直接粘贴以下两段
        Case NameR(0): showde = "允许商店售卖黑框眼镜." & vbCrLf _
        & "消耗" & ResV(0) & "s" & ",研究时长" & ResT(0) & "s"
        Case NameR(1): showde = "允许商店售卖《他改变了中国》." & vbCrLf _
        & "消耗" & ResV(1) & "s" & ",研究时长" & ResT(1) & "s"
        Case NameR(2): showde = "允许商店售卖赛艇." & vbCrLf _
        & "消耗" & ResV(2) & "s" & ",研究时长" & ResT(2) & "s"
        Case NameR(3): showde = "允许商店售卖三手表套装." & vbCrLf _
        & "消耗" & ResV(3) & "s" & ",研究时长" & ResT(3) & "s"
        Case NameR(4): showde = "为工作区添加一所房子以聘请工人." & vbCrLf _
        & "消耗" & ResV(4) & "s" & ",研究时长" & ResT(4) & "s" & vbCrLf & "每次点击效率为2"
        Case NameR(5): showde = "将黑框眼镜升级为意大利窄边眼镜." & vbCrLf _
        & "消耗" & ResV(5) & "s" & ",研究时长" & ResT(5) & "s" & vbCrLf & "黑框眼镜效率+50%"
        Case Else: showde = "点击研究项目显示描述" & vbCrLf & "点击'研究'按钮以开始研究"
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
        Case 4: Main.WorkPlace.Caption = "这是多了一个房子的工作区"
        Case Else: Main.WorkPlace.Caption = "这是工作区"
    End Select
End Sub

Public Sub UpdEve(str$)
    Main.EventS = str & vbCrLf & Main.EventS
End Sub
