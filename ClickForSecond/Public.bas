Attribute VB_Name = "Public"
Option Explicit
Public Const NumTopS = 4 - 1, NumTopR = 5 - 1
Public ItemV(NumTopS) As Long, ClickP As Integer
Public ResV(NumTopR) As Integer, ResT(NumTopR) As Integer, ResTI(1, NumTopR)
Public NumTotalS(NumTopS) As Integer, sper&, chg%, NumTotalR(NumTopR) As Boolean, NameR(NumTopR) As String
Public NumTotalRN(NumTopR) As Boolean
Public Sub Mainload() '常数表
    '研究名
    NameR(0) = "黑框眼镜制造"
    NameR(1) = "《他改变了中国》出版"
    NameR(2) = "赛艇出厂"
    NameR(3) = "将三个手表整合起来"
    NameR(4) = "工作区房屋建造"
    '商品费用
    ItemV(0) = 10
    ItemV(1) = 20
    ItemV(2) = 45
    ItemV(3) = 90
    '研究费用
    ResV(0) = 20
    ResV(1) = 30
    ResV(2) = 60
    ResV(3) = 150
    ResV(4) = 100
    '研究时间
    ResT(0) = 10
    ResT(1) = 30
    ResT(2) = 60
    ResT(3) = 115
    ResT(4) = 70
End Sub
Public Sub Refe()
    Main.Total = str(Ts)
    For I = 0 To NumTopS
        ShopF.NumI(I) = "目前共" & str(NumTotalS(I)) & "个"
    Next I
    Call NumPer
End Sub

Public Sub NumPer()
    sper = NumTotalS(0) * 1
    sper = sper + NumTotalS(1) * 2
    sper = sper + NumTotalS(2) * 5
    sper = sper + NumTotalS(3) * 10
    Main.Persec = "现在1s最少能续:" & str(sper) & "s"
End Sub

Public Sub ResRef()
    ResearchF.Resing.Clear
    ResearchF.Resed.Clear
    ResearchF.Resable.Clear
    For I = 0 To NumTopR
        If ResTI(0, I) Then ResearchF.Resing.AddItem NameR(I)
        If NumTotalRN(I) Then ResearchF.Resable.AddItem NameR(I)
        If NumTotalR(I) Then ResearchF.Resed.AddItem NameR(I)
    Next I
End Sub

Public Sub ResShop()
    For I = 0 To NumTopS
        ShopF.NumI(I) = "目前共" & str(NumTotalS(I)) & "个"
    Next I
    If NumTotalR(0) Then ShopF.BuyI(0).Enabled = True
    If NumTotalR(1) Then ShopF.BuyI(1).Enabled = True
    If NumTotalR(2) Then ShopF.BuyI(2).Enabled = True
    If NumTotalR(3) Then ShopF.BuyI(3).Enabled = True
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
    '0-用户名 1-总共秒数 2~(4+2)-物品数量 4+3-点击增加量
    '4+4-十六进制研究数(研究情况+研究中) (4+5)~(4+5+5)-研究剩余
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
        & "消耗" & ResV(4) & "s" & ",研究时长" & ResT(4) & "s" & vbCrLf & "每次点击效率+1"
        Case Else: showde = "点击研究项目显示描述" & vbCrLf & "点击'研究'按钮以开始研究"
    End Select
End Function

Public Function ResNum(ind As String) As Integer
    For I = 0 To NumTopR
        If NameR(I) = ind Then ResNum = I: Exit Function
    Next I
    ResNum = -1
End Function
