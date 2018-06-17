Attribute VB_Name = "SaveorLoad"
Option Explicit

Public Function ResSave() As String
Dim boo() As Integer, stuff As String, IRS%
    ReDim boo(NumTopR)
    ResSave = ""
    stuff = ""
    For IRS = 0 To NumTopR
        boo(IRS) = -NumTotalR(IRS)
        ResSave = ResSave & boo(IRS)
    Next IRS
    For IRS = 0 To NumTopR
        boo(IRS) = -ResTI(0, IRS)
        stuff = stuff & boo(IRS)
    Next IRS
    ResSave = bitHex(ResSave) & "+" & bitHex(stuff)
    stuff = ""
    For IRS = 0 To NumTopR
        boo(IRS) = -NumTotalRN(IRS)
        stuff = stuff & boo(IRS)
    Next IRS
    ResSave = ResSave & "+" & bitHex(stuff)
End Function


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

Public Sub loadF()
Dim str As String, stuffstr, bitR, iL As Integer, ResTIstuff() As Boolean
    ReDim ResTIstuff(NumTopR)
    Main.Common.ShowOpen
    If Main.Common.FileName = "" Then Exit Sub Else
    Open Main.Common.FileName For Input As #1
    Line Input #1, str
    Close #1
    '载入文档处理
    stuffstr = Split(str, "|")
    '0-用户名 1-总共秒数 2~(2+NumTopI)-物品数量 3+NumTopI-点击增加量
    '4+NumTopI-十六进制研究数(研究完+研究中+可研究) (5+NumTopI)~(5+NumTopI+NumTopR)-研究剩余
    '6+NumTopI+NumTopR-十六进制建筑数 (7+NumTopI+NumTopR)~(7+NumTopI+NumTopR+NumTopB)-建筑剩余(功能未调试完毕)
    '末尾-版本号(不使用)
    Main.User = stuffstr(0)
    Ts = stuffstr(1)
    For iL = 0 To NumTopI
        NumTotalI(iL) = stuffstr(iL + 2)
    Next iL
    ClickP = stuffstr(NumTopI + 3)
    bitR = Split(stuffstr(NumTopI + 4), "+", 3)
    Call bitBoo(hexBit(bitR(0)), NumTotalR())
    Call bitBoo(hexBit(bitR(1)), ResTIstuff())
    Call bitBoo(hexBit(bitR(2)), NumTotalRN())
    For iL = 0 To NumTopR
        ResTI(0, iL) = ResTIstuff(iL)
        ResTI(1, iL) = stuffstr(iL + NumTopI + 5)
    Next iL
    If NumTotalR(7) Or NumTotalRN(7) Or ResTI(0, 7) Then updCed(7) = True
    If NumTotalR(8) Or NumTotalRN(8) Or ResTI(0, 8) Then updCed(8) = True
    If NumTotalR(9) Or NumTotalRN(9) Or ResTI(0, 9) Then updCed(9) = True
    If NumTotalR(10) Or NumTotalRN(10) Or ResTI(0, 10) Then updCed(10) = True
    If NumTotalR(11) Or NumTotalRN(11) Or ResTI(0, 11) Then updCed(11) = True
    If NumTotalR(12) Or NumTotalRN(12) Or ResTI(0, 12) Then updCed(12) = True
    If NumTotalR(13) Or NumTotalRN(13) Or ResTI(0, 13) Then updCed(13) = True
    If NumTotalR(14) Or NumTotalRN(14) Or ResTI(0, 14) Then updCed(14) = True
    If NumTotalR(15) Or NumTotalRN(15) Or ResTI(0, 15) Then updCed(15) = True
    If NumTotalR(16) Or NumTotalRN(16) Or ResTI(0, 16) Then updCed(16) = True
    If NumTotalR(17) Or NumTotalRN(17) Or ResTI(0, 17) Then updCed(17) = True
    If NumTotalR(18) Or NumTotalRN(18) Or ResTI(0, 18) Then updCed(18) = True
    If NumTotalR(19) Or NumTotalRN(19) Or ResTI(0, 19) Then updCed(19) = True
    If NumTotalR(20) Or NumTotalRN(20) Or ResTI(0, 20) Then updCed(20) = True
    If NumTotalR(23) Then Call showWP(0)
    If NumTotalR(24) Then Call showWP(1)
    If NumTotalR(25) Then Call showWP(2)
    If NumTotalR(26) Then Call showWP(3)
    Call Refe
    Call ResRef
    Call ResRefresh
    UpdEve "存档'" & Main.Common.FileName & "'已加载"
End Sub

Public Function loadLang(Name As String, FileA As String) As String
Dim str1 As String, str2
    loadLang = Empty
    Open FileA For Input As #1
    Do While Not EOF(1)
        Line Input #1, str1
        If Left(str1, 1) <> "#" And str1 <> "" Then
            str2 = Split(str1, "=", 2)
            If str2(0) = Name Then
                loadLang = str2(1)
                Exit Do
            End If
        End If
    Loop
    Close #1
End Function

Public Sub loadC()
Dim I%, str1 As String
    CFSVersion = loadLang("Version", ConfigA)
    CraftP = loadLang("CraftingProbability", ConfigA)
    '数组初始化
    ReDim updCed(MaxNum): ReDim updPSed(1, MaxNum): ReDim updSkill(MaxNum)
    ReDim NameI(MaxNum): ReDim NameII(2, MaxNum): ReDim ItemPS(MaxNum)
    ReDim ResV(MaxNum): ReDim ResT(MaxNum): ReDim ResTI(1, MaxNum): ReDim ResVI(1, MaxNum)
    ReDim NumTotalI(MaxNum): ReDim Reventlist(2, MaxNum): ReDim EventList(MaxNum): ReDim Reventlist(2, MaxNum)
    ReDim NumTotalRN(MaxNum): ReDim NumTotalR(MaxNum): ReDim NameR(2, MaxNum): ReDim NameS(1, MaxNum)
    ReDim BuildV(MaxNum): ReDim BuildT(MaxNum): ReDim NumTotalBN(MaxNum): ReDim NumTotalB(MaxNum): ReDim NameB(1, MaxNum): ReDim BuildTI(1, MaxNum): ReDim BuildVI(1, MaxNum)
    ReDim ResNeed(MaxNum): ReDim Crafting(1, MaxNum): ReDim WPevent(MaxNum): ReDim ItemV(MaxNum)
End Sub

Public Sub loadL()
Dim I%, J%, str1, str2
    I = 0
    Do
        For J = 0 To 2
        NameII(J, I) = loadLang("Item.name_" & I & "-" & J, LangA)
        NameI(I) = NameII(0, I)
        Next J
        I = I + 1
    Loop While loadLang("Item.name_" & I & "-0", LangA) <> ""
    NumTopI = I - 1: ReDim Preserve NameII(2, NumTopI): ReDim Preserve NameI(NumTopI)
    I = 0
    Do
        NameR(0, I) = loadLang("Research.name_" & I, LangA)
        NameR(1, I) = loadLang("Research.tip_" & I, LangA)
        NameR(2, I) = loadLang("Research.event_" & I, LangA)
        ResV(I) = loadLang("Research.cost_" & I, LangA)
        ResT(I) = loadLang("Research.time_" & I, LangA)
        str1 = Split(loadLang("Research.needItem_" & I, LangA), "+")
        If UBound(str1) >= 1 Then
            For J = 0 To UBound(str1)
                str2 = Split(str1(J), "*")
                ResVI(0, I) = ResVI(0, I) & str2(0) & "|"
                ResVI(1, I) = ResVI(1, I) & str2(1) & "|"
            Next J
        End If
        I = I + 1
    Loop While loadLang("Research.time_" & I, LangA) <> ""
    NumTopR = I - 1
    ReDim Preserve NameR(2, NumTopR)
    ReDim Preserve ResV(NumTopR)
    ReDim Preserve ResT(NumTopR)
    ReDim Preserve ResVI(1, NumTopR)
    I = 0
    Do
        NameS(0, I) = loadLang("Skill.name_" & I, LangA)
        NameS(1, I) = loadLang("Skill.event_" & I, LangA)
        I = I + 1
    Loop While loadLang("Skill.event_" & I, LangA) <> ""
    NumTopS = I - 1: ReDim Preserve NameS(1, NumTopS)
    I = 0
    Do
        EventList(I) = loadLang("Event.Normal_" & I, LangA)
        I = I + 1
    Loop While loadLang("Event.Normal_" & I, LangA) <> ""
    NumTopE = I - 1: ReDim Preserve EventList(NumTopE)
    I = 0
    Do
        NameB(0, I) = loadLang("Building.Name_" & I, LangA)
        NameB(1, I) = loadLang("Building.Tip_" & I, LangA)
        BuildV(I) = loadLang("Building.Cost_" & I, LangA)
        BuildT(I) = loadLang("Building.Time_" & I, LangA)
        str1 = Split(loadLang("Building.NeedItem_" & I, LangA), "+")
        If UBound(str1) >= 1 Then
            For J = 0 To UBound(str1)
                str2 = Split(str1(J), "*")
                BuildVI(0, I) = BuildVI(0, I) & str2(0) & "|"
                BuildVI(1, I) = BuildVI(1, I) & str2(1) & "|"
            Next J
        End If
        I = I + 1
    Loop While loadLang("Building.Time_" & I, LangA) <> ""
    NumTopB = I - 1
    ReDim Preserve NameB(1, NumTopB)
    ReDim Preserve BuildV(NumTopB)
    ReDim Preserve BuildT(NumTopB)
    ReDim Preserve BuildVI(1, NumTopB)
    I = 0
    Do
        Reventlist(0, I) = loadLang("Event.Random.Name_" & I, LangA)
        Reventlist(1, I) = loadLang("Event.Random.Tip_" & I, LangA)
        Reventlist(2, I) = loadLang("Event.Random.Probability_" & I, LangA)
        I = I + 1
    Loop While loadLang("Event.Random.Probability_" & I, LangA) <> ""
    NumTopRevent = I - 1: ReDim Preserve Reventlist(2, NumTopRevent)
    I = 0
    Do
        ItemV(I) = loadLang("Item.value_" & I, LangA)
        I = I + 1
    Loop While loadLang("Item.value_" & I, LangA) <> ""
    SellI = I - 1: ReDim Preserve ItemV(SellI)
    I = 0
    Do
        WPevent(I) = loadLang("WorkPlace_" & I, LangA)
        I = I + 1
    Loop While loadLang("WorkPlace_" & I, LangA) <> ""
    NumWPE = I - 1: ReDim Preserve WPevent(NumWPE)
    I = 0
    Do
        Crafting(0, I) = loadLang("Item.Crafting_" & I, LangA)
        I = I + 1
    Loop While loadLang("Item.Crafting_" & I, LangA) <> ""
    NumTopC = I - 1: ReDim Preserve Crafting(1, NumTopC)
    I = 0
    Do
        ResNeed(I) = loadLang("Research.Need_" & I, ConfigA)
        I = I + 1
    Loop While loadLang("Research.Need_" & I, ConfigA) <> ""
    NumTopRN = I - 1: ReDim Preserve ResNeed(NumTopRN)
End Sub

Public Sub mainload()
    Call loadC
    Call loadL
    Open "MainOption.ini" For Output As #1
    Print #1, ConfigA
    Print #1, LangA
    Close #1
End Sub

