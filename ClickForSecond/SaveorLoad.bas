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
    Main.Common.Filter = "保存文档(*.savesecond)|*.savesecond|全部文件(*.*)|*.*"
    Main.Common.ShowSave
    If Main.Common.FileName = "" Then Exit Sub Else
    Open Main.Common.FileName For Output As #1
    Print #1, Main.User & "|" & Ts & "|";
    For ISa = 0 To NumTopI
        Print #1, NumTotalI(ISa) & "|";
    Next ISa
    Print #1, ResHex & "|";
    For ISa = 0 To NumTopR
        Print #1, ResTI(1, ISa) & "|";
    Next ISa
    Print #1, OnlineTime & "|";
    Print #1, CFSVersion & "|";
    Close #1
    UpdEve "已保存至'" & Main.Common.FileName & "'"
End Sub

Public Sub loadF()
Dim str As String, stuffstr, bitR, iL As Integer, ResTIstuff() As Boolean
    ReDim ResTIstuff(NumTopR)
    Main.Common.Filter = "保存文档(*.savesecond)|*.savesecond"
    Main.Common.ShowOpen
    If Main.Common.FileName = "" Then Exit Sub Else
    Open Main.Common.FileName For Input As #1
    Line Input #1, str
    Close #1
    '载入文档处理
    stuffstr = Split(str, "|")
    '0-用户名 1-总共秒数 2~(2+NumTopI)-物品数量
    '3+NumTopI-十六进制研究数(研究完+研究中+可研究) (4+NumTopI)~(4+NumTopI+NumTopR)-研究剩余 5+NumTopI+NumTopR-总共在线时间
    '6+NumTopI+NumTopR-十六进制建筑数 (7+NumTopI+NumTopR)~(7+NumTopI+NumTopR+NumTopB)-建筑剩余(WIP)
    '末尾-版本号(不使用)
    Main.User = stuffstr(0)
    Ts = CDec(stuffstr(1))
    For iL = 0 To NumTopI
        NumTotalI(iL) = stuffstr(iL + 2)
    Next iL
    bitR = Split(stuffstr(NumTopI + 3), "+", 3)
    Call bitBoo(hexBit(bitR(0)), NumTotalR())
    Call bitBoo(hexBit(bitR(1)), ResTIstuff())
    Call bitBoo(hexBit(bitR(2)), NumTotalRN())
    For iL = 0 To NumTopR
        ResTI(0, iL) = ResTIstuff(iL)
        ResTI(1, iL) = stuffstr(iL + NumTopI + 4)
    Next iL
    OnlineTime = CDec(stuffstr(NumTopI + NumTopR + 5))
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
        str2 = Split(str1, "#", 2)
        str1 = str2(1)
        If str1 <> "" Then
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
    Call loadL
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
        str1 = Split(loadLang("Item.Crafting_" & I, LangA), "|")
        Crafting(0, I) = str1(0)
        Crafting(1, I) = str1(1)
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
    Open "MainOption.ini" For Output As #1
    Print #1, ConfigA
    Print #1, LangA
    Close #1
End Sub

