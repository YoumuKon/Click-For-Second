Attribute VB_Name = "SaveorLoad"
Option Explicit

Public Function ResSave() As String
Dim boo() As Integer, stuff As String, IRS%
    ReDim boo(NumTopR)
    ResSave = ""
    stuff = ""
    For IRS = 0 To NumTopR
        boo(IRS) = -(RO(IRS).Status = CFSisdone)
        ResSave = ResSave & boo(IRS)
    Next IRS
    For IRS = 0 To NumTopR
        boo(IRS) = -(RO(IRS).Status = CFSisdoing)
        stuff = stuff & boo(IRS)
    Next IRS
    ResSave = bitHex(ResSave) & "+" & bitHex(stuff)
    stuff = ""
    For IRS = 0 To NumTopR
        boo(IRS) = -(RO(IRS).Status = CFSisable)
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
        Print #1, RO(ISa).TimeNow & "|";
    Next ISa
    Print #1, OnlineTime & "|";
    Print #1, CFSVersion & "|";
    Close #1
    UpdEve "已保存至'" & Main.Common.FileName & "'"
End Sub

Public Sub loadF()
Dim str As String, stuffstr, bitR, I As Integer, ResStuff() As Boolean
    ReDim ResStuff(NumTopR)
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
    For I = 0 To NumTopI
        NumTotalI(iL) = stuffstr(iL + 2)
    Next I
    bitR = Split(stuffstr(NumTopI + 3), "+", 3)
    Call bitBoo(hexBit(bitR(0)), ResStuff())
    For I = 0 To NumTopR
        If ResStuff(I) Then RO(I).Status = CFSisdone
    Next I
    Call bitBoo(hexBit(bitR(1)), ResStuff())
    For I = 0 To NumTopR
        If ResStuff(I) Then RO(I).Status = CFSisdoing
    Next I
    Call bitBoo(hexBit(bitR(2)), ResStuff())
    For I = 0 To NumTopR
        If ResStuff(I) Then RO(I).Status = CFSisable
    Next I
    For I = 0 To NumTopR
        RO(I).TimeNow = stuffstr(iL + NumTopI + 4)
    Next I
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
        If str1 <> "" Then
            str2 = Split(str1, "#", 2)
            str1 = str2(0)
            If str1 <> "" Then
                str2 = Split(str1, "=", 2)
                If str2(0) = Name Then
                    loadLang = str2(1)
                    Exit Do
                End If
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
        With RO(I)
            .Name = loadLang("Research.name_" & I, LangA)
            .Description = loadLang("Research.tip_" & I, LangA)
            .Event = loadLang("Research.event_" & I, LangA)
            .Valve = loadLang("Research.cost_" & I, ConfigA)
            .Time = loadLang("Research.time_" & I, ConfigA)
        End With
        str1 = Split(loadLang("Research.needItem_" & I, ConfigA), "+")
        If UBound(str1) >= 1 Then
            For J = 0 To UBound(str1)
                str2 = Split(str1(J), "*")
                RO(I).NeedItem = RO(I).NeedItem & str2(0) & "|"
                RO(I).NeedItemNumber = RO(I).NeedItemNumber & str2(1) & "|"
            Next J
        End If
        I = I + 1
    Loop While loadLang("Research.name_" & I, LangA) <> ""
    NumTopR = I - 1
    ReDim Preserve RO(NumTopR)
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
    Loop While loadLang("Event.Normal_" & I, ConfigA) <> ""
    NumTopE = I - 1: ReDim Preserve EventList(NumTopE)
    I = 0
    Do
        NameB(0, I) = loadLang("Building.Name_" & I, LangA)
        NameB(1, I) = loadLang("Building.Tip_" & I, LangA)
        BuildV(I) = loadLang("Building.Cost_" & I, ConfigA)
        BuildT(I) = loadLang("Building.Time_" & I, ConfigA)
        str1 = Split(loadLang("Building.NeedItem_" & I, ConfigA), "+")
        If UBound(str1) >= 1 Then
            For J = 0 To UBound(str1)
                str2 = Split(str1(J), "*")
                BuildVI(0, I) = BuildVI(0, I) & str2(0) & "|"
                BuildVI(1, I) = BuildVI(1, I) & str2(1) & "|"
            Next J
        End If
        I = I + 1
    Loop While loadLang("Building.Time_" & I, ConfigA) <> ""
    NumTopB = I - 1
    ReDim Preserve NameB(1, NumTopB)
    ReDim Preserve BuildV(NumTopB)
    ReDim Preserve BuildT(NumTopB)
    ReDim Preserve BuildVI(1, NumTopB)
    I = 0
    Do
        Reventlist(0, I) = loadLang("Event.Random.Name_" & I, LangA)
        Reventlist(1, I) = loadLang("Event.Random.Tip_" & I, LangA)
        Reventlist(2, I) = loadLang("Event.Random.Probability_" & I, ConfigA)
        I = I + 1
    Loop While loadLang("Event.Random.Probability_" & I, ConfigA) <> ""
    NumTopRevent = I - 1: ReDim Preserve Reventlist(2, NumTopRevent)
    I = 0
    Do
        ItemV(I) = loadLang("Item.value_" & I, ConfigA)
        I = I + 1
    Loop While loadLang("Item.value_" & I, ConfigA) <> ""
    SellI = I - 1: ReDim Preserve ItemV(SellI)
    I = 0
    Do
        WPevent(I) = loadLang("WorkPlace_" & I, LangA)
        I = I + 1
    Loop While loadLang("WorkPlace_" & I, LangA) <> ""
    NumWPE = I - 1: ReDim Preserve WPevent(NumWPE)
    I = 0
    Do
        str1 = Split(loadLang("Item.Crafting_" & I, ConfigA), "|")
        Crafting(0, I) = str1(0)
        Crafting(1, I) = str1(1)
        I = I + 1
    Loop While loadLang("Item.Crafting_" & I, ConfigA) <> ""
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

