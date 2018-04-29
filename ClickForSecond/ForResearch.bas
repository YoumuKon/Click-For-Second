Attribute VB_Name = "ForResearch"
Option Explicit
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

Public Function showde(ind As String) As String
    Select Case ind
        '添加新研究时直接粘贴
        Case NameR(0): showde = "允许商店售卖黑框眼镜." & vbCrLf _
        & "消耗" & ResV(0) & "s" & ",研究时长" & ResT(0) & "s"
        Case NameR(1): showde = "允许商店售卖《他改变了中国》." & vbCrLf _
        & "消耗" & ResV(1) & "s" & ",研究时长" & ResT(1) & "s"
        Case NameR(2): showde = "允许商店售卖机械手表套装." & vbCrLf _
        & "消耗" & ResV(2) & "s" & ",研究时长" & ResT(2) & "s"
        Case NameR(3): showde = "允许商店售卖普通鸭嘴笔." & vbCrLf _
        & "消耗" & ResV(3) & "s" & ",研究时长" & ResT(3) & "s"
        Case NameR(4): showde = "允许商店售卖赛艇." & vbCrLf _
        & "消耗" & ResV(4) & "s" & ",研究时长" & ResT(4) & "s"
        Case NameR(5): showde = "允许商店售卖《Aloha 'Oe》黑胶唱片." & vbCrLf _
        & "消耗" & ResV(5) & "s" & ",研究时长" & ResT(5) & "s"
        Case NameR(6): showde = "将黑框眼镜升级为意大利窄边眼镜." & vbCrLf _
        & "消耗" & ResV(6) & "s" & ",研究时长" & ResT(6) & "s" & vbCrLf & "效率+50%"
        Case NameR(7): showde = "将《他改变了中国》升级为《江泽民文选》." & vbCrLf _
        & "消耗" & ResV(7) & "s" & ",研究时长" & ResT(7) & "s" & vbCrLf & "效率+50%"
        Case NameR(8): showde = "将整合中的手表升级为电子手表." & vbCrLf _
        & "消耗" & ResV(8) & "s" & ",研究时长" & ResT(8) & "s" & vbCrLf & "效率+50%"
        Case NameR(9): showde = "将鸭嘴笔升级为高效鸭嘴笔." & vbCrLf _
        & "消耗" & ResV(9) & "s" & ",研究时长" & ResT(9) & "s" & vbCrLf & "效率+50%"
        Case NameR(10): showde = "将赛艇的材料升级为复合材料." & vbCrLf _
        & "消耗" & ResV(10) & "s" & ",研究时长" & ResT(10) & "s" & vbCrLf & "效率+50%"
        Case NameR(11): showde = "将黑胶唱片刻录到VCD." & vbCrLf _
        & "消耗" & ResV(11) & "s" & ",研究时长" & ResT(11) & "s" & vbCrLf & "效率+50%"
        Case NameR(12): showde = "为工作区建造一所房子以聘请工人." & vbCrLf _
        & "消耗" & ResV(12) & "s" & ",研究时长" & ResT(12) & "s" & vbCrLf & "每次点击效率+1"
        Case NameR(13): showde = "为工作区的房子升级为员工宿舍以聘请更多工人." & vbCrLf _
        & "消耗" & ResV(13) & "s" & ",研究时长" & ResT(13) & "s" & vbCrLf & "每次点击效率+3"
        Case NameR(14): showde = "为工作区添加一个广场以提高工人积极性." & vbCrLf _
        & "消耗" & ResV(14) & "s" & ",研究时长" & ResT(14) & "s" & vbCrLf & "每次点击效率+40%"
        Case NameR(15): showde = "允许商店售卖枸杞茶." & vbCrLf _
        & "消耗" & ResV(15) & "s" & ",研究时长" & ResT(15) & "s"
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
        Case 0: Main.WorkPlace.Caption = "这是多了一个房子的工作区"
        Case 1: Main.WorkPlace.Caption = "这是有员工宿舍配套的工作区"
        Case 2: Main.WorkPlace.Caption = "这是有员工宿舍配套的工作区广场"
        Case Else: Main.WorkPlace.Caption = "这是工作区"
    End Select
End Sub

Public Sub CheckRes()
Dim updateR As Boolean
    updateR = False
    '目前判定：
    '买够10个商品0时解锁研究1和6
    If NumTotalS(0) >= 10 And NumTotalR(0) And Not updCed(0) Then _
    NumTotalRN(6) = True: NumTotalRN(1) = True: updCed(0) = True: updateR = True
    '买够10个商品1时解锁研究2和7
    If NumTotalS(1) >= 10 And NumTotalR(1) And Not updCed(1) Then _
    NumTotalRN(7) = True: NumTotalRN(2) = True: updCed(1) = True: updateR = True
    '买够10个商品2时解锁研究3, 15和8
    If NumTotalS(2) >= 10 And NumTotalR(2) And Not updCed(2) Then _
    NumTotalRN(8) = True: NumTotalRN(15) = True: NumTotalRN(3) = True: updCed(2) = True: updateR = True
    '买够10个商品3时解锁研究4和9
    If NumTotalS(3) >= 10 And NumTotalR(3) And Not updCed(3) Then _
    NumTotalRN(9) = True: NumTotalRN(4) = True: updCed(3) = True: updateR = True
    '买够10个商品4时解锁研究5和10
    If NumTotalS(4) >= 10 And NumTotalR(4) And Not updCed(4) Then _
    NumTotalRN(10) = True: NumTotalRN(5) = True: updCed(4) = True: updateR = True
    '买够10个商品5时解锁研究11
    If NumTotalS(5) >= 10 And NumTotalR(5) And Not updCed(5) Then _
    NumTotalRN(11) = True: updCed(5) = True: updateR = True
    '研究12,13,14完成后更新工作区信息
    If NumTotalR(12) And Not updCed(12) Then _
    ClickP = ClickP + 1: updCed(12) = True: NumTotalRN(13) = True: Call showWP(0): updateR = True
    If NumTotalR(13) And Not updCed(13) Then _
    ClickP = ClickP + 3: updCed(13) = True: NumTotalRN(14) = True: Call showWP(1): updateR = True
    If NumTotalR(14) And Not updCed(14) Then _
    ClickP = ClickP * 1.4: updCed(14) = True: Call showWP(2): updateR = True
    '更新研究列表
    If updateR Then Call ResRef: Call ResShop
End Sub

