Attribute VB_Name = "ForResearch"
Option Explicit
Public Sub ResRef()
Dim IRR%
    ResearchF.Resing.Clear
    ResearchF.Resed.Clear
    ResearchF.Resable.Clear
    For IRR = 0 To NumTopR
        If ResTI(0, IRR) Then ResearchF.Resing.AddItem NameR(0, IRR)
        If NumTotalRN(IRR) Then ResearchF.Resable.AddItem NameR(0, IRR)
        If NumTotalR(IRR) Then ResearchF.Resed.AddItem NameR(0, IRR)
    Next IRR
End Sub

Public Function ResNum(ind As String) As Integer
Dim IRM%
    ResNum = -1
    For IRM = 0 To NumTopR
        If NameR(0, IRM) = ind Then ResNum = IRM: Exit Function
    Next IRM
End Function

Public Sub showWP(ind As Integer)
    Select Case ind
        Case 0: ClickP = 2 '1+1=2
        Case 1: ClickP = 5 '2+3=5
        Case 2: ClickP = 7 '5*1.4=7
        Case 3: ItemPST = 1.1 '1+0.1=1.1
    End Select
    Main.WorkPlace.Caption = WPevent(ind + 1)
End Sub

Public Function checkWP() As Integer
Dim I
    For I = 0 To NumWPE
        If Main.WorkPlace.Caption = WPevent(I) Then checkWP = I - 1
    Next I
End Function

Public Sub CheckRes()
Dim updateR As Boolean, I As Integer, J As Integer, strF, strI, strT, str1
Dim CanUpd As Boolean
    updateR = False
    'ÅÐ¶¨£º
    For I = 0 To NumTopRN
        CanUpd = True
        Call Needcele(ResNeed(I), strF, strI, strT)
        For J = 0 To UBound(strF)
            If Not NumTotalR(strF(J)) Then CanUpd = False
        Next J
        If UBound(strI) >= 0 Then
            For J = 0 To UBound(strI)
                str1 = Split(strI(J), "*")
                If NumTotalI(str1(0)) < str1(1) Then CanUpd = False
            Next J
        End If
        If CanUpd Then
            For J = 0 To UBound(strT)
                If Not updCed(I) Then
                    RO(strT(J)).Status = CFSIsable
                    updateR = True
                End If
            Next J
            updCed(I) = True
        End If
    Next I
    If NumTotalR(26) And checkWP() < 3 Then Call showWP(3)
    If NumTotalR(25) And checkWP() < 2 Then Call showWP(2)
    If NumTotalR(24) And checkWP() < 1 Then Call showWP(1)
    If NumTotalR(23) And checkWP() < 0 Then Call showWP(0)
    '¸üÐÂ
    If updateR Then
        Call ResRef
        Call ResRefresh
    End If
End Sub

