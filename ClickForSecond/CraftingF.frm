VERSION 5.00
Begin VB.Form CraftingF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "制造工厂"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5250
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Craftstart 
      Caption         =   "制作"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.ListBox CraftList 
      Height          =   3120
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Craftde 
      BackColor       =   &H8000000F&
      Height          =   3735
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "CraftingF.frx":0000
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "CraftingF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CraftList_Click()
    Craftde = ShowCraftde(CraftList.List(CraftList.ListIndex))
End Sub

Private Sub Craftstart_Click()
Dim CN As Integer, str1, str2, strI, strN
    If CraftList.ListIndex = -1 Then
        MsgBox "请选择目标物品!", vbCritical, "未选择物品"
        Else
        CN = CraftNum(CraftList.List(CraftList.ListIndex))
        str1 = Split(Crafting(0, CN), "+")
        For I = 0 To UBound(str1)
            str2 = Split(str1(I), "*")
            strI = strI & str2(0) & "|"
            strN = strN & str2(1) & "|"
        Next I
        If NeedItemCheck(CStr(strI), CStr(strN)) Then
            If ProCele(CraftP) Then
                NumTotalI(1 + SellI + CN) = NumTotalI(1 + SellI + CN) + 1
                Call UpdEve(StrEnc(EventList(7), StrMem1, NameI(1 + SellI + CN)))
                Else: UpdEve (StrEnc(EventList(8), StrMem1, NameI(1 + SellI + CN)))
            End If
            Else: MsgBox "所需材料不足!", 16, "材料不足"
        End If
    End If
    Call CraftList_Click
End Sub

Private Sub Form_Load()
    Call RefCraft
End Sub

Public Function ShowCraftde(ind As String) As String
Dim NumC%, str1, str2, I%
    NumC = CraftNum(ind)
    If NumC < 0 Then ShowCraftde = "点击物品显示所需材料" & vbCrLf & "点击 '制作'按钮以开始制作": Exit Function
    ShowCraftde = "现在共有" & NumTotalI(1 + SellI + NumC) & "个"
    str1 = Split(Crafting(0, NumC), "+")
    ShowCraftde = ShowCraftde & vbCrLf & "所需物品:"
    For I = 0 To UBound(str1)
        str2 = Split(str1(I), "*")
        ShowCraftde = ShowCraftde & vbCrLf & NameI(str2(0)) & ":" & str2(1)
    Next I
    ShowCraftde = ind & vbCrLf & ShowCraftde
End Function

Public Sub RefCraft()
Dim I%
    CraftingF.CraftList.Clear
    For I = 0 To NumTopC
        If Crafting(1, I) <> "" Then
            If RO(Crafting(1, I)).Status = CFSisdone Then CraftingF.CraftList.AddItem NameII(0, I + SellI + 1)
        End If
    Next I
End Sub

