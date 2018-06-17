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

Private Sub Craftde_Change()

End Sub

Private Sub CraftList_Click()
    Craftde = ShowCraftde(CraftList.List(CraftList.ListIndex))
End Sub

Private Sub Form_Load()
    Call ResCraft
End Sub

Private Function ShowCraftde(Ind As String)
Dim NumR%, strI, strN, I%
    NumC = -1
    For I = 0 To NumTopI
        If NameII(0, I) = Ind Then NumC = I: Exit Function
    Next I
    If NumC < 0 Then ShowCraftde = "点击物品显示所需材料" & vbCrLf & "点击 '制作'按钮以开始制作": Exit Function
    strI = Split(Crafting(0, NumC), "+")
    strN = Split(ResVI(1, NumC), "*")
    If UBound(strN) > 0 Then
        ShowCraftde = ShowCraftde & vbCrLf & "以及:"
        For I = 0 To UBound(strN) - 1
            ShowCraftde = ShowCraftde & " " & NameI(strI(I)) & ":" & strN(I)
        Next I
    End If
    ShowCraftde = Ind & vbCrLf & ShowCraftde
End Function
