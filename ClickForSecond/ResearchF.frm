VERSION 5.00
Begin VB.Form ResearchF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "时间研究中心"
   ClientHeight    =   5865
   ClientLeft      =   10020
   ClientTop       =   1785
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6225
   Begin VB.ListBox Resed 
      Height          =   5280
      Left            =   4200
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton ResStart 
      Caption         =   "研究"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   3840
      Width           =   1935
   End
   Begin VB.ListBox Resing 
      Height          =   3300
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1320
      Top             =   0
   End
   Begin VB.TextBox Resde 
      BackColor       =   &H8000000F&
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "ResearchF.frx":0000
      Top             =   4440
      Width           =   3975
   End
   Begin VB.ListBox Resable 
      Height          =   3840
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "已研究项目"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "研究中项目"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label label1 
      Caption         =   "可用研究项目"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu MnuUser 
      Caption         =   "用户菜单"
      Begin VB.Menu MnuUskill 
         Caption         =   "技能"
         Begin VB.Menu USkill0 
            Caption         =   "喝枸杞茶(1&)"
         End
      End
   End
End
Attribute VB_Name = "ResearchF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call ResRef
End Sub

Private Sub Resing_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu MnuUser
End Sub

Private Sub Resable_Click()
    Resde = showde(Resable.List(Resable.ListIndex))
    Resing.ListIndex = -1
    Resed.ListIndex = -1
End Sub

Private Sub Resing_Click()
    If Resing.ListIndex = -1 Then
        Resde = showde("")
        Else: Resde = showde(Resing.List(Resing.ListIndex)) & vbCrLf & _
        "现在还剩" & ResTI(1, ResNum(Resing.List(Resing.ListIndex))) & "s"
    End If
    Resable.ListIndex = -1
    Resed.ListIndex = -1
End Sub

Private Sub Resed_Click()
    Resde = showde(Resed.List(Resed.ListIndex))
    Resable.ListIndex = -1
    Resing.ListIndex = -1
End Sub


Private Sub ResStart_Click()
Dim resN As Integer, RV As Double
    If Resable.ListIndex = -1 Then
        MsgBox "请选择研究项目!", vbCritical, "未选择研究"
        Else: resN = ResNum(Resable.List(Resable.ListIndex)): RV = ResV(resN)
        If BuyCheck(RV, Ts) Then
            NumTotalRN(resN) = False
            ResTI(0, resN) = True
            ResTI(1, resN) = ResT(resN)
            Resing.AddItem Resable.List(Resable.ListIndex)
            Resable.RemoveItem Resable.ListIndex
            Else: MsgBox "秒数不够!", 16, "秒数不够"
        End If
    End If
End Sub

Private Sub Timer1_Timer()
Dim Resin As Integer, TRes%
    If Resing.ListCount <> 0 Then
        For TRes = Resing.ListCount - 1 To 0 Step -1
            Resin = -1
            Do While Resin = -1
                Resin = ResNum(Resing.List(TRes))
            Loop
            If ResTI(1, Resin) = 0 Then
                ResTI(0, Resin) = False
                NumTotalR(Resin) = True
                Resed.AddItem NameR(Resin)
                Resing.RemoveItem TRes
                UpdEve "“" & NameR(Resin) & "”" & "研究成功!"
                Select Case Resin
                    '添加新研究时直接粘贴
                    Case 0: UpdEve "现在已经可以购买黑框眼镜了!"
                    Case 1: UpdEve "现在已经可以购买《他改变了中国》了!"
                    Case 2: UpdEve "现在已经可以购买机械手表套装了!"
                    Case 3: UpdEve "现在已经可以购买普通鸭嘴笔了!"
                    Case 4: UpdEve "现在已经可以购买赛艇了!"
                    Case 5: UpdEve "现在已经可以购买《Aloha 'Oe》黑胶唱片了!"
                    Case 6: UpdEve "黑框眼镜已升级为意大利窄边眼镜!"
                    Case 7: UpdEve "《他改变了中国》已升级为《江泽民文选》!"
                    Case 8: UpdEve "机械手表套装已升级为电子手表套装!"
                    Case 9: UpdEve "普通鸭嘴笔已升级为高效鸭嘴笔!"
                    Case 10: UpdEve "普通材料赛艇已升级为复合材料赛艇!"
                    Case 11: UpdEve "黑胶唱片已升级为VCD!"
                    Case 12: UpdEve "工作区房屋已建造完毕!"
                    Case 13: UpdEve "工作区员工宿舍已建造完毕!"
                    Case 14: UpdEve "工作区广场已建造完毕!"
                    Case 15: UpdEve "现在已经可以购买枸杞茶了!"
                End Select
                Call ResShop
                ElseIf ResTI(1, Resin) > 0 Then ResTI(1, Resin) = ResTI(1, Resin) - 1
            End If
        Next TRes
    End If
    Call CheckRes
End Sub

Private Sub USkill0_Click()
Dim ReS0%, TS0%
    If MsgBox("技能需要消耗1枸杞茶" & Chr(13) & "现在有" & NumTotalS(6) & "个枸杞茶" & Chr(13) & "确定要使用技能吗?", _
    vbYesNo, "喝枸杞茶") = vbYes Then
        If BuyCheck(ItemV(6), Ts) Then
            NumTotalS(6) = NumTotalS(6) - 1
            For TS0 = Resing.ListCount - 1 To 0 Step -1
                ReS0 = -1
                Do While ReS0 = -1
                    ReS0 = ResNum(Resing.List(TS0))
                Loop
                If ResTI(1, ReS0) > 0 Then ResTI(1, ReS0) = ResTI(1, ReS0) - 60
                If ResTI(1, ReS0) < 0 Then ResTI(1, ReS0) = 0
            Next TS0
            MsgBox "技能使用成功!", 0, "使用成功"
            Else: MsgBox "物品数不够!", 16, "使用失败"
        End If
    End If
End Sub
