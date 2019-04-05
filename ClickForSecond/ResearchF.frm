VERSION 5.00
Begin VB.Form ResearchF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "时间研究室"
   ClientHeight    =   5865
   ClientLeft      =   8325
   ClientTop       =   1470
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
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
      ScrollBars      =   2  'Vertical
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
   Begin VB.Label Label1 
      Caption         =   "可用研究项目"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1935
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

Private Sub Resable_Click()
    Resde = showde(Resable.List(Resable.ListIndex))
    Resing.ListIndex = -1
    Resed.ListIndex = -1
End Sub

Private Sub Resing_Click()
    If Resing.ListIndex = -1 Then
        Resde = showde("")
        Else: Resde = showde(Resing.List(Resing.ListIndex)) & vbCrLf & _
        "现在还剩" & RO(ResNum(Resing.List(Resing.ListIndex))).TimeNow & "s"
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
        Else: resN = ResNum(Resable.List(Resable.ListIndex)): RV = RO(resN).Valve
        If BuyCheck(RV, Ts) Then
            If NeedItemCheck(RO(resN).NeedItem, RO(resN).NeedItemNumber) Then
                RO(resN).Status = CFSisdoing
                RO(resN).TimeNow = RO(resN).Time
                Resing.AddItem Resable.List(Resable.ListIndex)
                Resable.RemoveItem Resable.ListIndex
                Else: MsgBox "所需材料不足!", 16, "材料不足"
            End If
            Else: MsgBox "秒数不够!", 16, "秒数不够"
        End If
    End If
End Sub

Private Sub Timer1_Timer()
Dim Resin As Integer, TRes%, i As Integer
    If Resing.ListCount <> 0 Then
        For TRes = Resing.ListCount - 1 To 0 Step -1
            Resin = -1
            Do While Resin = -1
                Resin = ResNum(Resing.List(TRes))
            Loop
            If RO(Resin).TimeNow = 0 Then
                RO(Resin).Status = CFSisdone
                Resed.AddItem RO(Resin).Name
                Resing.RemoveItem TRes
                UpdEve StrEnc(EventList(1), StrMem1, RO(Resin).Name)
                For i = 0 To NumTopR
                    If i = Resin Then
                        UpdEve StrEnc(RO(Resin).Event, StrMem1, RO(Resin).Name)
                    End If
                Next i
                Call ResRefresh
                ElseIf RO(Resin).TimeNow > 0 Then RO(Resin).TimeNow = RO(Resin).TimeNow - 1
            End If
        Next TRes
    End If
    Call CheckRes
End Sub

Private Function showde(ind As String) As String
Dim NumR%, strI, strN, i%
    NumR = ResNum(ind)
    If NumR < 0 Then
        showde = "点击研究项目显示描述" & vbCrLf & "点击'研究'按钮以开始研究": Exit Function
        Else: showde = StrEnc(RO(NumR).Description, StrCrLf, vbCrLf) & vbCrLf & "消耗" & RO(NumR).Valve & "s" & ",研究时长" & RO(NumR).Time & "s"
    End If
    If RO(NumR).NeedItem <> "" Then
        strI = Split(RO(NumR).NeedItem, "|")
        strN = Split(RO(NumR).NeedItemNumber, "|")
        showde = showde & vbCrLf & "所需物品:"
        For i = 0 To UBound(strN) - 1
            showde = showde & vbCrLf & NameI(strI(i)) & ":" & strN(i)
        Next i
    End If
    showde = ind & vbCrLf & showde
End Function


