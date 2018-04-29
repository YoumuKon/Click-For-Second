VERSION 5.00
Begin VB.Form ResearchF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "时间研究中心"
   ClientHeight    =   5865
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6225
   StartUpPosition =   3  '窗口缺省
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
   Begin VB.Menu MunR 
      Caption         =   "菜单"
      Begin VB.Menu MnuNow 
         Caption         =   "现在的效率"
      End
   End
End
Attribute VB_Name = "ResearchF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TRes%
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
Dim resN As Integer, RV&
    If Resable.ListIndex = -1 Then
        MsgBox "请选择研究项目!", vbCritical, "未选择研究"
        Else: resN = ResNum(Resable.List(Resable.ListIndex)): RV = ResV(resN)
        If BuyCheck(RV, Ts) Then
            ResTI(1, resN) = ResT(resN)
            NumTotalRN(resN) = False
            ResTI(0, resN) = True
            Resing.AddItem Resable.List(Resable.ListIndex)
            Resable.RemoveItem Resable.ListIndex
            Else: MsgBox "秒数不够!", 16, "秒数不够"
        End If
    End If
End Sub

Private Sub Timer1_Timer()
Dim updateR As Boolean, Resin As Integer
    If Resing.ListCount <> 0 Then
        For TRes = 0 To Resing.ListCount - 1
            Resin = -1
            Do While Resin = -1
                Resin = ResNum(Resing.List(TRes))
            Loop
            If ResTI(1, Resin) = 0 Then
                ResTI(0, Resin) = False
                NumTotalR(Resin) = True
                Resed.AddItem NameR(Resin)
                Resing.RemoveItem TRes
                Else: ResTI(1, Resin) = ResTI(1, Resin) - 1
            End If
        Next TRes
    End If
    If NumTotalS(0) = 10 And NumTotalR(0) And Not NumTotalRN(1) Then NumTotalRN(1) = True: updateR = True
    If NumTotalS(1) = 10 And NumTotalR(1) And Not NumTotalRN(2) Then NumTotalRN(2) = True: updateR = True
    If NumTotalS(2) = 10 And NumTotalR(2) And Not NumTotalRN(3) Then NumTotalRN(3) = True: updateR = True
    If NumTotalR(4) Then ClickP = ClickP + 1: updateR = True
    If updateR Then Call ResRef
End Sub
