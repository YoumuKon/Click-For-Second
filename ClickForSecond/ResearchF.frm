VERSION 5.00
Begin VB.Form ResearchF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ʱ���о�����"
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
      Caption         =   "�о�"
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
      Caption         =   "���о���Ŀ"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "�о�����Ŀ"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label label1 
      Caption         =   "�����о���Ŀ"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu MunR 
      Caption         =   "�˵�"
      Begin VB.Menu MnuNow 
         Caption         =   "���ڵ�Ч��"
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
        "���ڻ�ʣ" & ResTI(1, ResNum(Resing.List(Resing.ListIndex))) & "s"
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
        MsgBox "��ѡ���о���Ŀ!", vbCritical, "δѡ���о�"
        Else: resN = ResNum(Resable.List(Resable.ListIndex)): RV = ResV(resN)
        If BuyCheck(RV, Ts) Then
            NumTotalRN(resN) = False
            ResTI(0, resN) = True
            ResTI(1, resN) = ResT(resN)
            Resing.AddItem Resable.List(Resable.ListIndex)
            Resable.RemoveItem Resable.ListIndex
            Else: MsgBox "��������!", 16, "��������"
        End If
    End If
End Sub

Private Sub Timer1_Timer()
Dim Resin As Integer
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
                UpdEve "��" & NameR(Resin) & "��" & "�о��ɹ�!"
                Select Case Resin
                    '������о�ʱֱ��ճ��
                    Case 0: UpdEve "�����Ѿ����Թ���ڿ��۾���!"
                    Case 1: UpdEve "�����Ѿ����Թ������ı����й�����!"
                    Case 2: UpdEve "�����Ѿ����Թ�����ͧ��!"
                    Case 3: UpdEve "�����Ѿ����Թ������ֱ���װ��!"
                    Case 4: UpdEve "�����Ѿ����Թ���Ѽ�����װ��!"
                    Case 5: UpdEve "�ڿ��۾�������Ϊ�����խ���۾�!"
                    Case 6: UpdEve "�����ı����й���������Ϊ����������ѡ��!"
                    Case 7: UpdEve "��ͨ������ͧ������Ϊ���ϲ�����ͧ!"
                    Case 8: UpdEve "��е�ֱ���װ������Ϊ�����ֱ���װ!"
                    Case 9: UpdEve "Ѽ�����װ������Ϊ��ЧѼ���!"
                    Case 10: UpdEve "�����������ѽ������!"
                    Case 11: UpdEve "������Ա�������ѽ������!"
                    Case 12: UpdEve "�������㳡�ѽ������!"
                End Select
                Call ResShop
                ElseIf ResTI(1, Resin) > 0 Then ResTI(1, Resin) = ResTI(1, Resin) - 1
            End If
        Next TRes
    End If
    Call CheckRes
End Sub
