VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CFS+"
   ClientHeight    =   8385
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   8055
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton ShotlistY 
      Caption         =   "ǹ������"
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton ItemList 
      Caption         =   "ͳ����Ʒ"
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton ModSet 
      Caption         =   "Mod..."
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Setting 
      Caption         =   "����"
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton CopyE 
      Caption         =   "���Ƽ�¼"
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      ToolTipText     =   "������¼���Ƶ�������"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Research 
      Caption         =   "�о�"
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   4440
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog Common 
      Left            =   2040
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1440
      Top             =   2280
   End
   Begin VB.CommandButton shop 
      Caption         =   "�̵�"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Clear 
      Caption         =   "��ռ�¼"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "��մ���¼"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox User 
      Height          =   270
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Youmu"
      ToolTipText     =   "�޸����ֻᵼ�¼�¼����"
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox EventS 
      Height          =   4815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3480
      Width           =   5175
   End
   Begin VB.Label Persec 
      Caption         =   "����1����������: 0s"
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label WorkPlace 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���ǹ�����"
      BeginProperty Font 
         Name            =   "������"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label Label4 
      Caption         =   "����¼��"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "s"
      Height          =   255
      Left            =   7680
      TabIndex        =   5
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "����ʱ�䷨�󴢴������Ϊ��"
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "ͬ־����������ǣ�"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Menu Menu 
      Caption         =   "�˵�"
      Begin VB.Menu MnuSaveData 
         Caption         =   "�浵"
         Begin VB.Menu MnuSave 
            Caption         =   "����浵(S&)"
         End
         Begin VB.Menu MnuLoad 
            Caption         =   "����浵(L&)"
         End
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "����(A&)"
      End
      Begin VB.Menu GiveAwaySecond 
         Caption         =   "�Ͻ�ȫ������(G&)"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ClickEB As Boolean
Private Sub CopyE_Click()
    Clipboard.Clear
    Clipboard.SetText "��ǰ����:" & Date & vbCrLf & EventS
    MsgBox "�Ѹ��Ƴɹ�!", 0, "���Ƴɹ�"
End Sub

Private Sub Form_Load()
Dim I%
    On Error Resume Next
    Call Mainconst
    '��ʼ��
    Ts = 0
    EventS = ""
    chg = 0
    Total = 0
    For I = 0 To NumTopI
        ShopF.BuyI(I).Enabled = False
        NumTotalS(I) = 0
        ItemPS(I) = 1
    Next I
    For I = 0 To NumTopR
        NumTotalR(I) = False
        ResTI(0, I) = False
        NumTotalRN(I) = False
    Next I
    For I = 0 To 9
        Shotlist(0, I) = "���ϰ�"
        Shotlist(1, I) = 0
    Next I
    ClickP = 1
    'Ĭ������
    ResearchF.Resable.AddItem NameR(0)
    ResearchF.Resable.AddItem NameR(12)
    NumTotalRN(0) = True
    NumTotalRN(12) = True
    ClickEB = True
    Common.Filter = "�����ĵ�(*.savesecond)|*.savesecond|ȫ���ļ�(*.*)|*.*"
    Call showWP(-1)
    Call NumPer
    Call ResRef
End Sub

Private Sub Clear_Click()
    EventS = ""
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub GiveAwaySecond_Click()
Dim tsN As Double
    tsN = 0
    If MsgBox("���Ҫ��ȫ������������������?" & Chr(13) & "�������������Ϊ0!", vbExclamation + vbYesNo, "����") = vbYes Then
        If Ts > 0 Then
            tsN = Ts
            Ts = 0
            Total = Ts
            MsgBox "�����ɹ�!"
            UpdEve User & "�����˳���" & tsN & "s"
            Else: MsgBox "��������Ϊ0!", 16, "��������"
        End If
    End If
    If tsN > Shotlist(1, 0) Then
        If MsgBox("������������ʷ�¸�! Ϊ" & tsN & "s" & Chr(13) & "����ǹ��������?", vbQuestion + vbYesNo) = vbYes Then
            Call Shotadd(User, tsN)
            SecondList.Show
            UpdEve User & "������ǹ����������!"
        End If
    End If
End Sub

Private Sub ItemList_Click()
Dim I%
    For I = 0 To NumTopI
        UpdEve NameI(I) & ":" & NumTotalS(I)
    Next I
End Sub

Private Sub MnuAbout_Click()
    MsgBox "�һ���ϷClickForSecond   By YoumuKon" & Chr(13) & "�汾��: " & CFSVersion
End Sub

Private Sub MnuLoad_Click()
    Call loadf
End Sub

Private Sub MnuSave_Click()
    Call saveF
End Sub

Private Sub ModSet_Click()
    ModSetting.Show
End Sub

Private Sub Research_Click()
    ResearchF.Show
End Sub

Private Sub Setting_Click()
    SettingF.Show
End Sub

Private Sub shop_Click()
    ShopF.Show
End Sub

Private Sub ShotlistY_Click()
    SecondList.Show
End Sub

Private Sub Timer1_Timer()
    sper = 0
    Call NumPer
    Ts = Ts + sper
    Total = str(Ts)
End Sub

Private Sub User_Change()
    Call Form_Load
End Sub

Private Sub User_Click()
    chg = MsgBox("��Ҫ�ı����������?" & Chr(13) & "һ���ı佫���ü�¼!", 4 + 48, "���ָı侯��")
    If chg = vbYes Then
        User.Text = InputBox("����������")
    End If
End Sub

Private Sub WorkPlace_Click()
    Ts = Ts + ClickP
    Total = Ts
    If ClickEB Then UpdEve User & "������" & ClickP & "s"
End Sub

