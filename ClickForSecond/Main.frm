VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CFS+"
   ClientHeight    =   8385
   ClientLeft      =   2025
   ClientTop       =   2685
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   9045
   Begin VB.CommandButton NBuild 
      Caption         =   "��������"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton ItemCraft 
      Caption         =   "��Ʒ�ϳ�"
      Height          =   375
      Left            =   7680
      TabIndex        =   16
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton ShotlistY 
      Caption         =   "ǹ������"
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton ItemList 
      Caption         =   "ͳ����Ʒ"
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Setting 
      Caption         =   "����"
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton CopyE 
      Caption         =   "���Ƽ�¼"
      Height          =   375
      Left            =   960
      TabIndex        =   12
      ToolTipText     =   "������¼���Ƶ�������"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Research 
      Caption         =   "�о�"
      Height          =   375
      Left            =   7680
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
      Left            =   6360
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Clear 
      Caption         =   "��ռ�¼"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      ToolTipText     =   "��մ���¼"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox User 
      CausesValidation=   0   'False
      Height          =   270
      Left            =   6360
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
      Width           =   6135
   End
   Begin VB.Label Persec 
      Caption         =   "����1����������: 0s"
      Height          =   255
      Left            =   6360
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
      Width           =   9015
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
      Left            =   8640
      TabIndex        =   5
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6360
      TabIndex        =   4
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "����ʱ�䷨�󴢴������Ϊ��"
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "ͬ־����������ǣ�"
      Height          =   255
      Left            =   4800
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
         Caption         =   "��ȥȫ������(G&)"
      End
      Begin VB.Menu MnuSkill 
         Caption         =   "����"
         Visible         =   0   'False
         Begin VB.Menu MnuSkill0 
            Caption         =   "����轲�(0&)"
            Enabled         =   0   'False
         End
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
    'Ϊ��ֹ�������ִ������
    'On Error Resume Next
    If Dir("MainOption.ini") <> "" Then
        If FileLen("MainOption.ini") <> 0 Then
            Open "MainOption.ini" For Input As #1
            Line Input #1, ConfigA
            Line Input #1, LangA
            Close #1
        End If
    End If
    If ConfigA = "" Then
        SettingF.Common.Filter = "�����ļ�(*.CFSconfig)|*.CFSconfig|ȫ���ļ�(*.*)|*.*"
        SettingF.Common.ShowOpen
        ConfigA = SettingF.Common.FileName
    End If
    If LangA = "" Then
        SettingF.Common.Filter = "�����ļ�(*.CFSlang)|*.CFSlang|ȫ���ļ�(*.*)|*.*"
        SettingF.Common.ShowOpen
        LangA = SettingF.Common.FileName
    End If
    SettingF.ConfigAddress = ConfigA
    SettingF.LangAddress = LangA
    Call mainload
    '��ʼ��
    Ts = 0
    EventS = ""
    Total = 0
    ClickP = 1
    ItemPST = 1
    UserN = User.Text
    For I = 0 To SellI
        ShopF.BuyI(I).Enabled = False
        NumTotalI(I) = 0
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
    'Ĭ������
    ClickEB = False
    Common.Filter = "�����ĵ�(*.savesecond)|*.savesecond|ȫ���ļ�(*.*)|*.*"
    Call showWP(-1)
    Call NumPer
    'Ĭ���о�
    NumTotalRN(0) = True
    NumTotalRN(23) = True
    'Ĭ�ϼ���
    MnuSkill0.Enabled = True
    'Ĭ�Ϻϳ�
    Crafting(1, 0) = True
End Sub

Private Sub Clear_Click()
    EventS = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close
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
            UpdEve StrEnc(StrEnc(EventList(2), "&U", UserN), "&Mem1", tsN)
            Else: MsgBox "��������Ϊ0!", 16, "��������"
        End If
    End If
    If tsN > Shotlist(1, 0) Then
        If MsgBox("������������ʷ�¸�! Ϊ" & tsN & "s" & Chr(13) & "����ǹ��������?", vbQuestion + vbYesNo) = vbYes Then
            Call Shotadd(UserN, tsN)
            SecondList.Show
            UpdEve EventList(3)
        End If
    End If
End Sub

Private Sub ItemCraft_Click()
    CraftingF.Show
End Sub

Private Sub ItemList_Click()
Dim I%
    For I = 0 To NumTopI
        UpdEve StrEnc(StrEnc(EventList(4), "&Mem1", NameI(I)), "&Mem2", NumTotalI(I))
    Next I
End Sub

Private Sub MnuAbout_Click()
    MsgBox "ClickForSecond   By YoumuKon" & Chr(13) & "�汾��: " & CFSVersion
End Sub

Private Sub MnuLoad_Click()
    Call loadF
End Sub

Private Sub MnuSave_Click()
    Call saveF
End Sub

Private Sub MnuSkill0_Click()
    RunSkill 0
End Sub

Private Sub NBuild_Click()
    BuildingF.Show
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
    If MsgBox("��Ҫ�ı����������?" & Chr(13) & "һ���ı佫���ü�¼!", 4 + 48, "���ָı侯��") = vbYes Then
        User.Text = InputBox("����������")
    End If
End Sub

Private Sub WorkPlace_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MnuSkill
        Else
    End If
End Sub

Private Sub WorkPlace_Click()
    Ts = Ts + ClickP
    Total = Ts
    If ClickEB Then UpdEve StrEnc(StrEnc(EventList(0), "&U", UserN), "&Mem1", ClickP)
End Sub

Private Sub WorkPlace_DblClick()
    Ts = Ts + ClickP
    Total = Ts
    If ClickEB Then UpdEve StrEnc(StrEnc(EventList(0), "&U", UserN), "&Mem1", ClickP)
End Sub
