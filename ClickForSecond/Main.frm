VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CFS+"
   ClientHeight    =   6255
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6630
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Research 
      Caption         =   "�о�"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   3720
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog Common 
      Left            =   1800
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1320
      Top             =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�̵�"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Clear 
      Caption         =   "��ռ�¼"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "��մ���¼"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox User 
      Height          =   270
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Youmu"
      ToolTipText     =   "�޸����ֻᵼ�¼�¼����"
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox EventS 
      Height          =   3855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label Persec 
      Caption         =   "����1s��������: 0s"
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   3360
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
      Height          =   1815
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label Label4 
      Caption         =   "����¼��"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "s"
      Height          =   255
      Left            =   6240
      TabIndex        =   5
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "��ĿǰΪֹ���������ˣ�"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "ͬ־����������ǣ�"
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Menu Menu 
      Caption         =   "�˵�"
      Begin VB.Menu MnuSave 
         Caption         =   "����浵(S&)"
      End
      Begin VB.Menu MnuLoad 
         Caption         =   "����浵(L&)"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    Call Mainload
    Common.Filter = "�����ĵ�(*.savesecond)|*.savesecond|ȫ���ļ�(*.*)|*.*"
    Ts = 0
    EventS = ""
    chg = 0
    Total = 0
    For I = 0 To NumTopS
        NumTotalS(I) = 0
    Next I
    For I = 0 To NumTopR
        NumTotalR(I) = False
        NumTotalRN(I) = False
        Select Case I
        End Select
    Next I
    ResearchF.Resable.AddItem NameR(0)
    ResearchF.Resable.AddItem NameR(4)
    NumTotalRN(0) = True
    NumTotalRN(4) = True
    ClickP = 1
    Load ShopF
    Load ResearchF
End Sub

Private Sub Clear_Click()
    EventS = ""
End Sub

Private Sub Command1_Click()
    ShopF.Show
End Sub

Private Sub MnuLoad_Click()
    Call loadf
End Sub

Private Sub MnuSave_Click()
    Call saveF
End Sub

Private Sub Research_Click()
    ResearchF.Show
End Sub

Private Sub Timer1_Timer()
    sper = 0
    Call NumPer
    Ts = Ts + sper
    Total = str(Ts)
End Sub

Private Sub User_Change()
    Ts = 0
    Total = str(Ts)
    For I = 0 To NumTopS
        NumTotalS(I) = 0
    Next I
    Call Refe
End Sub

Private Sub User_Click()
    chg = MsgBox("��Ҫ�ı����������?" & Chr(13) & "һ���ı佫���ü�¼!", 4 + 48, "���ָı侯��")
    If chg = vbYes Then
        User.Text = InputBox("����������")
    End If
End Sub

Private Sub WorkPlace_Click()
    Ts = Ts + ClickP
    Total = str(Ts)
    EventS = User & "Ϊ��ҵ������" & ClickP & "s" & vbCrLf & EventS
End Sub

