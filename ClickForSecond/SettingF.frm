VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SettingF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   2550
   ClientLeft      =   3645
   ClientTop       =   3660
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5565
   Begin VB.TextBox LangAddress 
      CausesValidation=   0   'False
      Height          =   270
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "˫����ѡ���ļ�"
      Top             =   1560
      Width           =   5295
   End
   Begin MSComDlg.CommonDialog Common 
      Left            =   3240
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox ConfigAddress 
      CausesValidation=   0   'False
      Height          =   270
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "˫����ѡ���ļ�"
      Top             =   840
      Width           =   5295
   End
   Begin VB.CommandButton ButtonN 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton ButtonY 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox ClickE 
      Caption         =   "�ڴ���¼���µ���¼�(������)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "��ǰLang�ļ���ַ��"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "��ǰConfig�ļ���ַ��"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "SettingF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ConfigAddress_DblClick()
    SettingF.Common.Filter = "�����ļ�(*.CFSconfig)|*.CFSconfig|ȫ���ļ�(*.*)|*.*"
    Common.ShowOpen
    If Common.FileName = "" Then Exit Sub Else
    ConfigAddress = Common.FileName
End Sub

Private Sub Form_Load()
    If Main.ClickEB Then ClickE.Value = 1
    LangAddress = LangA
    ConfigAddress = ConfigA
End Sub

Private Sub ButtonY_Click()
    If ClickE.Value = 1 Then Main.ClickEB = True Else: Main.ClickEB = False
    ConfigA = ConfigAddress
    LangA = LangAddress
    Call loadC
    Unload Me
End Sub

Private Sub ButtonN_Click()
    Unload Me
End Sub

Private Sub LangAddress_DblClick()
    SettingF.Common.Filter = "�����ļ�(*.CFSlang)|*.CFSlang|ȫ���ļ�(*.*)|*.*"
    Common.ShowOpen
    If Common.FileName = "" Then Exit Sub Else
    LangAddress = Common.FileName
End Sub
