VERSION 5.00
Begin VB.Form ShopF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "С�̵�"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5415
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton BuyI 
      Caption         =   "���ֱ���װ 90s"
      Enabled         =   0   'False
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton BuyI 
      Caption         =   "��ͧ 45s"
      Enabled         =   0   'False
      Height          =   615
      Index           =   2
      Left            =   3480
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton BuyI 
      Caption         =   "�ڿ��۾� 10s"
      Enabled         =   0   'False
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton BuyI 
      Caption         =   "�����ı����й��� 20s"
      Enabled         =   0   'False
      Height          =   615
      Index           =   1
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label NumI 
      Caption         =   "Ŀǰ��0��"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label NumI 
      Caption         =   "Ŀǰ��0��"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "�Զ�+1sװ��"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label NumI 
      Caption         =   "Ŀǰ��0��"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label NumI 
      Caption         =   "Ŀǰ��0��"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "ShopF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BuyI_Click(Index As Integer)
    If BuyCheck(ItemV(Index), Ts) Then
        NumTotalS(Index) = NumTotalS(Index) + 1
        Else: MsgBox "��������!", 16, "��������"
    End If
    Call Refe
End Sub

Private Sub Form_Load()
    For I = 0 To NumTopS
        NumI(I) = "Ŀǰ��" & str(NumTotalS(I)) & "��"
    Next I
    For I = 0 To NumTopS
        If NumTotalR(I) Then BuyI(I).Enabled = True
    Next I
End Sub

