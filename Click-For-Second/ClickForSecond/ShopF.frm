VERSION 5.00
Begin VB.Form ShopF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "С�̵�"
   ClientHeight    =   4050
   ClientLeft      =   6345
   ClientTop       =   1830
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   10230
   Begin VB.CommandButton BuyI 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   615
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton BuyI 
      Caption         =   "��轲�"
      Enabled         =   0   'False
      Height          =   615
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton BuyI 
      Caption         =   "�ڽ���Ƭ"
      Enabled         =   0   'False
      Height          =   615
      Index           =   5
      Left            =   8520
      TabIndex        =   11
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton BuyI 
      Caption         =   "��ͧ"
      Enabled         =   0   'False
      Height          =   615
      Index           =   4
      Left            =   6840
      TabIndex        =   9
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton BuyI 
      Caption         =   "Ѽ�����װ"
      Enabled         =   0   'False
      Height          =   615
      Index           =   3
      Left            =   5160
      TabIndex        =   7
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton BuyI 
      Caption         =   "��е�ֱ���װ"
      Enabled         =   0   'False
      Height          =   615
      Index           =   2
      Left            =   3480
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton BuyI 
      Caption         =   "�ڿ��۾�"
      Enabled         =   0   'False
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton BuyI 
      Caption         =   "�����ı����й���"
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
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "һ����������Ʒ"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label NumI 
      Caption         =   "Ŀǰ��0��"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label NumI 
      Caption         =   "Ŀǰ��0��"
      Height          =   255
      Index           =   5
      Left            =   8520
      TabIndex        =   12
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label NumI 
      Caption         =   "Ŀǰ��0��"
      Height          =   255
      Index           =   4
      Left            =   6840
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label NumI 
      Caption         =   "Ŀǰ��0��"
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   8
      Top             =   1080
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
      Caption         =   "�Զ�������Ʒ"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1575
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
    If BuyCheck(ItemV(Index) * (1 + NumTotalI(Index) * (0.25 * ItemPS(Index))), Ts) Then
        NumTotalI(Index) = NumTotalI(Index) + 1
        Else: MsgBox "��������!", 16, "��������"
    End If
    Call Refe
    Call ResRefresh
End Sub

Private Sub Form_Load()
    Call ResRefresh
End Sub

