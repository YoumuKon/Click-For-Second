VERSION 5.00
Begin VB.Form SecondList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ǹ������"
   ClientHeight    =   3075
   ClientLeft      =   4185
   ClientTop       =   5850
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3420
   Begin VB.CommandButton Exit 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "�粻��ʾ��������"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
End
Attribute VB_Name = "SecondList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Form_Click()
Dim I%
    SecondList.Cls
    For I = 0 To 9
        Print Tab(9); Shotlist(0, I); Tab(20); Shotlist(1, I) & "s"
    Next I
End Sub
