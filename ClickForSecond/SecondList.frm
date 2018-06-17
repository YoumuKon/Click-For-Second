VERSION 5.00
Begin VB.Form SecondList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "枪毙名单"
   ClientHeight    =   3075
   ClientLeft      =   4185
   ClientTop       =   5850
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   2745
   Begin VB.CommandButton Exit 
      Caption         =   "退出"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "如不显示请点击窗体"
      Height          =   255
      Left            =   480
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
        Print Tab(8); Shotlist(0, I); Tab(20); Shotlist(1, I) & "s"
    Next I
End Sub
