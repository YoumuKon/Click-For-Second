VERSION 5.00
Begin VB.Form SettingF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   3270
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ButtonN 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton ButtonY 
      Caption         =   "确定"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CheckBox ClickE 
      Caption         =   "在大事录记下点击事件(调试用)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "SettingF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If Main.ClickEB Then ClickE.Value = 1
End Sub

Private Sub ButtonY_Click()
    If ClickE.Value = 1 Then Main.ClickEB = True Else: Main.ClickEB = False
    Unload Me
End Sub

Private Sub ButtonN_Click()
    Unload Me
End Sub
