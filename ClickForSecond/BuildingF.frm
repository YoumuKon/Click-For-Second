VERSION 5.00
Begin VB.Form BuildingF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "建筑规划室(WIP)"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6270
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox Buildable 
      Height          =   3840
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Builde 
      BackColor       =   &H8000000F&
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "BuildingF.frx":0000
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1320
      Top             =   0
   End
   Begin VB.ListBox Building 
      Height          =   3300
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton BuildStart 
      Caption         =   "建造"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   3840
      Width           =   1935
   End
   Begin VB.ListBox Builded 
      Height          =   5280
      Left            =   4200
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "可用建筑项目"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "建筑中一览"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "现有建筑物"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "BuildingF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call Buildref
End Sub
