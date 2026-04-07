VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "喷漆计算"
   ClientHeight    =   5415
   ClientLeft      =   2520
   ClientTop       =   465
   ClientWidth     =   8805
   LinkTopic       =   "Form8"
   ScaleHeight     =   5415
   ScaleWidth      =   8805
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command3 
      Caption         =   " 圆 管   喷 漆"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   " 方 管   喷 漆"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回首页 "
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "喷漆计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   39.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form8.Hide
Form9.Show
End Sub

Private Sub Command2_Click()
Form8.Hide
Form2.Show
End Sub

Private Sub Command3_Click()
Form8.Hide
Form10.Show
End Sub
