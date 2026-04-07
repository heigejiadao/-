VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "钢材计算器"
   ClientHeight    =   5415
   ClientLeft      =   7650
   ClientTop       =   4950
   ClientWidth     =   9165
   LinkTopic       =   "Form2"
   ScaleHeight     =   5415
   ScaleWidth      =   9165
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command9 
      Caption         =   "椭圆管"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   0
      Left            =   7440
      TabIndex        =   9
      Top             =   2640
      Width           =   1000
   End
   Begin VB.CommandButton Command8 
      Caption         =   "CZ型钢"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   1
      Left            =   5880
      TabIndex        =   8
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton Command7 
      Caption         =   "运 费  合 算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   0
      Left            =   4200
      TabIndex        =   7
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton Command6 
      Caption         =   "喷 漆"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   2520
      TabIndex        =   6
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton Command5 
      Caption         =   "高频焊  H型钢"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   840
      TabIndex        =   5
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF0000&
      Caption         =   "带 钢  锌 层"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   5880
      TabIndex        =   4
      Top             =   2640
      Width           =   1000
   End
   Begin VB.CommandButton Command3 
      Caption         =   " 废 料  计 算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   1
      Left            =   4200
      TabIndex        =   3
      Top             =   2640
      Width           =   1000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "圆 管"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   2640
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "方 管"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   840
      TabIndex        =   1
      Top             =   2640
      Width           =   1000
   End
   Begin VB.Label Label1 
      Caption         =   "钢材计算器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   39.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2640
      TabIndex        =   0
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
Form3.Show
End Sub
Private Sub Command2_Click(Index As Integer)
Form2.Hide
Form4.Show
End Sub
Private Sub Command3_Click(Index As Integer)
Form2.Hide
Form5.Show
End Sub
Private Sub Command4_Click()
Form2.Hide
Form6.Show
End Sub
Private Sub Command5_Click()
Form2.Hide
Form7.Show
End Sub
Private Sub Command6_Click()
Form2.Hide
Form8.Show
End Sub
Private Sub Command7_Click(Index As Integer)
Form2.Hide
Form13.Show
End Sub
Private Sub Command8_Click(Index As Integer)
Form2.Hide
Form11.Show
End Sub

Private Sub Command9_Click(Index As Integer)
Form2.Hide
Form14.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
