VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "钢材计算器"
   ClientHeight    =   5415
   ClientLeft      =   7515
   ClientTop       =   4890
   ClientWidth     =   9165
   HasDC           =   0   'False
   LinkTopic       =   "欢迎使用钢材计算器"
   ScaleHeight     =   5415
   ScaleWidth      =   9165
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "登  录"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   2520
      TabIndex        =   1
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "欢迎使用钢材计算器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   39.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   7335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Hide
Form2.Show
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer) '在form上敲回车触发事件
If KeyAscii = 13 Then '如果按下的是回车键，注意回车Asc码是13
Call Command1_Click '那么执行command1点击事件
End If
End Sub
