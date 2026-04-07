VERSION 5.00
Begin VB.Form Form14 
   Caption         =   "涨磅计算"
   ClientHeight    =   5715
   ClientLeft      =   7650
   ClientTop       =   4950
   ClientWidth     =   9165
   LinkTopic       =   "Form14"
   ScaleHeight     =   5715
   ScaleWidth      =   9165
   Begin VB.CommandButton Command3 
      Caption         =   "清  空"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      TabIndex        =   16
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计  算"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      TabIndex        =   15
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回首页"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6840
      TabIndex        =   13
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6840
      TabIndex        =   11
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   9
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "吨"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   8640
      TabIndex        =   22
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "kg"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   8520
      TabIndex        =   21
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "吨"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   3840
      TabIndex        =   20
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "μm"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   3720
      TabIndex        =   19
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   3720
      TabIndex        =   18
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   3720
      TabIndex        =   17
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "总计涨磅重量"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4800
      TabIndex        =   12
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "每吨黑材涨磅重量"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4560
      TabIndex        =   10
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "黑料总重"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   600
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   $"Form14.frx":0000
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   480
      TabIndex        =   7
      Top             =   4680
      Width           =   7935
   End
   Begin VB.Label Label2 
      Caption         =   "镀锌层厚度"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "用料厚度"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "展开宽度"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "涨磅计算"
      BeginProperty Font 
         Name            =   "华文新魏"
         Size            =   39.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text5.Text = Format(((Val(Text1.Text) * Val(Text3.Text) * 7.14 * 2) / ((Val(Text1.Text) * Val(Text3.Text) * 7.14 * 2) + (Val(Text1.Text) * Val(Text2.Text) * 0.00785)) * 10), "0.00")
Text6.Text = Format(Val(Text4.Text) * (Val(Text5.Text) / 10), "0.00")
End Sub

Private Sub Command2_Click()
Form14.Hide
Form2.Show
End Sub

Private Sub Text1_Change(Index As Integer)
 If KeyAscii = 46 And Not CBool(InStr(txbNumber, ".")) Then Exit Sub
 If KeyAscii = 8 Then Exit Sub
 If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text2_Change(Index As Integer)
 If KeyAscii = 46 And Not CBool(InStr(txbNumber, ".")) Then Exit Sub
 If KeyAscii = 8 Then Exit Sub
 If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text3_Change(Index As Integer)
 If KeyAscii = 46 And Not CBool(InStr(txbNumber, ".")) Then Exit Sub
 If KeyAscii = 8 Then Exit Sub
 If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text4_Change(Index As Integer)
 If KeyAscii = 46 And Not CBool(InStr(txbNumber, ".")) Then Exit Sub
 If KeyAscii = 8 Then Exit Sub
 If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

