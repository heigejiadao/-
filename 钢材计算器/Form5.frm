VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "管材废料计算器"
   ClientHeight    =   5415
   ClientLeft      =   2520
   ClientTop       =   465
   ClientWidth     =   8805
   LinkTopic       =   "Form5"
   ScaleHeight     =   5415
   ScaleWidth      =   8805
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command3 
      Caption         =   "清  空"
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
      Left            =   6840
      TabIndex        =   7
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计  算"
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
      Left            =   4800
      TabIndex        =   6
      Top             =   4080
      Width           =   1575
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
      Height          =   420
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3120
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
      Height          =   420
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2280
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
      Height          =   420
      Left            =   1800
      TabIndex        =   4
      Top             =   4080
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
      Height          =   420
      Left            =   1800
      TabIndex        =   3
      Top             =   3360
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
      Height          =   420
      Left            =   1800
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
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
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回首页 "
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "成本单价"
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
      Left            =   360
      TabIndex        =   21
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "元/吨"
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
      Left            =   7905
      TabIndex        =   20
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "成品单价"
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
      Left            =   4800
      TabIndex        =   19
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "吨"
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
      Left            =   7920
      TabIndex        =   17
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   "吨"
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
      Left            =   3500
      TabIndex        =   16
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "元/吨"
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
      Left            =   3500
      TabIndex        =   15
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "吨"
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
      Left            =   3500
      TabIndex        =   14
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "元/吨"
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
      Left            =   3500
      TabIndex        =   13
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "成品吨位"
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
      Left            =   4800
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "废料吨位"
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
      Left            =   360
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "废料单价"
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
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "成本吨位"
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
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "管材废料计算器"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   5775
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text5.Text = Format(Val(Text2.Text) - Val(Text4.Text), "0.00")
Text6.Text = Format((Val(Text1.Text) * Val(Text2.Text) - Val(Text3.Text) * Val(Text4.Text)) / Val(Text5.Text), "0.00")
End Sub
Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Form5.Hide
Form2.Show
End Sub
Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer) '在form上敲回车触发事件
If KeyAscii = 13 Then '如果按下的是回车键，注意回车Asc码是13
Call Command1_Click '那么执行command1点击事件
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 46 And Not CBool(InStr(txbNumber, ".")) Then Exit Sub
 If KeyAscii = 8 Then Exit Sub
 If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 46 And Not CBool(InStr(txbNumber, ".")) Then Exit Sub
 If KeyAscii = 8 Then Exit Sub
 If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 46 And Not CBool(InStr(txbNumber, ".")) Then Exit Sub
 If KeyAscii = 8 Then Exit Sub
 If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
 If KeyAscii = 46 And Not CBool(InStr(txbNumber, ".")) Then Exit Sub
 If KeyAscii = 8 Then Exit Sub
 If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

