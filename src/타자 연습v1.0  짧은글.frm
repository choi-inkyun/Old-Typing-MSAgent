VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000A&
   Caption         =   "짧은글"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   6075
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   12
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   4365
   ScaleWidth      =   6075
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   1920
      Top             =   2400
   End
   Begin VB.CommandButton Command3 
      Caption         =   "메뉴(&E)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "중단하기(&S)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5160
      Top             =   3840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "시작하기(&S)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label Label14 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   39
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label13 
      Caption         =   "0 "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   38
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   37
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "정확도 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   36
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   35
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "타수 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   34
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label8 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   33
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "확인 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   32
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "확인 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   30
      Top             =   3480
      Width           =   4935
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   22
      Left            =   5400
      TabIndex        =   29
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   21
      Left            =   5160
      TabIndex        =   28
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   20
      Left            =   4920
      TabIndex        =   27
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   19
      Left            =   4680
      TabIndex        =   26
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   18
      Left            =   4440
      TabIndex        =   25
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   17
      Left            =   4200
      TabIndex        =   24
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   16
      Left            =   3960
      TabIndex        =   23
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   15
      Left            =   3720
      TabIndex        =   22
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   14
      Left            =   3480
      TabIndex        =   21
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   13
      Left            =   3240
      TabIndex        =   19
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   12
      Left            =   3000
      TabIndex        =   18
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   2760
      TabIndex        =   17
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   2520
      TabIndex        =   16
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   2280
      TabIndex        =   15
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   2040
      TabIndex        =   14
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   1800
      TabIndex        =   13
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   1560
      TabIndex        =   12
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   1320
      TabIndex        =   11
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   1080
      TabIndex        =   10
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   840
      TabIndex        =   9
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label5 
      Caption         =   "맞춘갯수:"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   180
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 문장(50), v, k, c, d

Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
Label6.Visible = True
Timer1.Enabled = True
Text1.SetFocus
For i = 0 To k
Label1(i).Caption = "x"
Label1(i).FontSize = Text1.FontSize
Label1(i).Visible = True
Timer2.Enabled = False
Next
End Sub

Private Sub Command2_Click()
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Label6.Visible = False
Timer1.Enabled = False
For i = 0 To k
Label1(i).Visible = False
Next
Timer2.Enabled = False
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
문장(1) = "말 한마디로 천냥빛을 갚는다."
문장(2) = "낮말은새가듣고 밤말은쥐가듣는다."
문장(3) = "가는말이 고와야 오는 말도 곱다."
문장(4) = "소 귀에 경 읽기"
문장(5) = "낫보고 기역자도 모른다."
문장(6) = "가까운 이웃이 먼 친척보다 낫다."
문장(7) = "가는 날이 장날이다."
문장(8) = "간이 콩알만하다."
문장(9) = "가재는 게 편"
문장(10) = "가지많은 나무에 바람잘 날이 없다."
문장(11) = "간에 기별도 가지 않는다."
문장(12) = "개밥에 도토리"
문장(13) = "간에 기별도 가지 않는다."
문장(14) = "개천에서 용 난다."
문장(15) = "검은 머리 파 뿌리 되도록"
문장(16) = "공든 탑이 무너지랴"
문장(17) = "까마귀 날자 배 떨어진다."
문장(18) = "꿩 대신 닭"
문장(19) = "꿩 먹고 알 먹는다."
문장(20) = "굳은 땅에 물이 고인다."
문장(21) = "누워서 떡 먹기"
문장(22) = "냉수 먹고 이쑤시기"
문장(23) = "돌다리도 두드려 보고 건너라"
문장(24) = "달걀로 바위 치기"
문장(25) = "들으면 병이요 안들으면 약이다."
문장(26) = "땅 짚고 헤엄치기"
문장(27) = "되로 주고 말로 받는다."
문장(28) = "바늘 가는데 실이 간다."
문장(29) = "소 잃고 외양간 고친다."
문장(30) = "수박 겉 햝기"
문장(31) = "아니 되면 조상 탓"
문장(32) = "발 없는 말이 천리 간다."
문장(33) = "물에 빠진 새앙쥐"
문장(34) = "만리길도 한 걸음으로 시작된다."
문장(35) = "마른 하늘에 날벼락"
문장(36) = "아는 것이 병"
문장(37) = "원수는외나무 다리에서 만난다."
문장(38) = "울며 겨자먹기"
문장(39) = "엎어지면 꼬 닿을 때"
문장(40) = "지렁이도 밟으면 꿈틀한다."
문장(41) = "탕약에 감초 빠질까"
문장(42) = "피는 물보다 진하다."
문장(43) = "제도끼에 제 발등을 찍는다."
문장(44) = "제 꾀에 넘어간다."
문장(45) = "정신 일도 하사 불성"
문장(46) = "점잖은 개가 부뚜막에 오른다."
문장(47) = "입술에 침이나 바르지"
문장(48) = "웃는 낮에 침 뱉으랴"
문장(49) = "우물을 파도 한 우물을 파라"
문장(50) = "웃물이맑아야 아랫물이 맑다"
a = Int(Rnd(1) * 50)
Label6.Caption = 문장(a)
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Label6.Visible = False
k = Len(문장(a))
v = Text1.Text
For i = 0 To k
Label1(i).Visible = False
Next
Timer2.Enabled = False
Label10.Caption = 0
Label13.Caption = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Timer2.Enabled = True
Label2.Caption = Text1.Text
If KeyAscii = 13 Then
Label13.Caption = Label13.Caption + 1
        For i = 0 To k
        Label1(i).Caption = ""
        Next
       ' Label10.Caption = Int((d / k) * 0.1)
        If Label6.Caption = Text1.Text Then
           Label4.Caption = Label4.Caption + 1
           a = Int(Rnd(1) * 50)
           Label6.Caption = 문장(a)
           k = Len(문장(a))
           For i = 0 To k
           Label1(i).Caption = "x"
           Next
           Label8.Caption = "맞았다"
        ElseIf v <> Label6.Caption Then
         Label8.Caption = "틀렸다"
           a = Int(Rnd(1) * 50)
           Label6.Caption = 문장(a)
           k = Len(문장(a))
           For i = 0 To k
           Label1(i).Caption = "x"
                       
            Beep
           Next
        End If
    Text1.Text = ""
    Text1.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
k = Len(Label6.Caption)

For i = 0 To k
    If Left(Text1.Text, i + 1) = Left(Label6.Caption, i + 1) Then
        Label1(i).Caption = "o"
        End If
         Next
For i = 0 To k
If Left(Text1.Text, i + 1) <> Left(Label6.Caption, i + 1) Then
       Label1(i).Caption = "x"
        End If
        Next

End Sub

Private Sub Timer2_Timer()
'd = Timer2.Interval + Timer2.Interval + 0.1
'Label12.Caption = Int((Val(Label13.Caption) / Val(Label4.Caption)) * 100) * 1 & "%"

End Sub
