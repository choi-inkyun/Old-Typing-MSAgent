VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   0  '없음
   Caption         =   "달리기"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
   Icon            =   "타자 연습 V1.0 9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer12 
      Interval        =   100
      Left            =   5880
      Top             =   3360
   End
   Begin VB.Timer Timer11 
      Interval        =   10
      Left            =   5400
      Top             =   3360
   End
   Begin VB.Timer Timer10 
      Interval        =   1
      Left            =   4920
      Top             =   3360
   End
   Begin VB.CheckBox Check2 
      Height          =   180
      Left            =   7680
      TabIndex        =   22
      Top             =   4630
      Width           =   180
   End
   Begin VB.CheckBox Check1 
      Height          =   180
      Left            =   7680
      TabIndex        =   21
      Top             =   4240
      Value           =   1  '확인
      Width           =   180
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "초급자"
      Height          =   180
      Left            =   9000
      TabIndex        =   20
      Top             =   4730
      Width           =   180
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   9000
      TabIndex        =   19
      Top             =   4470
      Value           =   -1  'True
      Width           =   180
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   9000
      TabIndex        =   18
      Top             =   4220
      Width           =   180
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3960
      TabIndex        =   12
      Top             =   4560
      Width           =   2805
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   600
      Top             =   3360
   End
   Begin VB.Timer Timer2 
      Interval        =   30
      Left            =   1080
      Top             =   3360
   End
   Begin VB.Timer Timer3 
      Interval        =   30
      Left            =   1560
      Top             =   3360
   End
   Begin VB.Timer Timer4 
      Interval        =   30
      Left            =   2040
      Top             =   3360
   End
   Begin VB.Timer Timer5 
      Interval        =   30
      Left            =   2520
      Top             =   3360
   End
   Begin VB.Timer Timer6 
      Interval        =   30
      Left            =   3000
      Top             =   3360
   End
   Begin VB.Timer Timer7 
      Interval        =   100
      Left            =   3480
      Top             =   3360
   End
   Begin VB.Timer Timer8 
      Interval        =   200
      Left            =   3960
      Top             =   3360
   End
   Begin VB.Timer Timer9 
      Interval        =   1000
      Left            =   4440
      Top             =   3360
   End
   Begin VB.Label Label28 
      BackStyle       =   0  '투명
      Caption         =   "Player"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   20.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   4800
      TabIndex        =   38
      Top             =   2770
      Width           =   1215
   End
   Begin VB.Label Label28 
      BackStyle       =   0  '투명
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   4320
      TabIndex        =   37
      Top             =   2320
      Width           =   495
   End
   Begin VB.Label Label28 
      BackStyle       =   0  '투명
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3920
      TabIndex        =   36
      Top             =   1910
      Width           =   495
   End
   Begin VB.Label Label28 
      BackStyle       =   0  '투명
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3540
      TabIndex        =   35
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label28 
      BackStyle       =   0  '투명
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3120
      TabIndex        =   34
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label Label28 
      BackStyle       =   0  '투명
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2750
      TabIndex        =   33
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '투명
      Height          =   735
      Left            =   9000
      MousePointer    =   2  '십자형
      TabIndex        =   32
      Top             =   5280
      Width           =   975
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   960
      Picture         =   "타자 연습 V1.0 9.frx":030A
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   960
      Picture         =   "타자 연습 V1.0 9.frx":0B4C
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   960
      Picture         =   "타자 연습 V1.0 9.frx":138E
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   960
      Picture         =   "타자 연습 V1.0 9.frx":1BD0
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   960
      Picture         =   "타자 연습 V1.0 9.frx":2412
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   960
      Picture         =   "타자 연습 V1.0 9.frx":2C54
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label20 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "초급자"
      Height          =   255
      Left            =   2520
      TabIndex        =   31
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label23 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "99"
      Height          =   255
      Left            =   2640
      TabIndex        =   30
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label26 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   2520
      TabIndex        =   29
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label19 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "중급자"
      Height          =   255
      Left            =   1440
      TabIndex        =   28
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label22 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "99"
      Height          =   255
      Left            =   1560
      TabIndex        =   27
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label Label25 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   1440
      TabIndex        =   26
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label18 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "상급자"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label21 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "99"
      Height          =   255
      Left            =   480
      TabIndex        =   24
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label24 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label27 
      BackStyle       =   0  '투명
      Height          =   615
      Left            =   5640
      MousePointer    =   2  '십자형
      TabIndex        =   17
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '투명
      Height          =   615
      Left            =   4440
      MousePointer    =   2  '십자형
      TabIndex        =   16
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   15
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
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
      Left            =   4920
      TabIndex        =   14
      Top             =   3930
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "단어 :"
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
      Left            =   4080
      TabIndex        =   13
      Top             =   3930
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   11
      Top             =   2770
      Width           =   495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   10
      Top             =   3010
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   9
      Top             =   2320
      Width           =   495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   8
      Top             =   2530
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   7
      Top             =   1900
      Width           =   495
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   6
      Top             =   2150
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   5
      Top             =   1450
      Width           =   495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   3
      Top             =   1030
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   0
      Top             =   840
      Width           =   495
   End
   Begin VB.Image Image8 
      Height          =   6000
      Left            =   0
      Picture         =   "타자 연습 V1.0 9.frx":3496
      Top             =   0
      Width           =   10500
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 낱말(200), k, b, c, d, e, f, g, h, i
Dim cnt As Integer
Dim aaa1 As Byte
Dim aaa2 As Byte
Dim aaa3 As Byte
Dim nam
Dim number

Private Sub Check1_Click()
If Check1.Value = 1 Then
Check2.Value = 0
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Check1.Value = 0
End If
End Sub

Private Sub Form_Load()
Timer11.Enabled = False
number = 0
Label10.Visible = True
Label9.Caption = " 나라에 달리기 대회가 열였습니다. 선수들 준비."
Open App.Path & "\txt\타자글.txt" For Input As #1
For i = 1 To 200
    Input #1, 낱말(i)
Next
Close #1

Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Label1.Visible = False
k = Len(낱말(a))
b = k * 40
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False


Text1.Enabled = False
Timer8.Enabled = False
cnt = 0
Timer9.Enabled = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
mumain.Agent1.Characters("Genie").Stop

End Sub

Private Sub Label10_Click()
mumain.Agent1.Characters("Genie").Stop
mumain.Show
Unload Me
End Sub

Private Sub Label17_Click()
Timer11.Enabled = True
Text1.IMEMode = vbIMEModeHangul '한글
Text1.Enabled = True
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Timer7.Enabled = True
Timer9.Enabled = True
Label1.Visible = True
Text1.SetFocus
Text1.Text = ""
a = Int(Rnd(1) * 200)
Label1.Caption = 낱말(a)
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label2.Visible = True
Label9.Caption = "경기가 시작되었습니다"
Timer8.Enabled = True
Label10.Visible = True
nam = "사용자"
If Check1.Value = 1 Then
        If mumain.Option12 = True Then
        mumain.Agent1.Characters("Genie").Speak Label1.Caption
        End If
End If

End Sub

Private Sub Label27_Click()
Timer11.Enabled = False
mumain.Agent1.Characters("Genie").Stop
Image1.Left = 960
Image2.Left = 960
Image3.Left = 960
Image4.Left = 960
Image5.Left = 960
Image6.Left = 960
Timer1.Enabled = False
Text1.Enabled = False
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
cnt = 0
Label1.Visible = False
End Sub

Private Sub Label9_Change()
If Check2.Value = 1 Then
        If mumain.Option12 = True Then
        mumain.Agent1.Characters("Genie").Speak Label9.Caption
        End If
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 32 Then
If Trim(Text1.Text) = Label1.Caption Then
a = Int(Rnd(1) * 200)
Label1.Caption = 낱말(a)
If Label1.Caption = "" Then
a = Int(Rnd(1) * 200)
Label1.Caption = 낱말(a)
End If
k = Len(낱말(a))
b = k * Int(Rnd(140) * 200)
Image6.Left = Image6.Left + b
Else
If Text1.Text <> "" Then
Image6.Left = Image6.Left - Int(Rnd(40) * 70)
a = Int(Rnd(1) * 200)
Label1.Caption = 낱말(a)
If Label1.Caption = "" Then
a = Int(Rnd(1) * 200)
Label1.Caption = 낱말(a)
End If
End If
End If
Text1.Text = ""
Text1.SetFocus
If Check1.Value = 1 Then
        If mumain.Option12 = True Then
        mumain.Agent1.Characters("Genie").Speak Label1.Caption
        End If
End If

End If
End Sub

Private Sub Timer1_Timer()


Image1.Left = Image1.Left + Int(Rnd(c) * d)
If Image1.Left >= 8800 Then
Label3.Visible = True
Label11.Visible = True
Timer1.Enabled = False
Image1.Left = 9150
If Label4.Caption = "" Then
If Label5.Caption = "" Then
If Label6.Caption = "" Then
If Label7.Caption = "" Then
If Label8.Caption = "" Then
Label3.Caption = "1등"
Label9.Caption = "1번이 1등을 하였습니다"
Label11.Caption = cnt & "초"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "초"
Label21.Visible = True
Label24.Caption = "1번"
End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "초"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "초"
Label22.Visible = True
Label25.Caption = "1번"

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "초"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "초"
Label23.Visible = True
Label26.Caption = "1번"
End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "초"
End If
End If


End If
End If
End If
End If
End If
If Label4.Caption = "1등" Then
Label3.Caption = "2등"
Label9.Caption = "1번는 2등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label5.Caption = "1등" Then
Label3.Caption = "2등"
Label9.Caption = "1번는 2등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label6.Caption = "1등" Then
Label3.Caption = "2등"
Label9.Caption = "1번는 2등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label7.Caption = "1등" Then
Label3.Caption = "2등"
Label9.Caption = "1번는 2등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label8.Caption = "1등" Then
Label3.Caption = "2등"
Label9.Caption = "1번는 2등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label4.Caption = "2등" Then
Label3.Caption = "3등"
Label9.Caption = "1번는 3등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label5.Caption = "2등" Then
Label3.Caption = "3등"
Label9.Caption = "1번는 3등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label6.Caption = "2등" Then
Label3.Caption = "3등"
Label9.Caption = "1번는 3등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label7.Caption = "2등" Then
Label3.Caption = "3등"
Label9.Caption = "1번는 3등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label8.Caption = "2등" Then
Label3.Caption = "3등"
Label9.Caption = "1번는 3등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label4.Caption = "3등" Then
Label3.Caption = "4등"
Label9.Caption = "1번는 4등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label5.Caption = "3등" Then
Label3.Caption = "4등"
Label9.Caption = "1번는 4등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label6.Caption = "3등" Then
Label3.Caption = "4등"
Label9.Caption = "1번는 4등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label7.Caption = "3등" Then
Label3.Caption = "4등"
Label9.Caption = "1번는 4등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label8.Caption = "3등" Then
Label3.Caption = "4등"
Label9.Caption = "1번는 4등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label4.Caption = "4등" Then
Label3.Caption = "5등"
Label9.Caption = "1번는 5등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label5.Caption = "4등" Then
Label3.Caption = "5등"
Label9.Caption = "1번는 5등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label6.Caption = "4등" Then
Label3.Caption = "5등"
Label9.Caption = "1번는 5등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label7.Caption = "4등" Then
Label3.Caption = "5등"
Label9.Caption = "1번는 5등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label8.Caption = "4등" Then
Label3.Caption = "5등"
Label9.Caption = "1번는 5등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label4.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "1번는 6등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label5.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "1번는 6등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label6.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "1번는 6등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label7.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "1번는 6등을 하였습니다"
Label11.Caption = cnt & "초"
End If
If Label8.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "1번는 6등을 하였습니다"
Label11.Caption = cnt & "초"
End If
Timer1.Enabled = False
If Timer1.Enabled = False Then
If Timer2.Enabled = False Then
If Timer3.Enabled = False Then
If Timer4.Enabled = False Then
If Timer5.Enabled = False Then
If Timer6.Enabled = False Then
Timer11.Enabled = False
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Timer11_Timer()
If number = 0 Then
If Timer1.Enabled = True Then
Image1.Picture = LoadPicture(App.Path & "\image\ani1-1.bmp")
End If
If Timer2.Enabled = True Then
Image2.Picture = LoadPicture(App.Path & "\image\ani2-1.bmp")
End If
If Timer3.Enabled = True Then
Image3.Picture = LoadPicture(App.Path & "\image\ani3-1.bmp")
End If
If Timer4.Enabled = True Then
Image4.Picture = LoadPicture(App.Path & "\image\ani4-1.bmp")
End If
If Timer5.Enabled = True Then
Image5.Picture = LoadPicture(App.Path & "\image\ani5-1.bmp")
End If
If Timer6.Enabled = True Then
Image6.Picture = LoadPicture(App.Path & "\image\ani6-1.bmp")
End If
End If
If number = 1 Then
If Timer1.Enabled = True Then
Image1.Picture = LoadPicture(App.Path & "\image\ani1-2.bmp")
End If
If Timer2.Enabled = True Then
Image2.Picture = LoadPicture(App.Path & "\image\ani2-2.bmp")
End If
If Timer3.Enabled = True Then
Image3.Picture = LoadPicture(App.Path & "\image\ani3-2.bmp")
End If
If Timer4.Enabled = True Then
Image4.Picture = LoadPicture(App.Path & "\image\ani4-2.bmp")
End If
If Timer5.Enabled = True Then
Image5.Picture = LoadPicture(App.Path & "\image\ani5-2.bmp")
End If
If Timer6.Enabled = True Then
Image6.Picture = LoadPicture(App.Path & "\image\ani6-2.bmp")
End If

End If
If number = 2 Then
If Timer1.Enabled = True Then
Image1.Picture = LoadPicture(App.Path & "\image\ani1-3.bmp")
End If
If Timer2.Enabled = True Then
Image2.Picture = LoadPicture(App.Path & "\image\ani2-3.bmp")
End If
If Timer3.Enabled = True Then
Image3.Picture = LoadPicture(App.Path & "\image\ani3-3.bmp")
End If
If Timer4.Enabled = True Then
Image4.Picture = LoadPicture(App.Path & "\image\ani4-3.bmp")
End If
If Timer5.Enabled = True Then
Image5.Picture = LoadPicture(App.Path & "\image\ani5-3.bmp")
End If
If Timer6.Enabled = True Then
Image6.Picture = LoadPicture(App.Path & "\image\ani6-3.bmp")
End If

End If
If number = 3 Then
If Timer1.Enabled = True Then
Image1.Picture = LoadPicture(App.Path & "\image\ani1-4.bmp")
End If
If Timer2.Enabled = True Then
Image2.Picture = LoadPicture(App.Path & "\image\ani2-4.bmp")
End If
If Timer3.Enabled = True Then
Image3.Picture = LoadPicture(App.Path & "\image\ani3-4.bmp")
End If
If Timer4.Enabled = True Then
Image4.Picture = LoadPicture(App.Path & "\image\ani4-4.bmp")
End If
If Timer5.Enabled = True Then
Image5.Picture = LoadPicture(App.Path & "\image\ani5-4.bmp")
End If
If Timer6.Enabled = True Then
Image6.Picture = LoadPicture(App.Path & "\image\ani6-4.bmp")
End If
End If
number = number + 1
If number >= 5 Then
number = 0
End If
End Sub

Private Sub Timer12_Timer()
If Timer2.Enabled = True Then
Check1.Enabled = False
Check2.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Label17.Enabled = False
For i = 0 To 5
Label28(i).Visible = False
Next
End If
If Timer2.Enabled = False Then
Check1.Enabled = True
Check2.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Label17.Enabled = True
For i = 0 To 5
Label28(i).Visible = True
Next
End If

End Sub

Private Sub Timer2_Timer()
Image2.Left = Image2.Left + Int(Rnd(d) * e)
If Image2.Left >= 8800 Then
Label4.Visible = True
Label12.Visible = True
d = 0
Timer2.Enabled = False
Image2.Left = 9150
If Label3.Caption = "" Then
If Label5.Caption = "" Then
If Label6.Caption = "" Then
If Label7.Caption = "" Then
If Label8.Caption = "" Then
Label4.Caption = "1등"
Label9.Caption = "2번이 1등을 하였습니다"
Label12.Caption = cnt & "초"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "초"
Label21.Visible = True
Label24.Caption = "2번"

End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "초"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "초"
Label22.Visible = True
Label25.Caption = "2번"

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "초"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "초"
Label23.Visible = True
Label26.Caption = "2번"

End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "초"
End If
End If

End If
End If
End If
End If
End If
If Label3.Caption = "1등" Then
Label4.Caption = "2등"
Label9.Caption = "2번는 2등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label5.Caption = "1등" Then
Label4.Caption = "2등"
Label9.Caption = "2번는 2등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label6.Caption = "1등" Then
Label4.Caption = "2등"
Label9.Caption = "2번는 2등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label7.Caption = "1등" Then
Label4.Caption = "2등"
Label9.Caption = "2번는 2등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label8.Caption = "1등" Then
Label4.Caption = "2등"
Label9.Caption = "2번는 2등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label3.Caption = "2등" Then
Label4.Caption = "3등"
Label9.Caption = "2번는 3등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label5.Caption = "2등" Then
Label4.Caption = "3등"
Label9.Caption = "2번는 3등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label6.Caption = "2등" Then
Label4.Caption = "3등"
Label9.Caption = "2번는 3등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label7.Caption = "2등" Then
Label4.Caption = "3등"
Label9.Caption = "2번는 3등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label8.Caption = "2등" Then
Label4.Caption = "3등"
Label9.Caption = "2번는 3등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label3.Caption = "3등" Then
Label4.Caption = "4등"
Label9.Caption = "2번는 4등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label5.Caption = "3등" Then
Label4.Caption = "4등"
Label9.Caption = "2번는 4등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label6.Caption = "3등" Then
Label4.Caption = "4등"
Label9.Caption = "2번는 4등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label7.Caption = "3등" Then
Label4.Caption = "4등"
Label9.Caption = "2번는 4등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label8.Caption = "3등" Then
Label4.Caption = "4등"
Label9.Caption = "2번는 4등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label3.Caption = "4등" Then
Label4.Caption = "5등"
Label9.Caption = "2번는 5등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label5.Caption = "4등" Then
Label4.Caption = "5등"
Label9.Caption = "2번는 5등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label6.Caption = "4등" Then
Label4.Caption = "5등"
Label9.Caption = "2번는 5등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label7.Caption = "4등" Then
Label4.Caption = "5등"
Label9.Caption = "2번는 5등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label8.Caption = "4등" Then
Label4.Caption = "5등"
Label9.Caption = "2번는 5등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label3.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "2번는 6등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label5.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "2번는 6등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label6.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "2번는 6등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label7.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "2번는 6등을 하였습니다"
Label12.Caption = cnt & "초"
End If
If Label8.Caption = "5등" Then
Label4.Caption = "6등"
Label9.Caption = "2번는 6등을 하였습니다"
Label12.Caption = cnt & "초"
End If
Timer2.Enabled = False
If Timer1.Enabled = False Then
If Timer2.Enabled = False Then
If Timer3.Enabled = False Then
If Timer4.Enabled = False Then
If Timer5.Enabled = False Then
If Timer6.Enabled = False Then
Timer11.Enabled = False
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Timer3_Timer()
Image3.Left = Image3.Left + Int(Rnd(e) * f)
If Image3.Left >= 8800 Then
Label5.Visible = True
Label13.Visible = True

Timer3.Enabled = False
e = 0
Image3.Left = 9150
If Label3.Caption = "" Then
If Label4.Caption = "" Then
If Label6.Caption = "" Then
If Label7.Caption = "" Then
If Label8.Caption = "" Then
Label5.Caption = "1등"
Label9.Caption = "3번이 1등을 하였습니다"
Label13.Caption = cnt & "초"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "초"
Label21.Visible = True
Label24.Caption = "3번"
End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "초"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "초"
Label22.Visible = True
Label25.Caption = "3번"
End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "초"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "초"
Label23.Visible = True
Label26.Caption = "3번"
End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "초"
End If
End If

End If
End If
End If
End If
End If
If Label3.Caption = "1등" Then
Label5.Caption = "2등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 2등을 하였습니다"
End If
If Label4.Caption = "1등" Then
Label5.Caption = "2등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 2등을 하였습니다"
End If
If Label6.Caption = "1등" Then
Label5.Caption = "2등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 2등을 하였습니다"
End If
If Label7.Caption = "1등" Then
Label5.Caption = "2등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 2등을 하였습니다"
End If
If Label8.Caption = "1등" Then
Label5.Caption = "2등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 2등을 하였습니다"
End If
If Label3.Caption = "2등" Then
Label5.Caption = "3등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 3등을 하였습니다"
End If
If Label4.Caption = "2등" Then
Label5.Caption = "3등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 3등을 하였습니다"
End If
If Label6.Caption = "2등" Then
Label5.Caption = "3등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 3등을 하였습니다"
End If
If Label7.Caption = "2등" Then
Label5.Caption = "3등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 3등을 하였습니다"
End If
If Label8.Caption = "2등" Then
Label5.Caption = "3등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 3등을 하였습니다"
End If
If Label3.Caption = "3등" Then
Label5.Caption = "4등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 4등을 하였습니다"
End If
If Label4.Caption = "3등" Then
Label5.Caption = "4등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 4등을 하였습니다"
End If
If Label6.Caption = "3등" Then
Label5.Caption = "4등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 4등을 하였습니다"
End If
If Label7.Caption = "3등" Then
Label5.Caption = "4등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 4등을 하였습니다"
End If
If Label8.Caption = "3등" Then
Label5.Caption = "4등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 4등을 하였습니다"
End If
If Label3.Caption = "4등" Then
Label5.Caption = "5등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 5등을 하였습니다"
End If
If Label4.Caption = "4등" Then
Label5.Caption = "5등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 5등을 하였습니다"
End If
If Label6.Caption = "4등" Then
Label5.Caption = "5등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 5등을 하였습니다"
End If
If Label7.Caption = "4등" Then
Label5.Caption = "5등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 5등을 하였습니다"
End If
If Label8.Caption = "4등" Then
Label5.Caption = "5등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 5등을 하였습니다"
End If
If Label3.Caption = "5등" Then
Label5.Caption = "6등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 6등을 하였습니다"
End If
If Label4.Caption = "5등" Then
Label5.Caption = "6등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 6등을 하였습니다"
End If
If Label6.Caption = "5등" Then
Label5.Caption = "6등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 6등을 하였습니다"
End If
If Label7.Caption = "5등" Then
Label5.Caption = "6등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 6등을 하였습니다"
End If
If Label8.Caption = "5등" Then
Label5.Caption = "6등"
Label13.Caption = cnt & "초"
Label9.Caption = "3번는 6등을 하였습니다"
End If
Timer3.Enabled = False
If Timer1.Enabled = False Then
If Timer2.Enabled = False Then
If Timer3.Enabled = False Then
If Timer4.Enabled = False Then
If Timer5.Enabled = False Then
If Timer6.Enabled = False Then
Timer11.Enabled = False
End If
End If
End If
End If
End If
End If
End If

End Sub

Private Sub Timer4_Timer()
Image4.Left = Image4.Left + Int(Rnd(f) * g)
If Image4.Left >= 8800 Then
Label6.Visible = True
Label14.Visible = True
Timer4.Enabled = False
f = 0
Image4.Left = 9150
If Label3.Caption = "" Then
If Label4.Caption = "" Then
If Label5.Caption = "" Then
If Label7.Caption = "" Then
If Label8.Caption = "" Then
Label6.Caption = "1등"
Label9.Caption = "4번이 1등을 하였습니다"
Label14.Caption = cnt & "초"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "초"
Label21.Visible = True
Label24.Caption = "4번"

End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "초"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "초"
Label22.Visible = True
Label25.Caption = "4번"

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "초"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "초"
Label23.Visible = True
Label26.Caption = "4번"
End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "초"
End If
End If

End If
End If
End If
End If
End If
If Label3.Caption = "1등" Then
Label6.Caption = "2등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 2등을 하였습니다"
End If
If Label4.Caption = "1등" Then
Label6.Caption = "2등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 2등을 하였습니다"
End If
If Label5.Caption = "1등" Then
Label6.Caption = "2등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 2등을 하였습니다"
End If
If Label7.Caption = "1등" Then
Label6.Caption = "2등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 2등을 하였습니다"
End If
If Label8.Caption = "1등" Then
Label6.Caption = "2등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 2등을 하였습니다"
End If
If Label3.Caption = "2등" Then
Label6.Caption = "3등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 3등을 하였습니다"
End If
If Label4.Caption = "2등" Then
Label6.Caption = "3등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 3등을 하였습니다"
End If
If Label5.Caption = "2등" Then
Label6.Caption = "3등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 3등을 하였습니다"
End If
If Label7.Caption = "2등" Then
Label6.Caption = "3등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 3등을 하였습니다"
End If
If Label8.Caption = "2등" Then
Label6.Caption = "3등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 3등을 하였습니다"
End If
If Label3.Caption = "3등" Then
Label6.Caption = "4등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 4등을 하였습니다"
End If
If Label4.Caption = "3등" Then
Label6.Caption = "4등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 4등을 하였습니다"
End If
If Label5.Caption = "3등" Then
Label6.Caption = "4등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 4등을 하였습니다"
End If
If Label7.Caption = "3등" Then
Label6.Caption = "4등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 4등을 하였습니다"
End If
If Label8.Caption = "3등" Then
Label6.Caption = "4등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 4등을 하였습니다"
End If
If Label3.Caption = "4등" Then
Label6.Caption = "5등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 5등을 하였습니다"
End If
If Label4.Caption = "4등" Then
Label6.Caption = "5등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 5등을 하였습니다"
End If
If Label5.Caption = "4등" Then
Label6.Caption = "5등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 5등을 하였습니다"
End If
If Label7.Caption = "4등" Then
Label6.Caption = "5등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 5등을 하였습니다"
End If
If Label8.Caption = "4등" Then
Label6.Caption = "5등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 5등을 하였습니다"
End If
If Label3.Caption = "5등" Then
Label6.Caption = "6등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 6등을 하였습니다"
End If
If Label4.Caption = "5등" Then
Label6.Caption = "6등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 6등을 하였습니다"
End If
If Label5.Caption = "5등" Then
Label6.Caption = "6등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 6등을 하였습니다"
End If
If Label7.Caption = "5등" Then
Label6.Caption = "6등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 6등을 하였습니다"
End If
If Label8.Caption = "5등" Then
Label6.Caption = "6등"
Label14.Caption = cnt & "초"
Label9.Caption = "4번는 6등을 하였습니다"
End If
Timer4.Enabled = False
If Timer1.Enabled = False Then
If Timer2.Enabled = False Then
If Timer3.Enabled = False Then
If Timer4.Enabled = False Then
If Timer5.Enabled = False Then
If Timer6.Enabled = False Then
Timer11.Enabled = False
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Timer5_Timer()
Image5.Left = Image5.Left + Int(Rnd(g) * c)
If Image5.Left >= 8800 Then
Label7.Visible = True
Label15.Visible = True
Timer5.Enabled = False
g = 0
Image5.Left = 9150
If Label3.Caption = "" Then
If Label4.Caption = "" Then
If Label5.Caption = "" Then
If Label6.Caption = "" Then
If Label8.Caption = "" Then
Label7.Caption = "1등"
Label9.Caption = "5번이 1등을 하였습니다"
Label15.Caption = cnt & "초"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "초"
Label21.Visible = True
Label24.Caption = "5번"
End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "초"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "초"
Label22.Visible = True
Label25.Caption = "5번"

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "초"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "초"
Label23.Visible = True
Label26.Caption = "5번"

End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "초"
End If
End If

End If
End If
End If
End If
End If
If Label7.Caption = "1등" Then
Label7.Caption = "2등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 2등을 하였습니다"
End If
If Label4.Caption = "1등" Then
Label7.Caption = "2등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 2등을 하였습니다"
End If
If Label5.Caption = "1등" Then
Label7.Caption = "2등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 2등을 하였습니다"
End If
If Label6.Caption = "1등" Then
Label7.Caption = "2등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 2등을 하였습니다"
End If
If Label8.Caption = "1등" Then
Label7.Caption = "2등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 2등을 하였습니다"
End If
If Label3.Caption = "2등" Then
Label7.Caption = "3등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 3등을 하였습니다"
End If
If Label4.Caption = "2등" Then
Label7.Caption = "3등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 3등을 하였습니다"
End If
If Label5.Caption = "2등" Then
Label7.Caption = "3등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 3등을 하였습니다"
End If
If Label6.Caption = "2등" Then
Label7.Caption = "3등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 3등을 하였습니다"
End If
If Label8.Caption = "2등" Then
Label7.Caption = "3등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 3등을 하였습니다"
End If
If Label3.Caption = "3등" Then
Label7.Caption = "4등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 4등을 하였습니다"
End If
If Label4.Caption = "3등" Then
Label7.Caption = "4등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 4등을 하였습니다"
End If
If Label5.Caption = "3등" Then
Label7.Caption = "4등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 4등을 하였습니다"
End If
If Label6.Caption = "3등" Then
Label7.Caption = "4등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 4등을 하였습니다"
End If
If Label8.Caption = "3등" Then
Label7.Caption = "4등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 4등을 하였습니다"
End If
If Label3.Caption = "4등" Then
Label7.Caption = "5등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 5등을 하였습니다"
End If
If Label4.Caption = "4등" Then
Label7.Caption = "5등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 5등을 하였습니다"
End If
If Label5.Caption = "4등" Then
Label7.Caption = "5등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 5등을 하였습니다"
End If
If Label6.Caption = "4등" Then
Label7.Caption = "5등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 5등을 하였습니다"
End If
If Label8.Caption = "4등" Then
Label7.Caption = "5등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 5등을 하였습니다"
End If
If Label3.Caption = "5등" Then
Label7.Caption = "6등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 6등을 하였습니다"
End If
If Label4.Caption = "5등" Then
Label7.Caption = "6등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 6등을 하였습니다"
End If
If Label5.Caption = "5등" Then
Label7.Caption = "6등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 6등을 하였습니다"
End If
If Label6.Caption = "5등" Then
Label7.Caption = "6등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 6등을 하였습니다"
End If
If Label8.Caption = "5등" Then
Label7.Caption = "6등"
Label15.Caption = cnt & "초"
Label9.Caption = "5번은 6등을 하였습니다"
End If
Timer5.Enabled = False
If Timer1.Enabled = False Then
If Timer2.Enabled = False Then
If Timer3.Enabled = False Then
If Timer4.Enabled = False Then
If Timer5.Enabled = False Then
If Timer6.Enabled = False Then
Timer11.Enabled = False
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Timer6_Timer()
If Image6.Left >= 8800 Then
Label8.Visible = True
Label16.Visible = True
Timer6.Enabled = False
b = 0
Label1.Visible = False
Image6.Left = 9150
Text1.Enabled = False
If Label3.Caption = "" Then
If Label4.Caption = "" Then
If Label5.Caption = "" Then
If Label6.Caption = "" Then
If Label7.Caption = "" Then
Label8.Caption = "1등"
Label9.Caption = "" & nam & "가 1등을 하였습니다"
mumain.Agent1.Characters("Genie").Speak "축하합니다"
mumain.Agent1.Characters("Genie").Play "Congratulate_2"
Label16.Caption = cnt & "초"

If Option1.Value = True Then
If aaa1 > cnt Then
Label21.Caption = cnt & "초"
Label21.Visible = True
Label24.Caption = nam

End If
If aaa1 <= cnt Then
Label21.Caption = aaa1 & "초"
End If
End If
If Option2.Value = True Then
If aaa2 > cnt Then
Label22.Caption = cnt & "초"
Label22.Visible = True
Label25.Caption = nam

End If
If aaa2 <= cnt Then
Label22.Caption = aaa2 & "초"
End If
End If
If Option3.Value = True Then
If aaa3 > cnt Then
Label23.Caption = cnt & "초"
Label23.Visible = True
Label26.Caption = nam

End If
If aaa3 <= cnt Then
Label23.Caption = aaa3 & "초"
End If
End If

End If
End If
End If
End If
End If
If Label3.Caption = "1등" Then
Label8.Caption = "2등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 2등을 하였습니다"
End If
If Label4.Caption = "1등" Then
Label8.Caption = "2등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 2등을 하였습니다"
End If
If Label5.Caption = "1등" Then
Label8.Caption = "2등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 2등을 하였습니다"
End If
If Label6.Caption = "1등" Then
Label8.Caption = "2등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 2등을 하였습니다"
End If
If Label7.Caption = "1등" Then
Label8.Caption = "2등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 2등을 하였습니다"
End If
If Label3.Caption = "2등" Then
Label8.Caption = "3등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 3등을 하였습니다"
End If
If Label4.Caption = "2등" Then
Label8.Caption = "3등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 3등을 하였습니다"
End If
If Label5.Caption = "2등" Then
Label8.Caption = "3등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 3등을 하였습니다"
End If
If Label6.Caption = "2등" Then
Label8.Caption = "3등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 3등을 하였습니다"
End If
If Label7.Caption = "2등" Then
Label8.Caption = "3등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 3등을 하였습니다"
End If
If Label3.Caption = "3등" Then
Label8.Caption = "4등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 4등을 하였습니다"
End If
If Label4.Caption = "3등" Then
Label8.Caption = "4등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 4등을 하였습니다"
End If
If Label5.Caption = "3등" Then
Label8.Caption = "4등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 4등을 하였습니다"
End If
If Label6.Caption = "3등" Then
Label8.Caption = "4등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 4등을 하였습니다"
End If
If Label7.Caption = "3등" Then
Label8.Caption = "4등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 4등을 하였습니다"
End If
If Label3.Caption = "4등" Then
Label8.Caption = "5등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 5등을 하였습니다"
End If
If Label4.Caption = "4등" Then
Label8.Caption = "5등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 5등을 하였습니다"
End If
If Label5.Caption = "4등" Then
Label8.Caption = "5등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 5등을 하였습니다"
End If
If Label6.Caption = "4등" Then
Label8.Caption = "5등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 5등을 하였습니다"
End If
If Label7.Caption = "4등" Then
Label8.Caption = "5등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 5등을 하였습니다"
End If
If Label3.Caption = "5등" Then
Label8.Caption = "6등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 6등을 하였습니다"
End If
If Label4.Caption = "5등" Then
Label8.Caption = "6등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 6등을 하였습니다"
End If
If Label5.Caption = "5등" Then
Label8.Caption = "6등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 6등을 하였습니다"
End If
If Label6.Caption = "5등" Then
Label8.Caption = "6등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 6등을 하였습니다"
End If
If Label7.Caption = "5등" Then
Label8.Caption = "6등"
Label16.Caption = cnt & "초"
Label9.Caption = "" & nam & "는 6등을 하였습니다"
End If
Timer6.Enabled = False
If Timer1.Enabled = False Then
If Timer2.Enabled = False Then
If Timer3.Enabled = False Then
If Timer4.Enabled = False Then
If Timer5.Enabled = False Then
If Timer6.Enabled = False Then
Timer11.Enabled = False
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Timer7_Timer()
If Option1.Value = True Then
c = Int(Rnd(10) * 40)
d = Int(Rnd(10) * 40)
e = Int(Rnd(10) * 40)
f = Int(Rnd(10) * 40)
g = Int(Rnd(10) * 40)
End If
If Option2.Value = True Then
c = Int(Rnd(5) * 30)
d = Int(Rnd(5) * 30)
e = Int(Rnd(5) * 30)
f = Int(Rnd(5) * 30)
g = Int(Rnd(5) * 30)
End If
If Option3.Value = True Then
c = Int(Rnd(1) * 20)
d = Int(Rnd(1) * 20)
e = Int(Rnd(1) * 20)
f = Int(Rnd(1) * 20)
g = Int(Rnd(1) * 20)
End If
aaa1 = Val(Label21.Caption)
aaa2 = Val(Label22.Caption)
aaa3 = Val(Label23.Caption)
End Sub

Private Sub Timer8_Timer()
If Image1.Left > Image2.Left Then
If Image1.Left > Image3.Left Then
If Image1.Left > Image4.Left Then
If Image1.Left > Image5.Left Then
If Image1.Left > Image6.Left Then
Select Case Image1.Left
Case 960 To 7200
Label9.Caption = "현재 1번이 선두로 달리고 있습니다"
End Select
End If
End If
End If
End If
End If

If Image2.Left > Image1.Left Then
If Image2.Left > Image3.Left Then
If Image2.Left > Image4.Left Then
If Image2.Left > Image5.Left Then
If Image2.Left > Image6.Left Then
Select Case Image2.Left
Case 960 To 7200
Label9.Caption = "현재 2번이 선두로 달리고 있습니다"
End Select
End If
End If
End If
End If
End If

If Image3.Left > Image1.Left Then
If Image3.Left > Image2.Left Then
If Image3.Left > Image4.Left Then
If Image3.Left > Image5.Left Then
If Image3.Left > Image6.Left Then
Select Case Image3.Left
Case 960 To 7200
Label9.Caption = "현재 3번이 선두로 달리고 있습니다"
End Select
End If
End If
End If
End If
End If

If Image4.Left > Image1.Left Then
If Image4.Left > Image2.Left Then
If Image4.Left > Image3.Left Then
If Image4.Left > Image5.Left Then
If Image4.Left > Image6.Left Then
Select Case Image4.Left
Case 960 To 7200
Label9.Caption = "현재 4번이 선두로 달리고 있습니다"
End Select
End If
End If
End If
End If
End If

If Image5.Left > Image1.Left Then
If Image5.Left > Image2.Left Then
If Image5.Left > Image3.Left Then
If Image5.Left > Image4.Left Then
If Image5.Left > Image6.Left Then
Select Case Image5.Left
Case 960 To 7200
Label9.Caption = "현재 5번이 선두로 달리고 있습니다"
End Select
End If
End If
End If
End If
End If

If Image6.Left > Image1.Left Then
If Image6.Left > Image2.Left Then
If Image6.Left > Image3.Left Then
If Image6.Left > Image4.Left Then
If Image6.Left > Image5.Left Then
Select Case Image6.Left
Case 960 To 7200
Label9.Caption = "현재 " & nam & "이 선두로 달리고 있습니다"
End Select
End If
End If
End If
End If
End If
If Image6.Left <= 500 Then
MsgBox "당신은 경기장에서 퇴장당했습니다" + Chr(13) + Chr(13) _
        + "그렇기 때문에 경기를 처음부터 시작합니다."
Image1.Left = 960
Image2.Left = 960
Image3.Left = 960
Image4.Left = 960
Image5.Left = 960
Image6.Left = 960
Timer1.Enabled = False
Text1.Enabled = False
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
cnt = 0
End If

End Sub

Private Sub Timer9_Timer()
cnt = cnt + 1
End Sub
