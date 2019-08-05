VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  '없음
   Caption         =   "알람설정"
   ClientHeight    =   5235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Label Label14 
      BackStyle       =   0  '투명
      Height          =   375
      Left            =   1560
      MousePointer    =   2  '십자형
      TabIndex        =   31
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '투명
      Height          =   375
      Left            =   120
      MousePointer    =   2  '십자형
      TabIndex        =   30
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   120
      MousePointer    =   2  '십자형
      TabIndex        =   29
      Top             =   3680
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   120
      MousePointer    =   2  '십자형
      TabIndex        =   28
      Top             =   3330
      Width           =   855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   120
      MousePointer    =   2  '십자형
      TabIndex        =   27
      Top             =   2970
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   120
      MousePointer    =   2  '십자형
      TabIndex        =   26
      Top             =   2620
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   120
      MousePointer    =   2  '십자형
      TabIndex        =   25
      Top             =   2260
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   120
      MousePointer    =   2  '십자형
      TabIndex        =   24
      Top             =   1900
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Height          =   280
      Left            =   120
      MousePointer    =   2  '십자형
      TabIndex        =   23
      Top             =   1520
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Height          =   280
      Left            =   120
      MousePointer    =   2  '십자형
      TabIndex        =   22
      Top             =   1160
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Height          =   270
      Left            =   120
      MousePointer    =   2  '십자형
      TabIndex        =   21
      Top             =   820
      Width           =   900
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Height          =   300
      Left            =   120
      MousePointer    =   2  '십자형
      TabIndex        =   20
      Top             =   420
      Width           =   855
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Index           =   9
      Left            =   1680
      TabIndex        =   19
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Index           =   8
      Left            =   1680
      TabIndex        =   18
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Index           =   7
      Left            =   1680
      TabIndex        =   17
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Index           =   5
      Left            =   1680
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   165
      Index           =   9
      Left            =   1080
      TabIndex        =   9
      Top             =   3720
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   165
      Index           =   8
      Left            =   1080
      TabIndex        =   8
      Top             =   3360
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   165
      Index           =   7
      Left            =   1080
      TabIndex        =   7
      Top             =   3000
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   165
      Index           =   6
      Left            =   1080
      TabIndex        =   6
      Top             =   2640
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   165
      Index           =   5
      Left            =   1080
      TabIndex        =   5
      Top             =   2280
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   165
      Index           =   4
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   165
      Index           =   3
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   165
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   165
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   165
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   5250
      Left            =   0
      Picture         =   "game3.frx":0000
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label10_Click()
오전오후(7) = Form1.Combo1.Text
시간(7) = Form1.Text1.Text
분(7) = Form1.Text2.Text
초(7) = Form1.Text4.Text
If Form1.Option2.Value = True Then
메세지(7) = Form1.Text3.Text
End If
If Form1.Option1.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(7) & 시간(7) & "시" & 분(7) & "분 소리로 설정돼었습니다"
Label1(7).Caption = 오전오후(7) & 시간(7) & "시" & 분(7) & "분 소리"
Label2(7).Caption = "소리"
End If
If Form1.Option2.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(7) & 시간(7) & "시" & 분(7) & 메세지(7) & "라는 음성으로 설정돼었습니다"
Label1(7).Caption = 오전오후(7) & 시간(7) & "시" & 분(7) & "분" & " " & 메세지(7) & "라는 음성"
Label2(7).Caption = 메세지(7)
End If
Form1.Show
Form3.Hide

End Sub

Private Sub Label11_Click()
오전오후(8) = Form1.Combo1.Text
시간(8) = Form1.Text1.Text
분(8) = Form1.Text2.Text
초(8) = Form1.Text4.Text
If Form1.Option2.Value = True Then
메세지(8) = Form1.Text3.Text
End If
If Form1.Option1.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(8) & 시간(8) & "시" & 분(8) & "분 소리로 설정돼었습니다"
Label1(8).Caption = 오전오후(8) & 시간(8) & "시" & 분(8) & "분 소리"
Label2(8).Caption = "소리"
End If
If Form1.Option2.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(8) & 시간(8) & "시" & 분(8) & 메세지(8) & "라는 음성으로 설정돼었습니다"
Label1(8).Caption = 오전오후(8) & 시간(8) & "시" & 분(8) & "분" & " " & 메세지(8) & "라는 음성"
Label2(8).Caption = 메세지(8)
End If
Form1.Show
Form3.Hide

End Sub

Private Sub Label12_Click()
오전오후(9) = Form1.Combo1.Text
시간(9) = Form1.Text1.Text
분(9) = Form1.Text2.Text
초(9) = Form1.Text4.Text
If Form1.Option2.Value = True Then
메세지(9) = Form1.Text3.Text
End If
If Form1.Option1.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(9) & 시간(9) & "시" & 분(9) & "분 소리로 설정돼었습니다"
Label1(9).Caption = 오전오후(9) & 시간(9) & "시" & 분(9) & "분 소리"
Label2(9).Caption = "소리"
End If
If Form1.Option2.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(9) & 시간(9) & "시" & 분(9) & 메세지(9) & "라는 음성으로 설정돼었습니다"
Label1(9).Caption = 오전오후(9) & 시간(9) & "시" & 분(9) & "분" & " " & 메세지(9) & "라는 음성"
Label2(9).Caption = 메세지(9)
End If
Form1.Show
Form3.Hide
End Sub

Private Sub Label13_Click()
Form1.Show
Form3.Hide
End Sub

Private Sub Label14_Click()
For i = 0 To 9
Label1(i).Caption = ""
Label2(i).Caption = ""
오전오후(i) = ""
시간(i) = ""
분(i) = ""
Next
End Sub

Private Sub Label3_Click()
오전오후(0) = Form1.Combo1.Text
시간(0) = Form1.Text1.Text
분(0) = Form1.Text2.Text
초(0) = Form1.Text4.Text
If Form1.Option2.Value = True Then
메세지(0) = Form1.Text3.Text
End If
If Form1.Option1.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(0) & 시간(0) & "시" & 분(0) & "분 소리로 설정돼었습니다"
Label1(0).Caption = 오전오후(0) & 시간(0) & "시" & 분(0) & "분 소리"
Label2(0).Caption = "소리"
End If
If Form1.Option2.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(0) & 시간(0) & "시" & 분(0) & 메세지(0) & "라는 음성으로 설정돼었습니다"
Label1(0).Caption = 오전오후(0) & 시간(0) & "시" & 분(0) & "분" & " " & 메세지(0) & "라는 음성"
Label2(0).Caption = 메세지(0)
End If
Form1.Show
Form3.Hide
End Sub

Private Sub Label4_Click()
오전오후(1) = Form1.Combo1.Text
시간(1) = Form1.Text1.Text
분(1) = Form1.Text2.Text
초(1) = Form1.Text4.Text
If Form1.Option2.Value = True Then
메세지(1) = Form1.Text3.Text
End If
If Form1.Option1.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(1) & 시간(1) & "시" & 분(1) & "분 소리로 설정돼었습니다"
Label1(1).Caption = 오전오후(1) & 시간(1) & "시" & 분(1) & "분 소리"
Label2(1).Caption = "소리"
End If
If Form1.Option2.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(1) & 시간(1) & "시" & 분(1) & 메세지(1) & "라는 음성으로 설정돼었습니다"
Label1(1).Caption = 오전오후(1) & 시간(1) & "시" & 분(1) & "분" & " " & 메세지(1) & "라는 음성"
Label2(1).Caption = 메세지(1)
End If
Form1.Show
Form3.Hide

End Sub

Private Sub Label5_Click()
오전오후(2) = Form1.Combo1.Text
시간(2) = Form1.Text1.Text
분(2) = Form1.Text2.Text
초(2) = Form1.Text4.Text
If Form1.Option2.Value = True Then
메세지(2) = Form1.Text3.Text
End If
If Form1.Option1.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(2) & 시간(2) & "시" & 분(2) & "분 소리로 설정돼었습니다"
Label1(2).Caption = 오전오후(2) & 시간(2) & "시" & 분(2) & "분 소리"
Label2(2).Caption = "소리"
End If
If Form1.Option2.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(2) & 시간(2) & "시" & 분(2) & 메세지(2) & "라는 음성으로 설정돼었습니다"
Label1(2).Caption = 오전오후(2) & 시간(2) & "시" & 분(2) & "분" & " " & 메세지(2) & "라는 음성"
Label2(2).Caption = 메세지(2)
End If
Form1.Show
Form3.Hide
End Sub

Private Sub Label6_Click()
오전오후(3) = Form1.Combo1.Text
시간(3) = Form1.Text1.Text
분(3) = Form1.Text2.Text
초(3) = Form1.Text4.Text
If Form1.Option2.Value = True Then
메세지(3) = Form1.Text3.Text
End If
If Form1.Option1.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(3) & 시간(3) & "시" & 분(3) & "분 소리로 설정돼었습니다"
Label1(3).Caption = 오전오후(3) & 시간(3) & "시" & 분(3) & "분 소리"
Label2(3).Caption = "소리"
End If
If Form1.Option2.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(3) & 시간(3) & "시" & 분(3) & 메세지(3) & "라는 음성으로 설정돼었습니다"
Label1(3).Caption = 오전오후(3) & 시간(3) & "시" & 분(3) & "분" & " " & 메세지(3) & "라는 음성"
Label2(3).Caption = 메세지(3)
End If
Form1.Show
Form3.Hide
End Sub

Private Sub Label7_Click()
오전오후(4) = Form1.Combo1.Text
시간(4) = Form1.Text1.Text
분(4) = Form1.Text2.Text
초(4) = Form1.Text4.Text
If Form1.Option2.Value = True Then
메세지(4) = Form1.Text3.Text
End If
If Form1.Option1.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(4) & 시간(4) & "시" & 분(4) & "분 소리로 설정돼었습니다"
Label1(4).Caption = 오전오후(4) & 시간(4) & "시" & 분(4) & "분 소리"
Label2(4).Caption = "소리"
End If
If Form1.Option2.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(4) & 시간(4) & "시" & 분(4) & 메세지(4) & "라는 음성으로 설정돼었습니다"
Label1(4).Caption = 오전오후(4) & 시간(4) & "시" & 분(4) & "분" & " " & 메세지(4) & "라는 음성"
Label2(4).Caption = 메세지(4)
End If
Form1.Show
Form3.Hide

End Sub

Private Sub Label8_Click()
오전오후(5) = Form1.Combo1.Text
시간(5) = Form1.Text1.Text
분(5) = Form1.Text2.Text
초(5) = Form1.Text4.Text
If Form1.Option2.Value = True Then
메세지(5) = Form1.Text3.Text
End If
If Form1.Option1.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(5) & 시간(5) & "시" & 분(5) & "분 소리로 설정돼었습니다"
Label1(5).Caption = 오전오후(5) & 시간(5) & "시" & 분(5) & "분 소리"
Label2(5).Caption = "소리"
End If
If Form1.Option2.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(5) & 시간(5) & "시" & 분(5) & 메세지(5) & "라는 음성으로 설정돼었습니다"
Label1(5).Caption = 오전오후(5) & 시간(5) & "시" & 분(5) & "분" & " " & 메세지(5) & "라는 음성"
Label2(5).Caption = 메세지(5)
End If
Form1.Show
Form3.Hide
End Sub

Private Sub Label9_Click()
오전오후(6) = Form1.Combo1.Text
시간(6) = Form1.Text1.Text
분(6) = Form1.Text2.Text
초(6) = Form1.Text4.Text
If Form1.Option2.Value = True Then
메세지(6) = Form1.Text3.Text
End If
If Form1.Option1.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(6) & 시간(6) & "시" & 분(6) & "분 소리로 설정돼었습니다"
Label1(6).Caption = 오전오후(6) & 시간(6) & "시" & 분(6) & "분 소리"
Label2(6).Caption = "소리"
End If
If Form1.Option2.Value = True Then
Form1.Agent1.Characters("Genie").Speak 오전오후(6) & 시간(6) & "시" & 분(6) & 메세지(6) & "라는 음성으로 설정돼었습니다"
Label1(6).Caption = 오전오후(6) & 시간(6) & "시" & 분(6) & "분" & " " & 메세지(6) & "라는 음성"
Label2(6).Caption = 메세지(6)
End If
Form1.Show
Form3.Hide

End Sub
