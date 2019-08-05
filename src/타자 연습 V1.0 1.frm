VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  '없음
   Caption         =   "산성비"
   ClientHeight    =   6000
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   10500
   FontTransparent =   0   'False
   Icon            =   "타자 연습 V1.0 1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   600
      Top             =   2760
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   2760
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   8680
      TabIndex        =   19
      Top             =   4720
      Width           =   180
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   8680
      TabIndex        =   18
      Top             =   4470
      Value           =   -1  'True
      Width           =   180
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   8680
      TabIndex        =   17
      Top             =   4210
      Width           =   180
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "바탕체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   16
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "한글 타자"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      MaskColor       =   &H00FF0000&
      Style           =   1  '그래픽
      TabIndex        =   15
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "영문 타자"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      Style           =   1  '그래픽
      TabIndex        =   14
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '투명
      Height          =   615
      Left            =   6240
      MousePointer    =   2  '십자형
      TabIndex        =   24
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '투명
      Height          =   735
      Left            =   5040
      MousePointer    =   2  '십자형
      TabIndex        =   23
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '투명
      Height          =   615
      Left            =   4200
      MousePointer    =   2  '십자형
      TabIndex        =   22
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '투명
      Height          =   375
      Left            =   7920
      MousePointer    =   2  '십자형
      TabIndex        =   21
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Height          =   735
      Left            =   9000
      MousePointer    =   2  '십자형
      TabIndex        =   20
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   13
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '투명
      Caption         =   "입력 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "남은 갯수 : "
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
      Left            =   360
      TabIndex        =   10
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "10"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "맞은 갯수 : "
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
      Left            =   360
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   4320
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   10440
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4560
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6000
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   7440
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   8760
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "타자 연습 V1.0 1.frx":030A
      Top             =   0
      Width           =   10500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 낱말(200), i, 맞춘갯수(10), 점수(10), 이름(10), 순위(10), 난이도(10), cnt, J, 영문(100), 종류(10)


Private Sub Form_Load()

Randomize


Command5.Enabled = False
For i = 0 To 6
Label1(i).Visible = False
Label1(i).Top = 0
Next
Label4.Caption = 0
Label5.Caption = 10
Timer1.Enabled = False
Open App.Path & "\txt\타자글.txt" For Input As #1
For i = 1 To 200
    Input #1, 낱말(i)
Next
Close #1
Text1.IMEMode = vbIMEModeHangul '한글
For i = 0 To 6
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Label1(i).Visible = False
Label1(i).Top = 0
Next
Label4.Caption = 0
Label5.Caption = 10
Timer1.Enabled = False
Label10.Caption = "1단계"
End Sub

Private Sub Form_Unload(Cancel As Integer)
mumain.Agent1.Characters("Genie").Stop

End Sub

Private Sub Label11_Click()
Randomize


Text1.IMEMode = vbIMEModeHangul '한글

Open App.Path & "\txt\타자글.txt" For Input As #1
For i = 1 To 200
    Input #1, 낱말(i)
Next
Close #1


For i = 0 To 6
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Label1(i).Visible = False
Label1(i).Top = 0
Next
Label4.Caption = 0
Label5.Caption = 10
Timer1.Enabled = False
Command5.Enabled = False
Command6.Enabled = True
Label10.Caption = "1단계"
End Sub

Private Sub Label12_Click()
Randomize


Text1.IMEMode = vbIMEModeAlpha  '영문

Open App.Path & "\txt\타자영문.txt" For Input As #1
For i = 1 To 100
    Input #1, 영문(i)
Next
Close #1

Label10.Caption = "1단계"
For i = 0 To 6
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Label1(i).Visible = False
Label1(i).Top = 0
Next
Label4.Caption = 0
Label5.Caption = 10
Timer1.Enabled = False
Command5.Enabled = True
Command6.Enabled = False
End Sub

Private Sub Label13_Click()
Randomize
Label4.Caption = 0
Label5.Caption = 10
For i = 0 To 6
If Command5.Enabled = False Then
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
ElseIf Command6.Enabled = False Then
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Label1(i).Top = 0
End If
Next


For i = 0 To 6
Label1(i).Visible = True
Next
Timer1.Enabled = True
Text1.SetFocus


If Command5.Enabled = True Then
If Command6.Enabled = True Then
MsgBox "타자 종류를 선택해 주세요." + Chr(13) + Chr(13) _
       + "맨 밑에 한글과 영어중 선택해 주세요", vbCritical




End If
End If
End Sub

Private Sub Label6_Click()
mumain.Show
Unload Me
mumain.Agent1.Characters("Genie").Stop
End Sub

Private Sub Label9_Click()
     ReDim PLAYER(1 To 11) As String
     ReDim SCORE(1 To 11) As Integer
     Dim PLAYERNAME As String
     Dim PLAYERTEMP As String
     Dim SCORETEMP As Integer
     Dim MAINLOOP As Integer
     Dim LOOPCTR As Integer
     Dim FOUNDSW As Integer
     Open App.Path & "\xkwk.txt" For Input As #1
     For LOOPCTR = 1 To 10
          Input #1, PLAYER(LOOPCTR), SCORE(LOOPCTR)
     Next LOOPCTR
     Close #1
     For LOOPCTR = 1 To 10
          If Val(Label4.Caption) > SCORE(LOOPCTR) Then FOUNDSW = 1
     Next LOOPCTR
     If FOUNDSW = 1 Then
          PLAYERNAME = InputBox$("이름을 입력하세요")
          If PLAYERNAME = "" Then PLAYERNAME = "이름없음"
          PLAYER(11) = PLAYERNAME
          SCORE(11) = Val(Label4.Caption)
     End If
     For MAINLOOPCTR = 1 To 10
          For LOOPCTR = 1 To 10
               If SCORE(LOOPCTR) < SCORE(LOOPCTR + 1) Then
                    PLAYERTEMP = PLAYER(LOOPCTR)
                    PLAYER(LOOPCTR) = PLAYER(LOOPCTR + 1)
                    PLAYER(LOOPCTR + 1) = PLAYERTEMP
                    SCORETEMP = SCORE(LOOPCTR)
                    SCORE(LOOPCTR) = SCORE(LOOPCTR + 1)
                    SCORE(LOOPCTR + 1) = SCORETEMP
               End If
          Next LOOPCTR
     Next MAINLOOPCTR
     Open "xkwk.txt" For Output As #1
     MESSAGE = "높은 점수" + Chr$(13)
     For LOOPCTR = 1 To 10
          MESSAGE = MESSAGE + Chr$(13)
          If SCORE(LOOPCTR) = SCORECTR Then
               MESSAGE = MESSAGE + "-> " + PLAYER(LOOPCTR) + " - " + Format$(SCORE(LOOPCTR), "00000")
          Else
               MESSAGE = MESSAGE + PLAYER(LOOPCTR) + " - " + Format$(SCORE(LOOPCTR), "00000")
          End If
          Write #1, PLAYER(LOOPCTR), SCORE(LOOPCTR)
     Next LOOPCTR
     Close #1
     MsgBox (MESSAGE)


Timer1.Enabled = False
'cnt = cnt + 1
'맞춘갯수(cnt) = Label4.Caption
'이름(cnt) = Text2.Text
'점수(cnt) = 맞춘갯수(cnt) * 10
'If Command5.Enabled = True Then
'종류(cnt) = "영문"
'ElseIf Command6.Enabled = True Then
'종류(cnt) = "한글"
'End If
'If Option1 = True Then
'난이도(cnt) = "상급자"
'ElseIf Option2 = True Then
'난이도(cnt) = "중급자"
'ElseIf Option3 = True Then
'난이도(cnt) = "초급자"
'Else
'난이도(cnt) = "보통"
'End If
'Form2.Show
'Form2.Print "순위    이름    맞춘갯수     점수    난이도     종류"

'For i = 1 To cnt
'순위(i) = 1
'Next

'For i = 1 To cnt
'For J = 1 To cnt
'If 점수(i) > 점수(J) Then
'    순위(J) = 순위(J) + 1
'End If
'Next
'Next


'For i = 1 To cnt - 1
'For J = i + 1 To cnt
'If 맞춘갯수(i) < 맞춘갯수(J) Then
'   im = 순위(i)
'   순위(i) = 순위(J)
'   순위(J) = im
   
 '  im = 이름(i)
 '  이름(i) = 이름(J)
 '  이름(J) = im
 '
 '  im = 맞춘갯수(i)
 '  맞춘갯수(i) = 맞춘갯수(J)
 '  맞춘갯수(J) = im
 '
 '  im = 점수(i)
 '  점수(i) = 점수(J)
 '  점수(J) = im
 '
 '  im = 난이도(i)
 '  난이도(i) = 난이도(J)
 '  난이도(J) = im
 '
 '  im = 종류(i)
 '  종류(i) = 종류(J)
 '  종류(J) = im
 '  End If
 '  Next
 '  Next

'For i = 1 To cnt
'Form2.Print Tab(2); 순위(i);
'Form2.Print Tab(7); 이름(i);
'Form2.Print Tab(16); 맞춘갯수(i);
'Form2.Print Tab(26); 점수(i);
'Form2.Print Tab(32); 난이도(i);
'Form2.Print Tab(42); 종류(i)
'Next


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Randomize
Label7.Caption = Text1.Text
If KeyAscii = 13 Or KeyAscii = 32 Then
For i = 0 To 6

        If Label1(i).Caption = Trim(Text1.Text) Then
           Label4.Caption = Label4.Caption + 1
           Label1(i).Top = 0

       If Command5.Enabled = False Then
      
           a = Int(Rnd(1) * 200)
           Label1(i).Caption = 낱말(a)
           If Label1(i).Caption = "" Then
           a = Int(Rnd(1) * 200)
           Label1(i).Caption = 낱말(a)
           End If
           ElseIf Command6.Enabled = False Then
             a = Int(Rnd(1) * 100)
           Label1(i).Caption = 영문(a)
          
End If
           
            End If
              Next
    Text1.Text = ""
    Text1.SetFocus
If Command5.Enabled = False Then
If Label4.Caption = 30 Then
Timer1.Enabled = False
MsgBox "2단계"
Label10.Caption = "2단계"
For i = 0 To 6
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Label1(i).Top = 0
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 65 Then
Timer1.Enabled = False
MsgBox "3단계"
Label10.Caption = "3단계"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 100 Then
Timer1.Enabled = False
MsgBox "4단계"
Label10.Caption = "4단계"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 140 Then
Timer1.Enabled = False
MsgBox "5단계"
Label10.Caption = "5단계"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 195 Then
Timer1.Enabled = False
MsgBox "6단계"
Label10.Caption = "6단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 250 Then
Timer1.Enabled = False
MsgBox "7단계"
Label10.Caption = "7단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 300 Then
Timer1.Enabled = False
MsgBox "8단계"
Label10.Caption = "8단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 360 Then
Timer1.Enabled = False
MsgBox "9단계"
Label10.Caption = "9단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 425 Then
Timer1.Enabled = False
MsgBox "10단계"
Label10.Caption = "10단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 500 Then
Timer1.Enabled = False
MsgBox "마지막 단계"
Label10.Caption = "마지막 단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 200)
Label1(i).Caption = 낱말(a)
Next
Timer1.Enabled = True
End If
End If



If Command6.Enabled = False Then
If Label4.Caption = 30 Then
Timer1.Enabled = False
MsgBox "2단계"
Label10.Caption = "2단계"
For i = 0 To 6
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Label1(i).Top = 0
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 65 Then
Timer1.Enabled = False
MsgBox "3단계"
Label10.Caption = "3단계"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 100 Then
Timer1.Enabled = False
MsgBox "4단계"
Label10.Caption = "4단계"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 140 Then
Timer1.Enabled = False
MsgBox "5단계"
Label10.Caption = "5단계"
For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 195 Then
Timer1.Enabled = False
MsgBox "6단계"
Label10.Caption = "6단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 250 Then
Timer1.Enabled = False
MsgBox "7단계"
Label10.Caption = "7단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 300 Then
Timer1.Enabled = False
MsgBox "8단계"
Label10.Caption = "8단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 360 Then
Timer1.Enabled = False
MsgBox "9단계"
Label10.Caption = "9단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 425 Then
Timer1.Enabled = False
MsgBox "10단계"
Label10.Caption = "10단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
ElseIf Label4.Caption = 500 Then
Timer1.Enabled = False
MsgBox "마지막 단계"
Label10.Caption = "마지막 단계"

For i = 0 To 6
Label1(i).Top = 0
a = Int(Rnd(1) * 100)
Label1(i).Caption = 영문(a)
Next
Timer1.Enabled = True
End If
End If

End If
End Sub

Private Sub Timer1_Timer()
Randomize

Timer1.Enabled = True


For i = 0 To 6
Label1(i).Visible = True
If Option1 = True Then
If Label10.Caption = "1단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(1) * 15)
ElseIf Label10.Caption = "2단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(5) * 20)
ElseIf Label10.Caption = "3단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(10) * 25)
ElseIf Label10.Caption = "4단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 25)
ElseIf Label10.Caption = "5단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(20) * 30)
ElseIf Label10.Caption = "6단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(20) * 30)
ElseIf Label10.Caption = "7단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(20) * 35)
ElseIf Label10.Caption = "8단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(25) * 40)
ElseIf Label10.Caption = "9단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(25) * 45)
ElseIf Label10.Caption = "10단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(30) * 45)
ElseIf Label10.Caption = "마지막 단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(35) * 50)
End If
ElseIf Option2 = True Then
If Label10.Caption = "1단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(1) * 10)
ElseIf Label10.Caption = "2단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(5) * 10)
ElseIf Label10.Caption = "3단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(10) * 15)
ElseIf Label10.Caption = "4단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(10) * 20)
ElseIf Label10.Caption = "5단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 20)
ElseIf Label10.Caption = "6단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(20) * 25)
ElseIf Label10.Caption = "7단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(20) * 30)
ElseIf Label10.Caption = "8단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(25) * 30)
ElseIf Label10.Caption = "9단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(25) * 35)
ElseIf Label10.Caption = "10단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(30) * 35)
ElseIf Label10.Caption = "마지막 단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(35) * 45)
End If
ElseIf Option3 = True Then
If Label10.Caption = "1단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(1) * 5)
ElseIf Label10.Caption = "2단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(5) * 10)
ElseIf Label10.Caption = "3단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(8) * 12)
ElseIf Label10.Caption = "4단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(10) * 15)
ElseIf Label10.Caption = "5단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(10) * 20)
ElseIf Label10.Caption = "6단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 20)
ElseIf Label10.Caption = "7단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(15) * 25)
ElseIf Label10.Caption = "8단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(20) * 25)
ElseIf Label10.Caption = "9단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(25) * 28)
ElseIf Label10.Caption = "10단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(30) * 32)
ElseIf Label10.Caption = "마지막 단계" Then
Label1(i).Top = Label1(i).Top + Int(Rnd(30) * 36)
End If
End If
Next

For i = 0 To 6
        
         If Label1(i).Top > 3550 Then
            Label5.Caption = Label5.Caption - 1
            Label1(i).Top = 0
            If Command5.Enabled = False Then
            a = Int(Rnd(1) * 200)
            Label1(i).Caption = 낱말(a)
            ElseIf Command6.Enabled = False Then
            a = Int(Rnd(1) * 100)
            Label1(i).Caption = 영문(a)
            End If
            ElseIf Label5.Caption = 0 Then
            
            
            Timer1.Enabled = False
            Label1(i).Visible = False
            MsgBox "Game over"
            Exit Sub
            ElseIf Label5.Caption < 0 Then
            Label5.Caption = 0
            
                                    
End If
        
            
Next

If Command5.Enabled = False Then
If Label1(0).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
End If
If Label1(0).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
End If
If Label1(0).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
End If
If Label1(0).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
End If
If Label1(0).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
End If
If Label1(0).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
End If

If Label1(1).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(1).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
End If
If Label1(1).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
End If
If Label1(1).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
End If
If Label1(1).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
End If
If Label1(1).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
End If

If Label1(2).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(2).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
End If
If Label1(2).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
End If
If Label1(2).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
End If
If Label1(2).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
End If
If Label1(2).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
End If

If Label1(3).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(3).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
End If
If Label1(3).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
End If
If Label1(3).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
End If
If Label1(3).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
End If
If Label1(3).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
End If

If Label1(4).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(4).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
End If
If Label1(4).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
End If
If Label1(4).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
End If
If Label1(4).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
End If
If Label1(4).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
End If

If Label1(5).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(5).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(1).Caption = 낱말(a)
End If
If Label1(5).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(2).Caption = 낱말(a)
End If
If Label1(5).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(3).Caption = 낱말(a)
End If
If Label1(5).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(4).Caption = 낱말(a)
End If
If Label1(5).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 200)
Label1(5).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
End If

If Label1(6).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(6).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(6).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(6).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(6).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
If Label1(6).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 200)
Label1(6).Caption = 낱말(a)
a = Int(Rnd(1) * 200)
Label1(0).Caption = 낱말(a)
End If
End If

If Command6.Enabled = False Then
If Label1(0).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
End If
If Label1(0).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
End If
If Label1(0).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
End If
If Label1(0).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
End If
If Label1(0).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
End If
If Label1(0).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
End If

If Label1(1).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(1).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
End If
If Label1(1).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
End If
If Label1(1).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
End If
If Label1(1).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
End If
If Label1(1).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
End If

If Label1(2).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(2).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
End If
If Label1(2).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
End If
If Label1(2).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
End If
If Label1(2).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
End If
If Label1(2).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
End If

If Label1(3).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(3).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
End If
If Label1(3).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
End If
If Label1(3).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
End If
If Label1(3).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
End If
If Label1(3).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
End If

If Label1(4).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(4).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
End If
If Label1(4).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
End If
If Label1(4).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
End If
If Label1(4).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
End If
If Label1(4).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
End If

If Label1(5).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(5).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(1).Caption = 영문(a)
End If
If Label1(5).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(2).Caption = 영문(a)
End If
If Label1(5).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(3).Caption = 영문(a)
End If
If Label1(5).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(4).Caption = 영문(a)
End If
If Label1(5).Caption = Label1(6).Caption Then
a = Int(Rnd(1) * 100)
Label1(5).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
End If

If Label1(6).Caption = Label1(0).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(6).Caption = Label1(1).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(6).Caption = Label1(2).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(6).Caption = Label1(3).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(6).Caption = Label1(4).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
If Label1(6).Caption = Label1(5).Caption Then
a = Int(Rnd(1) * 100)
Label1(6).Caption = 영문(a)
a = Int(Rnd(1) * 100)
Label1(0).Caption = 영문(a)
End If
End If


End Sub

Private Sub Timer2_Timer()
If Timer1.Enabled = True Then
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Label13.Visible = False
End If
If Timer2.Enabled = False Then
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Label13.Visible = True
End If


End Sub
