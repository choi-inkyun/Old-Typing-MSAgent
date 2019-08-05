VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   8865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
   Icon            =   "시작.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   5040
      Top             =   3000
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   187
      Left            =   3820
      TabIndex        =   5
      Top             =   8280
      Width           =   187
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   187
      Left            =   3820
      TabIndex        =   4
      Top             =   8520
      Width           =   187
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4080
      Top             =   3000
   End
   Begin VB.Timer 순차 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3600
      Top             =   3000
   End
   Begin VB.Timer 올라가라 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   4560
      Top             =   3000
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   230
      Left            =   6000
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "15"
      ToolTipText     =   "대사가 사라질때의 속도를 입력합니다"
      Top             =   9000
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   230
      Left            =   4800
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "15"
      ToolTipText     =   "대사가 생겨날때의 속도를 입력합니다"
      Top             =   8880
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   1440
      MaxLength       =   28
      TabIndex        =   0
      Text            =   "프로그램 공모전 최인균 자작 프로그램"
      ToolTipText     =   "대사를 입력합니다"
      Top             =   8880
      Width           =   2535
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   11
      Top             =   9000
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   360
      MousePointer    =   2  '십자형
      TabIndex        =   10
      Top             =   7600
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   240
      MousePointer    =   2  '십자형
      TabIndex        =   9
      Top             =   7080
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   360
      MousePointer    =   2  '십자형
      TabIndex        =   8
      Top             =   6300
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   360
      MousePointer    =   2  '십자형
      TabIndex        =   7
      Top             =   5800
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   360
      MousePointer    =   2  '십자형
      TabIndex        =   6
      Top             =   5290
      Width           =   3855
   End
   Begin VB.Image Image4 
      Height          =   4350
      Left            =   45
      Picture         =   "시작.frx":030A
      Top             =   4500
      Width           =   4260
   End
   Begin VB.Image Image3 
      Height          =   750
      Left            =   4320
      MousePointer    =   2  '십자형
      Picture         =   "시작.frx":3C876
      Top             =   1200
      Width           =   3000
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   3720
      MousePointer    =   2  '십자형
      Picture         =   "시작.frx":43DEA
      Top             =   360
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   8850
      Left            =   0
      Picture         =   "시작.frx":4B35E
      Top             =   0
      Width           =   7500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   840
      X2              =   6240
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Height          =   180
      Left            =   3960
      TabIndex        =   3
      Top             =   7800
      Width           =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aaa As Integer
Dim 숫자(100)
Dim 사라져 As Boolean
Dim 현재개수
Dim 대사
Dim aa, bb, cc, dd, ee, ff, gg, hh, jj
Dim a, b, c, d, e
Private Sub 모두없어져()
Dim I As Integer
For I = 1 To 현재개수
Unload Label1(I)
Next I
Timer1.Enabled = True
End Sub
Private Sub 순차_Timer()
Static a As Integer
Dim 표시
If a = Len(Text1.Text) Then 순차.Enabled = False: a = 0: GoTo 1
a = a + 1
표시 = Mid(대사, a, 1)
Load Label1(a)
Label1(a).Move Label1(a - 1).Left + Label1(a - 1).Width, Label1(a - 1).Top
Label1(a).ForeColor = RGB(0, 0, 0)
Label1(a).Visible = True
Label1(a).Caption = 표시
Label1(a).ForeColor = RGB(0, 0, 0)
올라가라.Enabled = True
현재개수 = a
If a = 1 And b = 1 And c = 1 And d = 1 And e = 1 Then
    Check1.Value = 1
End If
1 End Sub

Private Sub 올라가라_Timer()
Static 완료개수
Static b As Integer
Dim I As Integer
If 사라져 = True Then GoTo 2 Else GoTo 1
1 For I = 0 To 현재개수
숫자(I) = 숫자(I) + Val(Text2.Text)
Label1(I).ForeColor = RGB(숫자(I), 숫자(I), 숫자(I))
If 숫자(I) >= 255 Then 숫자(I) = 255: 완료개수 = I
Next I
GoTo 3
2 For I = 0 To 현재개수
숫자(I) = 숫자(I) - Val(Text3.Text)
If 숫자(I) <= 0 Then 숫자(I) = 0: 완료개수 = I
Label1(I).ForeColor = RGB(숫자(I), 숫자(I), 숫자(I))
Next I
GoTo 4
3 If 완료개수 = Len(대사) Then 사라져 = True: 완료개수 = 0: GoTo 5
4 If 완료개수 = Len(대사) Then 올라가라.Enabled = False: 완료개수 = 0: Text1.Enabled = True: Text2.Enabled = True: Text3.Enabled = True: 모두없어져: 사라져 = False: 현재개수 = 0: GoTo 5
5 End Sub




Private Sub asAssisPopup8_Click()
End
End Sub

Private Sub Check1_Click()
     Open App.Path & "\체크.txt" For Input As #1
          Input #1, aaa
     Close #1
Open App.Path & "\체크.txt" For Output As #1
          If Check1.Value = 1 Then
          aaa = 1
          Write #1, aaa
          End If
          If Check1.Value = 0 Then
          aaa = 0
          Write #1, aaa
          End If
          
    Close #1
End Sub

Private Sub Form_Load()
     Open App.Path & "\체크.txt" For Input As #1
          Input #1, aaa
          If aaa = 1 Then
        Call Shell(App.Path & "\시간관리.exe", 1)
          End
          End If
     Close #1
     aa = 0
     bb = 0
     cc = 0
     dd = 0
     ee = 0
     ff = 0
     gg = 0
     hh = 0
     jj = 0
     a = 0
     b = 0
     c = 0
     d = 0
     e = 0
 Form2.Show
' Label2.Caption = "안녕하세요.^^.제가 만든 자작 프로그램의 설치 옵션입니다." + Chr(13) _
'                + "설치 해야할께 많죠?.필수 항목에 있는것들은 꼭 설치해 주시고. 밑에 선택항목에" + Chr(13) _
'                + "있는것은 설치를 하지 않으셔도 실행하시고 이용하실 수 있지만." + Chr(13) _
'                + "핵심인 음성출력부분이 않돼기 때문에 되도록이면꼭설치해주셨으면하는 바램입니다" + Chr(13) _
'                + "밑에체크하는것을 체크하시고다시 실행하시면 이 화면이다시는 않뜨는것을 볼수있습니다. 다시 뜨게 하고싶으시면 체크.txt 를 0으로 바꿔주세요.그럼유용하게쓰시길~"
Timer2.Enabled = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If jj = 0 Then
jj = 1
aa = 0
bb = 0
cc = 0
dd = 0
ee = 0
ff = 0
gg = 0
hh = 0
End If
Timer2.Enabled = True
End Sub

Private Sub Image2_Click()
Dim a As Integer
If Check2.Value = 1 Then
Kill App.Path & "\util\MSAgent.exe"
Kill App.Path & "\util\AgtX0412.exe"
Kill App.Path & "\util\Genie.exe"
Kill App.Path & "\util\lhttskok.exe"
Kill App.Path & "\util\spchapi.exe"
End If
a = MsgBox("설치는 모두 하셨습니까?", vbYesNo, "설치")
If a = vbYes Then
Check1.Value = 1
     Open App.Path & "\체크.txt" For Input As #1
          Input #1, aaa
     Close #1
     Open App.Path & "\체크.txt" For Output As #1
          If Check1.Value = 1 Then
          aaa = 1
          Write #1, aaa
          End If
          If Check1.Value = 0 Then
          aaa = 0
          Write #1, aaa
          End If
          
    Close #1
    Call Shell(App.Path & "\시간관리.exe", 1)
Unload Form2
End
End If



End Sub

Private Sub Image3_Click()
    a = MsgBox("종료하시겠습니까?", vbCritical + vbYesNo, "종료")
    
    If a = vbYes Then
    End
    End If
End Sub

Private Sub Label2_Click()
Call Shell(App.Path & "\util\MSAgent.exe", 1)
a = 1
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If aa = 0 Then
jj = 0
ee = 0
aa = 1
bb = 0
jj = 0
cc = 0
dd = 0
End If
Timer2.Enabled = True

End Sub

Private Sub Label4_Click()
Call Shell(App.Path & "\util\AgtX0412.exe", 1)
b = 1
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer2.Enabled = True
If bb = 0 Then
ee = 0
jj = 0
aa = 0
bb = 1
cc = 0
dd = 0
End If
End Sub

Private Sub Label5_Click()
Call Shell(App.Path & "\util\Genie.exe", 1)
c = 1
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer2.Enabled = True
If cc = 0 Then
ee = 0
jj = 0
aa = 0
bb = 0
cc = 1
dd = 0
End If
End Sub

Private Sub Label6_Click()
Call Shell(App.Path & "\util\lhttskok.exe", 1)
d = 1
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If dd = 0 Then
dd = 1
ee = 0
aa = 0
bb = 0
jj = 0
cc = 0
End If
Timer2.Enabled = True

End Sub

Private Sub Label7_Click()
Call Shell(App.Path & "\util\spchapi.exe", 1)
e = 1
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ee = 0 Then
ee = 1
aa = 0
bb = 0
cc = 0
jj = 0
dd = 0
End If
Timer2.Enabled = True

End Sub

Private Sub Timer1_Timer()
  If Len(Text2.Text) = 0 Then MsgBox "속도 1의 값을 넣어주세요.", vbExclamation, "에러": GoTo 2
2 If Len(Text3.Text) = 0 Then MsgBox "속도 2의 값을 넣어주세요.", vbExclamation, "에러": GoTo 3

 대사 = Text1.Text
 순차.Enabled = True
 Text1.Enabled = False
 Text2.Enabled = False
 Text3.Enabled = False
Timer1.Enabled = False

3 End Sub

Private Sub Timer2_Timer()
If aa = 1 Then
Image4.Picture = LoadPicture(App.Path & "\image\a1.bmp")
aa = 3
End If
If bb = 1 Then
Image4.Picture = LoadPicture(App.Path & "\image\a2.bmp")
bb = 3
End If
If cc = 1 Then
Image4.Picture = LoadPicture(App.Path & "\image\a3.bmp")
cc = 3
End If
If dd = 1 Then
Image4.Picture = LoadPicture(App.Path & "\image\a4.bmp")
dd = 3
End If
If ee = 1 Then
Image4.Picture = LoadPicture(App.Path & "\image\a5.bmp")
ee = 3
End If
If aa = 0 And bb = 0 And cc = 0 And dd = 0 And ee = 0 And jj = 1 Then
Image4.Picture = LoadPicture(App.Path & "\image\2.bmp")
jj = 3
End If

Timer2.Enabled = False



End Sub
