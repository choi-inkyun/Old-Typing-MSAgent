VERSION 5.00
Begin VB.Form mumain 
   BackColor       =   &H80000009&
   Caption         =   "타자 프로그램"
   ClientHeight    =   3765
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5085
   DrawStyle       =   3  '대시-점
   LinkTopic       =   "Form2"
   ScaleHeight     =   3765
   ScaleWidth      =   5085
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture1 
      Height          =   8535
      Left            =   0
      Picture         =   "타자연습 V1.0.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   0
      Width           =   10335
   End
   Begin VB.Menu mufile 
      Caption         =   "파일(&f)"
      Begin VB.Menu muend 
         Caption         =   "끝내기"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mugame 
      Caption         =   "게임(&g)"
      Begin VB.Menu mu111 
         Caption         =   "이슬비"
         Shortcut        =   ^A
      End
      Begin VB.Menu musdsd 
         Caption         =   "제작중"
         Shortcut        =   ^D
      End
      Begin VB.Menu mueksdj 
         Caption         =   "단어잡기"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu muxkwk 
      Caption         =   "타자(&s)"
      Begin VB.Menu murmf 
         Caption         =   "짧은글 연습"
         Shortcut        =   ^I
      End
      Begin VB.Menu mulong 
         Caption         =   "긴글 연습"
         Shortcut        =   ^L
      End
      Begin VB.Menu murjawjd 
         Caption         =   "타자 검정"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu muhelp 
      Caption         =   "도움말(&h)"
      Begin VB.Menu mu111help 
         Caption         =   "이슬비 도움말"
         Shortcut        =   ^H
      End
      Begin VB.Menu mume 
         Caption         =   "제작자"
      End
      Begin VB.Menu muhelp1 
         Caption         =   "도움"
      End
      Begin VB.Menu muver 
         Caption         =   "버전"
      End
   End
End
Attribute VB_Name = "mumain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub mu111_Click()
Form1.Show
End Sub

Private Sub mu111help_Click()
MsgBox "< 이슬비 도움말 >" + Chr(13) + Chr(13) _
       + "이슬비도 제가 자작한 프로그램 입니다." + Chr(13) + Chr(13) _
       + "처음에 시작할때 맨 밑에있는 한글과 영문중에서 선택해 주세요." + Chr(13) + Chr(13) _
       + "난이도 조절도 해주세요. 그리고 프로그램을 끄면 점수가 지워집니다." + Chr(13) + Chr(13) _
       + "점수계산을 포함해 문제가 많군요. 앞으로 계속 버전업할 생각입니다." + Chr(13) + Chr(13) _
       + "맞춘갯수가 10씩 올라갈때마다 속도가 빨라집니다." + Chr(13) + Chr(13) _
       + "즐겁게 하세요..^^" + Chr(13) + Chr(13)
       
       

End Sub

Private Sub mueksdj_Click()
Form8.Show
End Sub

Private Sub muend_Click()
    a = MsgBox("종료하시겠습니까?", vbCritical + vbYesNo, "종료")
    
    If a = vbYes Then
        End
    Else
        Exit Sub
    End If

End Sub

Private Sub muhelp1_Click()
MsgBox "이 프로그램은 누구나 쉽게 이용할수 있게" + Chr(13) + Chr(13) _
       + "만들었습니다.^^ 그래서 도움말은 생략하겠습니다." + Chr(13) + Chr(13) _
       + "궁금하신점이 있으시면 메일로 보내주세요."
End Sub

Private Sub mulong_Click()
Form4.Show
End Sub

Private Sub mume_Click()
MsgBox "안녕하세요..^^" + Chr(13) + Chr(13) _
                  + "이 프로그램의 제작자인 최인균 이라고 합니다." + Chr(13) + Chr(13) _
                  + "비록 잘 못짠 프로그램 일지라도 잘 봐주세요." + Chr(13) + Chr(13) _
                  + "제 소개를 할께요." + Chr(13) + Chr(13) _
                  + "86년 2월22일 태어났고요, 소년입니다." + Chr(13) + Chr(13) _
                  + "그리고 현재 능곡중학교 3학년에 재학중입니다" + Chr(13) + Chr(13) _
                  + "이메일 = dingpong@hitel.net" + Chr(13) + Chr(13) _
                  + "그럼 재미있게 즐기세요..^^" + Chr(13) + Chr(13)
End Sub

Private Sub murjawjd_Click()
Form5.Show
End Sub

Private Sub murmf_Click()
Form3.Show
End Sub

Private Sub musdsd_Click()
Form7.Show
End Sub

Private Sub muver_Click()
muloag.Show
End Sub

